/**
 * Copyright 2026 Mustafa Tunca (mustafatunca.com)
 * SPDX-License-Identifier: Apache-2.0
 *
 * outlook-followup-addin — Smart Follow-up Meeting Creator for Microsoft Outlook
 * https://github.com/mustafa-tunca/outlook-followup-addin
 */
/* global Office, document, window, console */

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/**
 * Augments AppointmentRead with properties that exist at runtime (Mailbox 1.5+)
 * but may be missing from the installed @types/office-js version.
 */
interface AppointmentReadEx extends Office.AppointmentRead {
  onlineMeetingUrl?: string | null;
  getAttachmentsAsync(
    callback: (result: Office.AsyncResult<Office.AttachmentDetails[]>) => void
  ): void;
}

interface AttachmentInfo {
  id: string;
  name: string;
  size: number;
  isInline: boolean;
  selected: boolean;
}

interface MeetingContext {
  rawSubject: string;
  newSubject: string;
  start: Date;
  end: Date;
  /** UTC+7-day start — calculated via epoch ms, not local-time arithmetic. */
  newStart: Date;
  /** UTC+7-day end — calculated via epoch ms, not local-time arithmetic. */
  newEnd: Date;
  location: string;
  isOnlineMeeting: boolean;
  isRecurring: boolean;
  attachments: AttachmentInfo[];
  organizer: string;
  requiredAttendees: string[];
  optionalAttendees: string[];
  originalBody: string;
}

// ---------------------------------------------------------------------------
// Entry point
// ---------------------------------------------------------------------------

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    loadMeetingData();
  }
});

// ---------------------------------------------------------------------------
// Subject logic
// ---------------------------------------------------------------------------

/**
 * Strips common reply/forward prefixes in any order/repetition.
 * Handles: RE:, FW:, FWD:, Antw:, WG: (German Outlook equivalents)
 */
function cleanSubject(subject: string): string {
  const prefixRe = /^\s*(RE|FW|FWD|Antw|WG)\s*:\s*/gi;
  let cleaned = subject;
  let prev: string;
  do {
    prev = cleaned;
    cleaned = cleaned.replace(prefixRe, "");
  } while (cleaned !== prev);
  return cleaned.trim();
}

/**
 * Appends or increments the "- Follow-up #N" suffix.
 */
function buildFollowUpSubject(cleaned: string): string {
  const followUpRe = /\s*-\s*Follow-up\s*#(\d+)\s*$/i;
  const match = cleaned.match(followUpRe);
  if (match) {
    const n = parseInt(match[1], 10);
    return cleaned.replace(followUpRe, ` - Follow-up #${n + 1}`);
  }
  return `${cleaned} - Follow-up #1`;
}

// ---------------------------------------------------------------------------
// UTC-safe date math
// ---------------------------------------------------------------------------

/** Adds exactly 7 × 24 × 60 × 60 × 1000 ms to a Date, working in UTC epoch. */
function addSevenDaysUTC(date: Date): Date {
  return new Date(date.getTime() + 604_800_000);
}

// ---------------------------------------------------------------------------
// Online-meeting detection
// ---------------------------------------------------------------------------

const ONLINE_KEYWORDS = [
  "teams.microsoft.com",
  "zoom.us",
  "webex.com",
  "meet.google.com",
  "gotomeeting.com",
  "microsoft teams",
  "teams meeting",
];

function detectOnlineMeeting(location: string, onlineMeetingUrl: string | null | undefined): boolean {
  if (onlineMeetingUrl) return true;
  const lower = location.toLowerCase();
  return ONLINE_KEYWORDS.some((kw) => lower.includes(kw));
}

// ---------------------------------------------------------------------------
// Data loading
// ---------------------------------------------------------------------------

let ctx: MeetingContext | null = null;

/** Fetches the item body as HTML (works in both read and compose mode). */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function fetchBodyHtml(item: any): Promise<string> {
  return new Promise((resolve) => {
    if (item.body && typeof item.body.getAsync === "function") {
      item.body.getAsync(Office.CoercionType.Html, (result: Office.AsyncResult<string>) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? (result.value ?? "") : "");
      });
    } else {
      resolve("");
    }
  });
}

/**
 * Resolves an Office.js field that is either a plain value (read mode)
 * or an async getter object with .getAsync() (compose/organizer mode).
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function resolveField<T>(field: any, fallback: T): Promise<T> {
  if (field != null && typeof field.getAsync === "function") {
    return new Promise((resolve) => {
      field.getAsync((result: Office.AsyncResult<T>) => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : fallback);
      });
    });
  }
  return Promise.resolve(field != null ? (field as T) : fallback);
}

/** Shows an error and makes the content pane visible so the banner is seen. */
function fatalError(msg: string): void {
  showSection("loading", false);
  showSection("content", true);   // must be visible for error-banner to show
  showSection("success-card", false);
  showError(msg);
}

async function loadMeetingData(): Promise<void> {
  showSection("loading", true);
  showSection("content", false);
  showSection("success-card", false);
  showSection("error-banner", false);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const item = Office.context.mailbox.item as any;

  if (!item) {
    fatalError("No appointment item found. Open the add-in from inside an appointment.");
    return;
  }

  try {
    const [subjectRaw, startRaw, endRaw, locationRaw, requiredRaw, optionalRaw, originalBody] = await Promise.all([
      resolveField<string>(item.subject, ""),
      resolveField<Date | string>(item.start, new Date()),
      resolveField<Date | string>(item.end, new Date()),
      resolveField<string>(item.location, ""),
      resolveField<Office.EmailAddressDetails[]>(item.requiredAttendees, []),
      resolveField<Office.EmailAddressDetails[]>(item.optionalAttendees, []),
      fetchBodyHtml(item),
    ]);

    const rawSubject = typeof subjectRaw === "string" ? subjectRaw : String(subjectRaw ?? "");
    const cleaned = cleanSubject(rawSubject);
    const newSubject = buildFollowUpSubject(cleaned);

    const start = new Date(startRaw as string | Date);
    const end   = new Date(endRaw   as string | Date);
    const newStart = addSevenDaysUTC(start);
    const newEnd   = addSevenDaysUTC(end);

    const location = typeof locationRaw === "string" ? locationRaw : String(locationRaw ?? "");
    const isOnlineMeeting = detectOnlineMeeting(location, item.onlineMeetingUrl);
    const isRecurring = item.recurrence != null;

    // Compose mode has no item.organizer — fall back to the current user's profile.
    let organizer = "";
    if (item.organizer && item.organizer.displayName !== undefined) {
      organizer = `${String(item.organizer.displayName ?? "")} <${String(item.organizer.emailAddress ?? "")}>`;
    } else {
      const profile = Office.context.mailbox.userProfile;
      if (profile) {
        organizer = `${String(profile.displayName ?? "")} <${String(profile.emailAddress ?? "")}>`;
      }
    }

    const requiredAttendees = Array.isArray(requiredRaw)
      ? requiredRaw.map((a: Office.EmailAddressDetails) => String(a.emailAddress ?? "")).filter(Boolean)
      : [];
    const optionalAttendees = Array.isArray(optionalRaw)
      ? optionalRaw.map((a: Office.EmailAddressDetails) => String(a.emailAddress ?? "")).filter(Boolean)
      : [];

    const doRender = (attachments: AttachmentInfo[]) => {
      ctx = {
        rawSubject, newSubject, start, end, newStart, newEnd,
        location, isOnlineMeeting, isRecurring, attachments,
        organizer, requiredAttendees, optionalAttendees, originalBody,
      };
      renderPreflight(ctx);
      showSection("loading", false);
      showSection("content", true);
    };

    // getAttachmentsAsync requires Mailbox 1.8 and is read-mode only.
    // Add a 5-second timeout so a silent API failure never leaves the spinner stuck.
    if (typeof item.getAttachmentsAsync === "function") {
      let settled = false;
      const timer = setTimeout(() => {
        if (!settled) { settled = true; doRender([]); }
      }, 5000);

      item.getAttachmentsAsync((attResult: Office.AsyncResult<Office.AttachmentDetails[]>) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);

        const attachments: AttachmentInfo[] = [];
        if (attResult.status === Office.AsyncResultStatus.Succeeded && attResult.value) {
          for (const att of attResult.value) {
            if (!att.isInline) {
              attachments.push({ id: att.id, name: att.name, size: att.size, isInline: false, selected: true });
            }
          }
        }
        doRender(attachments);
      });
    } else {
      doRender([]);
    }
  } catch (err) {
    fatalError(err instanceof Error ? err.message : String(err));
  }
}

// ---------------------------------------------------------------------------
// UI rendering
// ---------------------------------------------------------------------------

function renderPreflight(meeting: MeetingContext): void {
  setText("orig-subject", meeting.rawSubject);
  setText("orig-date", formatDate(meeting.start));
  setInputValue("new-subject", meeting.newSubject);
  setText("new-date", formatDate(meeting.newStart));
  setText("utc-detail", `UTC: ${meeting.newStart.toUTCString()}`);
  setInputValue("date-picker", toDatetimeLocalValue(meeting.newStart));

  // Attachments
  if (meeting.attachments.length > 0) {
    showSection("attachment-section", true);
    renderAttachments(meeting.attachments);
  } else {
    showSection("attachment-section", false);
  }

  // Recurrence
  showSection("recurrence-section", meeting.isRecurring);

  // Location
  if (meeting.location) {
    showSection("location-section", true);
    setText("location-value", meeting.location);

    if (meeting.isOnlineMeeting) {
      showSection("online-meeting-row", true);
      showSection("room-meeting-row", false);
    } else {
      showSection("online-meeting-row", false);
      showSection("room-meeting-row", true);
    }
  } else {
    showSection("location-section", false);
  }
}

function renderAttachments(attachments: AttachmentInfo[]): void {
  const list = document.getElementById("attachment-list");
  if (!list) return;

  list.innerHTML = "";

  attachments.forEach((att, idx) => {
    const label = document.createElement("label");
    label.className = "attachment-item";
    label.htmlFor = `att-${idx}`;

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.id = `att-${idx}`;
    checkbox.dataset["index"] = String(idx);
    checkbox.checked = att.selected;
    checkbox.className = "att-checkbox";
    checkbox.addEventListener("change", () => {
      if (ctx) ctx.attachments[idx].selected = checkbox.checked;
    });

    const nameSpan = document.createElement("span");
    nameSpan.className = "att-name";
    nameSpan.textContent = att.name;

    const sizeSpan = document.createElement("span");
    sizeSpan.className = "att-size";
    sizeSpan.textContent = formatBytes(att.size);

    label.appendChild(checkbox);
    label.appendChild(nameSpan);
    label.appendChild(sizeSpan);
    list.appendChild(label);
  });
}

// ---------------------------------------------------------------------------
// Meeting creation
// ---------------------------------------------------------------------------

async function createFollowUp(): Promise<void> {
  if (!ctx) return;

  const btn = document.getElementById("create-btn") as HTMLButtonElement | null;
  if (btn) { btn.disabled = true; btn.textContent = "Creating…"; }

  try {
    // Resolve final subject (user may have edited the input)
    const subjectInput = document.getElementById("new-subject") as HTMLInputElement | null;
    const finalSubject = (subjectInput?.value ?? "").trim() || ctx.newSubject;

    // Resolve date: auto (+7 days) or custom picker
    const isCustomDate = (document.getElementById("date-custom") as HTMLInputElement | null)?.checked ?? false;
    let finalStart: Date;
    let finalEnd: Date;
    if (isCustomDate) {
      const pickerEl = document.getElementById("date-picker") as HTMLInputElement | null;
      const pickerVal = pickerEl?.value;
      if (!pickerVal) { showError("Please select a date."); if (btn) { btn.disabled = false; btn.textContent = "Create Follow-up"; } return; }
      finalStart = new Date(pickerVal);
      const duration = ctx.end.getTime() - ctx.start.getTime();
      finalEnd = new Date(finalStart.getTime() + duration);
    } else {
      finalStart = ctx.newStart;
      finalEnd = ctx.newEnd;
    }

    // Resolve location
    const finalLocation = resolveLocation(ctx);

    // Build the Meeting Heritage block injected at the top of the body
    const newBody = buildMeetingHeritage(ctx);

    const params: Office.AppointmentForm = {
      subject: finalSubject,
      start: finalStart,
      end: finalEnd,
      location: finalLocation,
      body: newBody,
    };

    if (ctx.requiredAttendees.length > 0) {
      params.requiredAttendees = ctx.requiredAttendees;
    }
    if (ctx.optionalAttendees.length > 0) {
      params.optionalAttendees = ctx.optionalAttendees;
    }

    await displayAppointment(params);
    showSection("content", false);
    showSection("success-card", true);
    const usedDeepLink = (window as unknown as Record<string, unknown>)["_usedDeepLink"] === true;
    setText(
      "success-msg",
      usedDeepLink
        ? "A new calendar tab has been opened with subject, date, and attendees pre-filled. " +
          "The Meeting Heritage block could not be transferred — paste it from the taskpane manually if needed."
        : "Your follow-up meeting draft has been opened. Add attachments manually if needed, then send."
    );
  } catch (err) {
    const msg = err instanceof Error ? err.message : "An unexpected error occurred.";
    showError(msg);
    if (btn) { btn.disabled = false; btn.textContent = "Create Follow-up"; }
  }
}

function resolveLocation(meeting: MeetingContext): string {
  if (!meeting.location) return "";

  if (meeting.isOnlineMeeting) {
    // User can opt to convert to in-person by unchecking "Keep as online meeting"
    const keepOnline = (document.getElementById("keep-online") as HTMLInputElement | null)?.checked ?? true;
    // When keeping online, return empty — Outlook/Teams auto-generates the link.
    return keepOnline ? "" : "";
  } else {
    // Physical room — user can choose to keep or convert to Teams
    const keepRoom = (document.getElementById("keep-room") as HTMLInputElement | null)?.checked ?? false;
    return keepRoom ? meeting.location : "";
  }
}

function displayAppointment(params: Office.AppointmentForm): Promise<void> {
  return new Promise((resolve, reject) => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const mailbox = Office.context.mailbox as any;

    if (typeof mailbox.displayNewAppointmentFormAsync === "function") {
      mailbox.displayNewAppointmentFormAsync(params, (result: Office.AsyncResult<void>) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error?.message ?? "displayNewAppointmentFormAsync failed"));
        }
      });
      return;
    }

    if (typeof mailbox.displayNewAppointmentForm === "function") {
      try {
        mailbox.displayNewAppointmentForm(params);
        resolve();
      } catch (e) {
        reject(e instanceof Error ? e : new Error(String(e)));
      }
      return;
    }

    // OWA deep-link fallback: derive base URL from ewsUrl and open a compose tab.
    // Note: body is not transferred via deep link (OWA limitation).
    try {
      const ewsUrl: string | undefined = mailbox.ewsUrl;
      let base = "https://outlook.office365.com";
      if (ewsUrl) {
        try { base = new URL(ewsUrl).origin; } catch { /* keep default */ }
      }

      const start = params.start as Date;
      const end   = params.end   as Date;
      const qs = new URLSearchParams({
        subject:  String(params.subject ?? ""),
        startdt:  start.toISOString(),
        enddt:    end.toISOString(),
      });
      if (params.location) qs.set("location", String(params.location));
      if (Array.isArray(params.requiredAttendees) && params.requiredAttendees.length > 0) {
        qs.set("to", (params.requiredAttendees as string[]).join(";"));
      }
      if (Array.isArray(params.optionalAttendees) && params.optionalAttendees.length > 0) {
        qs.set("cc", (params.optionalAttendees as string[]).join(";"));
      }

      window.open(`${base}/calendar/action/compose?${qs.toString()}`, "_blank");
      // Flag so the success card can mention body must be added manually.
      (window as unknown as Record<string, unknown>)["_usedDeepLink"] = true;
      resolve();
    } catch (e) {
      reject(e instanceof Error ? e : new Error(String(e)));
    }
  });
}

// ---------------------------------------------------------------------------
// Meeting Heritage body block
// ---------------------------------------------------------------------------

function buildMeetingHeritage(meeting: MeetingContext): string {
  const selectedNames = meeting.attachments
    .filter((a) => a.selected)
    .map((a) => `<li style="margin:2px 0">${escHtml(a.name)}</li>`)
    .join("");

  const attachHtml = selectedNames
    ? `<ul style="margin:4px 0 0 16px;padding:0">${selectedNames}</ul>`
    : `<span style="color:#888">None</span>`;

  const attachWarning = meeting.attachments.some((a) => a.selected)
    ? `<p style="margin:8px 0 0;padding:8px 10px;background:#fff3cd;border:1px solid #ffc107;
         border-radius:4px;font-size:12px;color:#856404;">
        ⚠️ Attachments listed above must be added manually — the Outlook Add-in API does not
        support programmatic attachment transfer. Open the original meeting and drag files
        to this draft.
       </p>`
    : "";

  const recurrenceNote = meeting.isRecurring
    ? (() => {
        const instanceOnly =
          (document.getElementById("this-instance") as HTMLInputElement | null)?.checked ?? true;
        return `<tr>
          <td style="padding:2px 12px 2px 0;color:#555;font-weight:600;white-space:nowrap">Recurrence scope:</td>
          <td>${instanceOnly ? "This instance only" : "Entire series"}</td>
        </tr>`;
      })()
    : "";

  const originalBodySection = meeting.originalBody
    ? `<hr style="border:none;border-top:1px solid #ddd;margin:14px 0"/>
       <div style="font-family:Calibri,Arial,sans-serif;font-size:13px;color:#222">
         ${meeting.originalBody}
       </div>`
    : "";

  return `<div style="font-family:Calibri,Arial,sans-serif;font-size:13px;border:1px solid #ddd;
           border-radius:4px;padding:12px 14px;margin-bottom:14px;background:#f5f5f5;color:#222">
  <strong style="font-size:14px;color:#0078d4">📋 Meeting Heritage</strong>
  <table style="border-collapse:collapse;margin-top:8px;width:100%">
    <tr>
      <td style="padding:2px 12px 2px 0;color:#555;font-weight:600;white-space:nowrap">Original subject:</td>
      <td>${escHtml(meeting.rawSubject)}</td>
    </tr>
    <tr>
      <td style="padding:2px 12px 2px 0;color:#555;font-weight:600;white-space:nowrap">Original date:</td>
      <td>${formatDate(meeting.start)} (UTC: ${meeting.start.toUTCString()})</td>
    </tr>
    <tr>
      <td style="padding:2px 12px 2px 0;color:#555;font-weight:600;white-space:nowrap">Organizer:</td>
      <td>${escHtml(meeting.organizer || "Unknown")}</td>
    </tr>
    ${recurrenceNote}
    <tr>
      <td style="padding:2px 12px 2px 0;color:#555;font-weight:600;vertical-align:top;white-space:nowrap">Carry-over attachments:</td>
      <td>${attachHtml}</td>
    </tr>
  </table>
  ${attachWarning}
</div>
${originalBodySection}
`;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Returns a value suitable for <input type="datetime-local"> (local time, no seconds). */
function toDatetimeLocalValue(date: Date): string {
  const pad = (n: number) => String(n).padStart(2, "0");
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}T${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function onDateModeChange(): void {
  const isCustom = (document.getElementById("date-custom") as HTMLInputElement | null)?.checked ?? false;
  showSection("date-auto-display", !isCustom);
  showSection("date-picker-wrap", isCustom);
}

function formatDate(date: Date): string {
  return date.toLocaleString(undefined, {
    weekday: "short",
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
    timeZoneName: "short",
  });
}

function formatBytes(bytes: number): string {
  if (bytes < 1_024) return `${bytes} B`;
  if (bytes < 1_048_576) return `${(bytes / 1_024).toFixed(1)} KB`;
  return `${(bytes / 1_048_576).toFixed(1)} MB`;
}

function escHtml(str: string): string {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function showSection(id: string, visible: boolean): void {
  const el = document.getElementById(id);
  if (el) el.style.display = visible ? "" : "none";
}

function setText(id: string, text: string): void {
  const el = document.getElementById(id);
  if (el) el.textContent = text;
}

function setInputValue(id: string, value: string): void {
  const el = document.getElementById(id) as HTMLInputElement | null;
  if (el) el.value = value;
}

function showError(message: string): void {
  showSection("error-banner", true);
  setText("error-text", message);
}

// Expose to HTML onclick attributes
(window as unknown as Record<string, unknown>)["createFollowUp"] = createFollowUp;
(window as unknown as Record<string, unknown>)["dismissError"] = () => showSection("error-banner", false);
(window as unknown as Record<string, unknown>)["onDateModeChange"] = onDateModeChange;
