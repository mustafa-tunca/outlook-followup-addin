# Follow-up Meeting Add-in for Outlook

A Microsoft Outlook add-in that creates smart follow-up appointments with one click.
Open an existing meeting, click **"Create Follow-up"**, review the pre-filled details, and confirm.

**What it does automatically:**
- Cleans the subject (strips `RE:`, `FW:`, `FWD:`, `Antw:`, `WG:` prefixes)
- Appends `– Follow-up #1` (or increments an existing counter)
- Schedules the new meeting exactly 7 days later (UTC-safe, immune to daylight saving time)
- Carries over attendees, location, and the original invitation body
- Lists attachments with checkboxes (must be added manually — Office.js limitation)
- Lets you pick a custom date instead of +7 days
- Works whether you are the **meeting organizer or an attendee**

---

## Requirements

Before you start, make sure the following are installed:

| Requirement | Where to get it | Notes |
|---|---|---|
| **Node.js 18 LTS or newer** | https://nodejs.org → "LTS" download | Includes `npm`. Restart your PC after install. |
| **Microsoft 365 account** | Your company account | Outlook Web App (OWA) access required |
| **A modern browser** | Edge, Chrome, Firefox | For OWA |

> **Corporate IT note:** The add-in runs entirely on your own PC (`https://localhost:3000`).
> No data is sent to any external server. The HTTPS certificate is self-signed and only trusted locally.

---

## First-time setup (new PC)

### Step 1 — Install Node.js

1. Go to **https://nodejs.org** and download the **LTS** version.
2. Run the installer, keep all defaults, click through to finish.
3. **Restart your PC** so the `npm` command becomes available system-wide.

### Step 2 — Get the project files

Copy the entire `outlook-followup-addin` folder to your PC.
You can put it anywhere — for example `C:\Tools\outlook-followup-addin\` or keep it on OneDrive.

> Do **not** move the folder after setup. The auto-start task remembers the exact path.
> If you move it later, re-run `Install.bat`.

### Step 3 — Run the installer

Double-click **`Install.bat`** inside the folder.

The installer will automatically:
1. Check that Node.js is available
2. Download all required packages (`npm install`)
3. Build the add-in (`npm run build`)
4. Install and trust the HTTPS development certificate
   *(A Windows UAC prompt will appear — click **Yes**)*
5. Register a background task that starts the server silently at every login

> After the installer finishes, it prints the OWA sideloading instructions. Keep the window open for reference.

---

## Sideloading into Outlook Web App (OWA)

This step is done **once per browser / per user account**.

1. Open **Outlook Web App** in your browser (e.g. `https://outlook.office.com`)
2. Click any **calendar appointment** to open it in read mode
3. In the appointment window, click **"..."** (More options) in the toolbar
4. Choose **"Get Add-ins"** (or **"Manage add-ins"**)
5. In the dialog that opens, click **"My add-ins"** in the left panel
6. Scroll to the bottom and click **"+ Add a custom add-in"** → **"Add from file..."**
7. Browse to your `outlook-followup-addin` folder and select **`manifest.json`**
8. Confirm any security warning
9. Close the dialog

The **"Create Follow-up"** button now appears in the appointment ribbon.

> **Tip:** If you don't see the button, close and reopen the appointment.

---

## Daily use

### The server starts automatically

After running `Install.bat` once, the dev server starts **silently in the background** every time you log into Windows. You don't need to do anything.

The server runs at `https://localhost:3000` and serves the add-in files to OWA.

### Using the add-in

1. Open any calendar appointment in OWA
   (works for meetings you organised **and** meetings you were invited to)
2. Click **"Create Follow-up"** in the ribbon
3. A side panel opens with pre-filled details:
   - **Subject** — editable, already cleaned and numbered
   - **Date** — toggle between `+7 days` (automatic) or `Pick a date` (custom date picker)
   - **Attachments** — check or uncheck files to include in the reminder note
   - **Recurrence** — if the original is recurring, choose *This instance only* or *Entire series*
   - **Location** — keep as Teams / convert to in-person, or keep the physical room
4. Click **"Create Follow-up"** at the bottom
5. A new appointment draft opens — review, add attachments manually if needed, and send

---

## Stopping and restarting the server

The server runs silently in the background. To stop it if needed:

1. Open **Task Manager** (`Ctrl + Shift + Esc`)
2. Go to the **Details** tab
3. Find `node.exe` → right-click → **End task**

To start it again without rebooting, double-click **`Start-FollowupAddin.bat`**.

---

## Removing auto-start

To stop the server from launching at login, open a **Command Prompt** and run:

```
schtasks /delete /tn "FollowupMeetingAddinServer" /f
```

Or open **Task Scheduler** (search in Start menu), find `FollowupMeetingAddinServer`, and delete it.

---

## Troubleshooting

### "Create Follow-up" button doesn't appear
- Make sure the server is running — open `https://localhost:3000` in your browser.
  You should see a page load, not a connection error.
- If the browser shows a certificate warning, re-run `Install.bat` to reinstall the trusted certificate.
- Remove and re-add the add-in in OWA (repeat the Sideloading steps above).

### OWA shows "App error" or a blank panel
- The server is likely not running. Start it with `Start-FollowupAddin.bat` and reload OWA.

### Certificate error / browser refuses localhost
The HTTPS certificate is machine-specific and expires after 30 days. To renew:
```
npx office-addin-dev-certs install --force
```
Then restart the server.

### "npm is not recognized" error when running the bat file
- Node.js is not installed or not in PATH.
- Install from https://nodejs.org (LTS) and restart your PC.

### The server was working, then stopped after a reboot
- The auto-start task may not have been registered (IT policy sometimes blocks it).
- Double-click `Start-FollowupAddin.bat` to start manually.
- Re-run `Install.bat` to try registering the auto-start task again.

### IT policy blocks Task Scheduler registration
- `Install.bat` automatically falls back to the Windows Startup folder.
- If that also fails, pin `Start-FollowupAddin.bat` to your taskbar and run it manually when needed.

---

## Project structure

```
outlook-followup-addin/
|
+-- Install.bat                     Run this ONCE on each new PC
+-- Start-FollowupAddin.bat         Manual server start (if needed)
+-- Start-FollowupAddin-Hidden.vbs  Called by auto-start task (no window)
+-- Setup-AutoStart.bat             Re-register auto-start task only
|
+-- manifest.json                   Unified Manifest (Outlook on the Web / New Outlook)
+-- manifest.xml                    Legacy XML manifest (Classic Outlook)
|
+-- package.json                    Node.js project config + npm scripts
+-- tsconfig.json                   TypeScript compiler config
+-- webpack.config.js               Build config + HTTPS dev server setup
|
+-- assets/                         Icon PNG files (16, 32, 80, 192 px)
+-- scripts/
|   +-- generate-icons.js           Regenerate icons: node scripts/generate-icons.js
|
+-- src/
    +-- commands/
    |   +-- commands.html           Hidden function-file (required by manifest)
    |   +-- commands.ts             (reserved for future ribbon commands)
    +-- taskpane/
        +-- taskpane.html           Pre-flight UI (HTML + embedded CSS)
        +-- taskpane.ts             Core logic (TypeScript)
```

---

## For developers — making changes

### Development workflow

```bash
# Start the dev server with hot-reload
npm start

# Build a production bundle (output -> dist/)
npm run build

# Validate the manifest
npm run validate

# Regenerate icon PNG files
node scripts/generate-icons.js
```

Changes to `.ts` or `.html` files are picked up automatically when using `npm start` (hot module replacement). Just reload the taskpane in OWA after saving.

### Deploying to production (hosting on a real server)

1. Run `npm run build` — output lands in `dist/`
2. Deploy `dist/` and `assets/` to any HTTPS web server
3. In `manifest.json` **and** `manifest.xml`, replace every occurrence of
   `https://localhost:3000` with your production server URL
4. Distribute the updated `manifest.json` via your Exchange admin or an internal add-in catalog

### Key technical decisions

| Decision | Why |
|---|---|
| UTC epoch math for +7 days | `date.getTime() + 604_800_000` is immune to DST and timezone changes. Using `setDate(d+7)` breaks twice a year. |
| `resolveField()` helper | `item.subject` / `item.start` / etc. are plain values in attendee (read) mode but async getter objects in organizer (compose) mode. One helper handles both transparently. |
| Triple fallback for appointment creation | OWA does not consistently expose `displayNewAppointmentFormAsync` (Mailbox 1.9). Falls back to `displayNewAppointmentForm`, then to an OWA calendar deep-link URL. |
| Body via `item.body.getAsync(Html)` | Works in both read and compose mode; the full original invitation HTML is carried over into the new appointment body. |
| Subject prefix loop | Strips `RE:/FW:/Antw:/WG:` in a loop until stable — handles arbitrarily nested prefixes like `RE: RE: FW: RE: Meeting`. |

---

## Compatibility

| Platform | Status |
|---|---|
| Outlook on the Web (OWA) | Fully supported |
| New Outlook for Windows | Fully supported |
| Classic Outlook for Windows | Use `manifest.xml`; IT may restrict add-in deployment |
| Outlook for Mac | Partially supported — test with your version |
| Outlook Mobile | Not supported (`displayNewAppointmentFormAsync` unavailable on mobile) |
