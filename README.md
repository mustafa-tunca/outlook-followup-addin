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

| Requirement | Where to get it | Notes |
|---|---|---|
| **Node.js 18 LTS or newer** | https://nodejs.org → "LTS" download | Includes `npm`. Restart your PC after install. |
| **Microsoft 365 account** | Your company / personal M365 account | Outlook Web App (OWA) access required |
| **A modern browser** | Edge, Chrome, Firefox | For OWA |
| **Windows 10 / 11** | — | The auto-start scripts are Windows-only |

> **Corporate IT note:** The add-in runs entirely on your own PC (`https://localhost:3000`).
> No data is sent to any external server. The HTTPS certificate is self-signed and only trusted locally.

---

## First-time setup (new PC)

### Step 1 — Install Node.js

1. Go to **https://nodejs.org** and download the **LTS** version.
2. Run the installer, keep all defaults, click through to finish.
3. **Restart your PC** so `npm` becomes available in the command prompt.

### Step 2 — Get the project files

**Option A — Clone from GitHub (recommended):**
```
git clone https://github.com/mustafa-tunca/outlook-followup-addin.git
```

**Option B — Download ZIP:**
Click the green **Code** button on GitHub → **Download ZIP** → extract to any folder.

> Keep the folder in a permanent location. The auto-start task remembers the exact path.
> If you move the folder later, re-run `Install.bat`.

### Step 3 — Run the installer

Double-click **`Install.bat`** inside the folder.

> **Windows may warn** "Windows protected your PC". Click **More info → Run anyway**.
> This is expected for unsigned scripts.

The installer will:
1. Check that Node.js is available
2. Download all required packages (`npm install`)
3. Build the add-in (`npm run build`)
4. Install and trust the local HTTPS certificate *(a UAC prompt will appear — click **Yes**)*
5. Register a background task so the server starts **silently** at every login (no window)
6. Start the server immediately for this session

---

## Sideloading into Outlook Web App (OWA)

This step is done **once per browser / per user account**.

1. Open **Outlook Web App** in your browser (e.g. `https://outlook.office.com`)
2. Click any **calendar appointment** to open it
3. In the appointment window, click **"..."** (More options) in the toolbar
4. Choose **"Get Add-ins"** (or **"Manage add-ins"**)
5. In the dialog, click **"My add-ins"** in the left panel
6. Scroll to the bottom → **"+ Add a custom add-in"** → **"Add from file..."**
7. Browse to your `outlook-followup-addin` folder and select **`manifest.json`**
8. Confirm any security warning and close the dialog

The **"Create Follow-up"** button now appears in the appointment ribbon.

> **Tip:** If the button doesn't appear, close and reopen the appointment.

---

## Daily use

### The server starts automatically

After running `Install.bat` once, the dev server starts **silently in the background** every time you log into Windows — no window, no notification. It is ready within about 15 seconds of login.

The server runs at `https://localhost:3000` and serves the add-in to OWA.

### Using the add-in

1. Open any calendar appointment in OWA
   *(works for meetings you organised **and** meetings you were invited to)*
2. Click **"Create Follow-up"** in the ribbon
3. A side panel opens with pre-filled details:
   - **Subject** — editable, already cleaned and numbered
   - **Date** — toggle between `+7 days` (automatic) or `Pick a date` (custom date picker)
   - **Attendees / Location / Body** — carried over from the original meeting
   - **Attachments** — check or uncheck files to include in the reminder note
4. Click **"Create Follow-up"** at the bottom
5. A new appointment draft opens — review, add attachments manually if needed, and send

---

## Stopping and restarting the server

The server runs silently in the background. To **stop** it:

1. Open **Task Manager** (`Ctrl + Shift + Esc`)
2. Go to the **Details** tab
3. Find `node.exe` → right-click → **End task**

To **start** it again without rebooting, double-click **`Start-FollowupAddin.bat`** (this opens a visible window so you can see the server status).

---

## Re-registering auto-start (after moving the folder)

If you move the project folder, the auto-start task will point to the wrong path.
Fix it by running **`Setup-AutoStart.bat`** from the new location — it re-registers both the Startup folder shortcut and the Task Scheduler entry.

---

## Removing auto-start

To stop the server from launching at login:

```
schtasks /delete /tn "FollowupMeetingAddinServer" /f
```

And delete `FollowupAddinServer.lnk` from:
```
%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\
```

---

## Troubleshooting

### "Create Follow-up" button doesn't appear
- Check the server is running: open `https://localhost:3000` in your browser.
  You should see a page (not a connection error).
- If the browser shows a certificate warning, re-run `Install.bat` to reinstall the certificate.
- Remove and re-add the add-in in OWA (repeat the Sideloading steps above).

### OWA shows "App error" or a blank panel
- The server is likely not running. Start it with `Start-FollowupAddin.bat` and reload OWA.

### Certificate error / browser refuses localhost
The dev certificate may have expired. To renew:
```
npx office-addin-dev-certs install --force
```
Then restart the server.

### "npm is not recognized" error
- Node.js is not installed or not in PATH.
- Install from https://nodejs.org (LTS) and **restart your PC**.

### The server was working, then stopped after a reboot
- The auto-start task may not have been registered (IT policy sometimes blocks Task Scheduler).
- Run `Setup-AutoStart.bat` to try registering it again.
- If that also fails, double-click `Start-FollowupAddin.bat` to start manually.

### Bat files open and immediately close / show a syntax error
- Make sure you are running the `.bat` files from the correct project folder.
- Try right-clicking the bat file and choosing **"Run as administrator"** (needed for certificate install step).

### Install.bat shows "The system cannot find the path specified"
- This usually means Node.js is not in PATH yet. Restart your PC after installing Node.js.

---

## Project structure

```
outlook-followup-addin/
|
+-- Install.bat                     Run ONCE on each new PC
+-- Start-FollowupAddin.bat         Manual server start (visible window)
+-- Start-FollowupAddin-Hidden.vbs  Called by auto-start (no window)
+-- Setup-AutoStart.bat             Re-register auto-start after moving the folder
|
+-- manifest.json                   Unified Manifest (OWA / New Outlook)
+-- manifest.xml                    Legacy XML manifest (Classic Outlook)
|
+-- package.json
+-- tsconfig.json
+-- webpack.config.js
|
+-- assets/                         Icon PNG files (16, 32, 80, 192 px)
+-- scripts/
|   +-- generate-icons.js           Regenerate icons: node scripts/generate-icons.js
|
+-- src/
    +-- commands/
    |   +-- commands.html
    |   +-- commands.ts
    +-- taskpane/
        +-- taskpane.html           Side-panel UI (HTML + embedded CSS)
        +-- taskpane.ts             Core logic (TypeScript)
```

---

## How auto-start works

The setup registers **two** mechanisms so the server starts reliably even if one is blocked by IT policy:

| Mechanism | Trigger | Delay | Notes |
|---|---|---|---|
| **Startup folder shortcut** | Login | ~12 s | `%APPDATA%\...\Startup\FollowupAddinServer.lnk` → wscript.exe → VBScript → bat |
| **Task Scheduler** (`FollowupMeetingAddinServer`) | Login | 20 s | Backup; runs the same VBScript |

Both call `Start-FollowupAddin-Hidden.vbs`, which starts `Start-FollowupAddin.bat` with a hidden window. No console window ever appears.

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

Changes to `.ts` or `.html` files are picked up automatically when using `npm start`. Reload the taskpane in OWA after saving.

### Deploying to production (hosting on a real server)

1. Run `npm run build` — output lands in `dist/`
2. Deploy `dist/` and `assets/` to any HTTPS web server
3. In `manifest.json` **and** `manifest.xml`, replace every `https://localhost:3000` with your production URL
4. Distribute the updated manifest via your Exchange admin or an internal add-in catalog

### Key technical decisions

| Decision | Why |
|---|---|
| UTC epoch math for +7 days | `date.getTime() + 604_800_000` is immune to DST. Using `setDate(d+7)` breaks twice a year. |
| `resolveField()` helper | `item.subject` / `item.start` etc. are plain values in attendee (read) mode but async getter objects in organizer (compose) mode. One helper handles both. |
| Triple fallback for appointment creation | OWA does not consistently expose `displayNewAppointmentFormAsync`. Falls back to `displayNewAppointmentForm`, then to an OWA calendar deep-link URL. |
| `item.body.getAsync(Html)` | Works in both read and compose mode; carries the full original invitation HTML into the new appointment. |
| Plain ASCII in all `.bat` files | cmd.exe on non-English Windows reads batch files using the system ANSI code page. Unicode characters (even in comments) cause silent syntax errors. |

---

## Compatibility

| Platform | Status |
|---|---|
| Outlook on the Web (OWA) | ✅ Fully supported |
| New Outlook for Windows | ✅ Fully supported |
| Classic Outlook for Windows | Use `manifest.xml`; IT may restrict sideloading |
| Outlook for Mac | Partially supported — test with your version |
| Outlook Mobile | ❌ Not supported |

---

## License

Apache 2.0 — Copyright 2026 [Mustafa Tunca](https://mustafatunca.com)
