# Ugly Sweater Christmas Potluck — RSVP

Saturday, December 5, 2026 · 5:00 PM · 40008 Maddie Drive, Prairieville, LA 70769
RSVP by Tuesday, December 1.

**RSVP page:** https://k8tye.github.io/Christmas-Potluck-2026/

## How it works

```
Guest scans QR / taps link ──► RSVP page (GitHub Pages, free)
                                   │  sends answers
                                   ▼
                    Google Apps Script (free, runs as you)
                       ├─► adds/updates a row in your Google Sheet
                       ├─► emails you the RSVP
                       └─► returns the "Who's bringing what" list (dishes & drinks only)
```

| File | What it is |
|---|---|
| `index.html` | The RSVP page |
| `google-apps-script/Code.gs` | The script that saves RSVPs to your Google Sheet |
| `assets/invite-with-qr.jpg` | **Your invite with a working QR code.** Send this one. |
| `assets/rsvp-qr.png` / `.svg` | The QR code on its own |
| `invite.ics` | Calendar file for iPhone "Add to calendar" |

## Setup (about 10 minutes, one time)

### 1. Turn on the website (GitHub Pages)
1. On GitHub, open this repo and go to **Settings → Pages**.
2. Under **Build and deployment**, choose **Deploy from a branch**, branch **main**, folder **/ (root)**, then **Save**.
3. After about a minute, https://k8tye.github.io/Christmas-Potluck-2026/ goes live.
   - Free GitHub Pages needs the repo to be **public**. That's fine: guests' answers go to your private Google Sheet, never to GitHub.

### 2. Create the Google Sheet + script
1. Go to [sheets.new](https://sheets.new) and name the sheet something like **Potluck RSVPs 2026**.
2. Click **Extensions → Apps Script**.
3. Delete what's in the editor, paste in all of `google-apps-script/Code.gs`, and click **Save** (💾).
4. In the function dropdown at the top, pick **setup** and click **Run**. Google asks for permission:
   **Review permissions → your account → Advanced → Go to (project) (unsafe) → Allow**.
   It says "unsafe" only because you wrote the script yourself and Google hasn't reviewed it.
5. Click **Deploy → New deployment**, click the ⚙️ next to "Select type", and pick **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy** and **copy the Web app URL** (it ends in `/exec`).

### 3. Connect the page to the sheet
In `index.html`, find this line near the bottom:

```js
const RSVP_ENDPOINT = "";
```

Paste your Web app URL between the quotes. On GitHub you can click the file, then the ✏️ pencil, then **Commit changes**.

### 4. Test it
Open the RSVP page on your phone and submit a test RSVP. A new row should appear in the sheet, and you should get an email. Delete the test row afterward.

## Day-to-day
- **See all RSVPs:** open the Google Sheet (it's in your Google Drive). **File → Download → Microsoft Excel** if you want an Excel file.
- **Notifications:** every new or changed RSVP emails you. To also notify a co-host, add their address to `NOTIFY_EMAIL` at the top of the script, then **Deploy → Manage deployments → ✏️ → Version: New version → Deploy**.
- **Changing an RSVP:** guests submit again with the same name, and their row updates instead of duplicating.
- **Privacy:** the public "Who's bringing what" list shows only dishes, drinks and a head count. Names, contact info and dietary details stay in your sheet.

## Sending the invite
Use `assets/invite-with-qr.jpg` (the original's QR code was decorative and didn't scan), and **always include the link as text**, because people reading on their phone can't scan their own screen:

> You're invited to our Ugly Sweater Christmas Potluck! 🎄
> Sat, Dec 5 at 5 PM. Please RSVP by Dec 1. Scan the code or tap here:
> https://k8tye.github.io/Christmas-Potluck-2026/
