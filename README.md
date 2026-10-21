# MVB Serve & Pass Stats

Live serve and serve-receive stats for NYU men's volleyball. The app (one web page) is the input tool; two Google Sheets are where the numbers live and where coaches look: **MVB GAME Serve and Pass Master** (one tab per game) and **MVB PRACTICE Serve and Pass Master** (one tab per practice day). Same layout in both.

Files:

- `index.html` and `sw.js` — the app. Host on GitHub Pages (or open `index.html` on a laptop).
- `Code.gs` — the Google Apps Script that writes into the sheet.
- The two master workbooks in Drive (GAME and PRACTICE), converted to Google Sheets.

## One-time setup (about 15 minutes)

### 1. The Google Sheets

1. In Drive, open `MVB GAME Serve and Pass Master.xlsx` and choose **File > Save as Google Sheets**. Do the same for `MVB PRACTICE Serve and Pass Master.xlsx`. Work from the Google Sheets from now on; the .xlsx copies can be deleted.
2. Both have the tabs `Face Page`, `Roster`, `Sessions`, `Log`, and a hidden `Template`. Type player names into `Roster` column B in each (the app reads the roster from the GAME sheet).
3. Open the PRACTICE sheet and copy its ID from the address bar: the long string between `/d/` and `/edit`. You need it in step 2.
4. Share both sheets with the coaches as **Viewer** (or Commenter). The app does not need them to have edit access.

### 2. The script

1. In the **GAME** Google Sheet: **Extensions > Apps Script**. (One script serves both workbooks; it reaches the practice one by ID.)
2. Delete the sample code in `Code.gs`, paste in the contents of this repo's `Code.gs`, and save (Ctrl/Cmd+S).
3. **Project Settings (gear icon) > Script properties**. Add two properties: `PASSCODE` = anything you like (e.g. `nyumvb-2026`), and `PRACTICE_SHEET_ID` = the ID you copied from the practice sheet. Save.
4. Optional check: back in the editor, pick `selfTest` in the function dropdown and click **Run**. Approve the permissions prompt (it asks because the script edits your sheets). A tab `G1 <today> SELFTEST` appears in the GAME sheet with two entries for #1. Then run `selfTestPractice`; a tab named today's date appears in the PRACTICE sheet. Delete those tabs, their `Log` rows and their `Sessions` rows afterwards. (The Face Page formulas keep a harmless reference to a deleted tab; to be tidy, restore each Face Page from **File > Version history** to just before the test.)
5. Deploy: **Deploy > New deployment**. Type: **Web app**. Description: anything. Execute as: **Me**. Who has access: **Anyone**. Click **Deploy**, approve permissions, and copy the **Web app URL** (ends in `/exec`).

"Anyone" means anyone with the URL can call the script, which is why the passcode exists. The sheet itself stays private to whoever you share it with.

Whenever you change `Code.gs` later: **Deploy > Manage deployments > edit (pencil) > Version: New version > Deploy**. The URL stays the same.

### 3. The app

1. Put `index.html` and `sw.js` in a public GitHub repository. **Settings > Pages > Deploy from a branch** (main, root). The site is live a minute later at `https://<user>.github.io/<repo>/`.
2. On your own laptop, open the site, click **Not set up**, paste the web app URL and the passcode, and click **Test & save**. It should say "Connected" and load the roster from the sheet.
3. Under the status it now shows a **setup link**. Copy it and text it to the manager (and anyone else who will take stats). Opening that link once configures their device; after that the plain site URL is all they need. Settings are saved per device, so nobody types the URL again.
4. On the phone: Share > **Add to Home Screen**.

A game recorded before the connection was set up can be pushed to the sheet afterwards: open it, tap the status pill, then **Send this game / practice to the sheet again**.

## How it works day to day

- **Game**: pick Game, type the opponent, set the date, Home/Away, Start. Tap the players on the court (include the libero), Start set.
- **Practice**: pick Practice, set the date, Start. Everyone is preselected; untap anyone absent.
- Tap a player, then 0–4 under **Serve** or **Pass**. **Undo last** removes the most recent entry (from the sheet too).
- **Sub**: who is going out, then who is coming in. **New set / New round**: carries the lineup over.
- The pill in the top right shows **Live**, **N pending** (taps waiting for signal), or **offline**. Pending taps send on their own when the connection is back, in order, so nothing is lost if the gym Wi-Fi drops.

Each game gets its own tab in the GAME sheet (`G3 Nov 2 Stevens`) and each practice day its own tab in the PRACTICE sheet (`Oct 27`), in the usual layout. Each sheet's Face Page totals include its tabs automatically, `Sessions` lists every tab, and `Log` has one row per entry.

## If something goes wrong

- **"bad passcode"**: the passcode typed in the app doesn't match the script property.
- **"Unexpected reply from the script"**: the deployment isn't set to "Anyone", or the URL is the editor URL instead of the `/exec` web app URL.
- **"PRACTICE_SHEET_ID script property is not set"**: step 2.3 above; a practice was started before the practice sheet ID was added. Add it, then in the app tap the status pill and **Retry now**.
- **Numbers on a tab look wrong after someone edited cells by hand**: in Apps Script, run `rebuildWholeTab('G3 Nov 2 Stevens')` (games) or `rebuildWholeTab('Oct 27', 'practice')` to redo that tab from the Log.
- **Games started on another device (or by selfTest)** show up in the app's list automatically when it is connected; opening one loads its entries from the Log so it can be continued or deleted from any device.
- **Deleting a game or practice**: the trash icon next to it on the home screen (or the button at the bottom of its page). This removes its tab, its Log and Sessions rows, and its share of the Face Page totals from the workbook as well. There is no undo beyond the sheet's File > Version history.
- **Stuck queue**: open the connection settings in the app; **Retry now**, or **Discard queue** as a last resort (the entries stay on the device's Game totals but won't reach the sheet).
