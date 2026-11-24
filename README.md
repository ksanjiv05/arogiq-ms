# Sheet Forms

A phone-friendly web app for viewing and editing a Google Sheet. The page is plain HTML, CSS and JavaScript hosted on GitHub Pages. A small Google Apps Script attached to the sheet does the reading and writing.

You connect the sheet once. After that, anyone you give the link and access key can use it from any phone or browser. They don't need a Google account or access to the sheet itself.

## What it does

- Shows the rows of any visible tab in the sheet. Row 1 is treated as the column names.
- **Columns** button: choose which columns to **Show** in the list and which to **Edit**.
- **Views**: an admin saves column choices as named views (for example "Reception" or "Billing"). Everyone sees the saved views and can switch between them. Other users can change columns for themselves, but only until they reload the page.
- Tap a row to open a form with only the columns marked **Edit**.
- **+ Add row** opens a form with every column and adds the row below the last one.
- Search, and a card or table layout.

The forms adapt to the sheet:

| In the sheet | In the app |
|---|---|
| Dropdown (data validation) | A picker; "Reject input" dropdowns accept only their options |
| Checkbox | A checkbox |
| Mostly numbers | A number keyboard |
| Formula cells, and columns filled by `ARRAYFORMULA`, `QUERY`, `FILTER` and similar | Read-only |
| A tab whose cell A1 holds `QUERY`, `IMPORTRANGE`, `FILTER` or similar (a report tab) | Read-only |
| Other "Reject input" data validation (numbers, dates, …) | Enforced: nothing is saved and the app shows the sheet's message |
| Protected ranges (not "Show a warning") | Read-only |

New rows copy the formatting, dropdowns, checkboxes and per-row formulas of the row above.

## How it's put together

```
Phone browser ──> GitHub Pages (index.html, css/, js/)
      │
      └─────────> Apps Script web app (apps-script/Code.gs, runs as you) ──> your Google Sheet
```

- `js/config.js` holds the Apps Script URL. Everyone who opens the site uses that URL, so the connection is the same on every device.
- Saved views are stored with the sheet (Apps Script document properties), so they are shared too.
- The only things a device remembers are conveniences: the access key once entered, the last tab and view used, and the card/table layout.

## Setup (once)

### 1. Add the script to your sheet

1. Open the Google Sheet. Choose **Extensions → Apps Script**.
2. Delete everything in `Code.gs`, paste in the full contents of [apps-script/Code.gs](apps-script/Code.gs), and click **Save** (the disk icon).

### 2. Set access keys (strongly recommended)

The script URL ends up in a public file, so without a key anyone who finds it could read and edit the sheet.

1. In Apps Script, click **Project Settings** (gear icon on the left) and scroll to **Script properties**.
2. Click **Add script property**. Set the name to `ACCESS_KEY` and the value to your key.
3. Optionally add a second property named `ADMIN_KEY` with a different key. Only people with the admin key can save or delete views. Without it, everyone with the access key can manage views.
4. Click **Save script properties**. The keys are not saved until you do.

Choosing keys:

- Use at least 20 random letters and digits, for example from a password manager.
- Use only letters and digits, so the keys work in share links.
- The admin key also opens the app, so it must be just as strong.

You can change keys at any time; changes take effect immediately.

### 3. Deploy the script as a web app

1. Click **Deploy → New deployment**. Click the gear next to "Select type" and choose **Web app**.
2. Set **Execute as: Me** and **Who has access: Anyone**.
3. Click **Deploy** and authorize the script. Google warns that the app isn't verified, because it is your own script: click **Advanced → Go to (project name) (unsafe)**, then **Allow**.
4. Copy the **Web app URL**. It ends in `/exec`.

To check it, open the URL in a browser. It should show `{"ok":true,"data":{"service":"Sheet Forms",...}}`.

On a Google Workspace account, your admin may not allow "Anyone". The app needs "Anyone" to work for people who aren't signed in to Google.

### 4. Connect the web app

Put the URL in [js/config.js](js/config.js):

```js
window.SHEET_FORMS_CONFIG = {
  apiUrl: 'https://script.google.com/macros/s/XXXXXXXX/exec'
};
```

Never put access keys in this file.

### 5. Publish on GitHub Pages

The simple way, in a browser:

1. On GitHub, create a new repository (**+ → New repository**). Pick a name and choose **Public**; GitHub Pages on a free account needs a public repository.
2. On the new repository's page, click **uploading an existing file**. Drag in `index.html`, `.nojekyll`, and the `css`, `js` and `apps-script` folders, then click **Commit changes**. You don't need to upload the `.tmp` folder.
3. Go to **Settings → Pages**. Under **Build and deployment**, choose **Deploy from a branch**, then `main` and `/ (root)`, and click **Save**.
4. After a minute or two the app is at `https://YOUR-USERNAME.github.io/YOUR-REPO/`.

Or with git:

```sh
git init
git add index.html .nojekyll css js apps-script README.md
git commit -m "Sheet Forms"
git branch -M main
git remote add origin https://github.com/YOUR-USERNAME/YOUR-REPO.git
git push -u origin main
```

Then turn on Pages as in step 3.

### 6. Share it

Send people the link with the key after `#key=`. For example:

```
https://YOUR-USERNAME.github.io/YOUR-REPO/#key=THEACCESSKEY
```

The part after `#` is never sent to a server. The app saves the key on that device and removes it from the address bar. The browser's own history may still remember the full link, so on a shared computer, open the plain link and type the key instead.

Admins use the same link with the admin key instead.

## Updating the script later

After you edit `Code.gs`, the live web app keeps running the old version until you redeploy it:

1. Go to **Deploy → Manage deployments**, select the deployment, and click the edit (pencil) icon.
2. Set **Version: New version** and click **Deploy**.

The URL stays the same. Choosing **New deployment** instead creates a new URL.

## Test locally

Serve the folder rather than opening the file directly:

```sh
python3 -m http.server 8000
```

Then open <http://localhost:8000>.

## Good to know

**What people can see.** Anyone with a valid key can see every visible tab, including hidden columns, hidden rows and filtered-out rows in those tabs. To keep something private, move it to a separate tab and hide that tab: hidden tabs are never sent to the app.

**Formulas and protection.**
- Text starting with `=`, `+`, `-` or `@` is saved as plain text, so nobody can run formulas with your permissions through the app.
- Numbers with a leading zero (phone numbers, IDs) and very long numbers are saved as text, so they keep every digit.
- If saving into a cell would break an array formula's results, the app leaves that cell out and says so.
- Protected ranges are read-only in the app, unless they are set to "Show a warning".

**How to lay out the sheet.**
- Give every column a header in row 1. Columns without one are called "Column D" and so on, and saved views lose track of them when columns are inserted or deleted.
- A totals row under the data that is made only of formulas is never written over; new rows are inserted above it. Check that its formulas cover the new rows. A formula like `=SUM(D$2:INDEX(D:D,ROW()-1))` always adds up everything above it. A totals row that also has typed text (like "Total") counts as data, so new rows go below it. It is simplest to keep totals in another tab or above the header.
- A formula that fills several columns (for example `=FILTER(Orders!A2:C, …)`) only makes its own column read-only. The other columns it fills still look editable. The app refuses to save over them, but such results are best kept in their own tab.
- If saving a value would make a formula show `#REF!` (for example a lookup that suddenly returns two rows), the app leaves that field unsaved, keeps its old value, and says so.

**Two people editing at once.** Before saving, the app checks whether someone else changed the row since it was opened. If they did, it shows what changed and asks before overwriting.

**Seeing others' changes.** The app reloads when you come back to it after a couple of minutes. **Refresh** reloads at any time.

**Speed.** Apps Script typically takes 1–3 seconds per request, and longer for tabs with many thousands of rows.

**Edit history.** Every edit is made by the script as the sheet owner. The sheet's version history shows the owner's name, not the person using the app.
