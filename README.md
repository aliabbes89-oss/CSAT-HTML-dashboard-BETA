# Contact Centre KPI Explorer — auto-updating setup (Google Drive)

This folder is a complete, self-hosted version of the dashboard that refreshes
itself whenever you replace the Excel export in Google Drive — no chat with
Claude required after today.

## How it works

- `index.html` is the dashboard "shell" — same charts/filters/Response Review
  as before, but it no longer has data baked in. On load it does
  `fetch('./dashboard_data.json')`.
- `dashboard_data.json` is the small (~14MB) data file the shell reads. This
  is the file that gets replaced daily.
- `transform.py` converts a raw `lastminute_com_group_cc_post_call_responses_*.xlsx`
  export into `dashboard_data.json` — the exact same logic Claude used to
  build every version of this dashboard, verified to produce byte-identical
  output.
- `.github/workflows/update-data.yml` is a GitHub Action that runs on a
  schedule, downloads the latest export from Google Drive, and — only if the
  file actually changed — regenerates `dashboard_data.json`, commits it, and
  pushes. GitHub Pages then redeploys automatically.

## One-time setup (about 15–20 minutes)

### 1. Share the file and get its ID
1. In Google Drive, right-click the daily export file → **Share** → change
   access to **"Anyone with the link" / Viewer**.
2. Copy the share link. It looks like:
   `https://drive.google.com/file/d/1AbCdEfGhIJkLmNoPQRstuVWxyz1234567/view?usp=sharing`
3. The long string between `/d/` and `/view` is the **file ID** — copy just
   that part, e.g. `1AbCdEfGhIJkLmNoPQRstuVWxyz1234567`. You'll need it in
   step 3.

### 2. Create a GitHub repository
1. github.com → New repository → e.g. `kpi-dashboard` → keep it **Public**
   (needed for the free version of GitHub Pages) or **Private** if you have
   GitHub Pro/Team/Enterprise.
2. Upload every file from this package, keeping the same folder structure
   (including the `.github/workflows/update-data.yml` path — GitHub only
   recognizes workflows in exactly that location).

### 3. Add the file ID as a secret
Repository → **Settings → Secrets and variables → Actions → New repository
secret**:
- Name: `GDRIVE_FILE_ID`
- Value: the file ID from step 1 (just the ID, not the full URL)

### 4. Turn on GitHub Pages
Repository → **Settings → Pages** → Source: **Deploy from a branch** →
Branch: `main`, folder `/ (root)` → Save.

After a minute or two, your dashboard is live at:
`https://<your-username>.github.io/<repo-name>/`

That's your permanent link. Share it with anyone — no login needed, and it
keeps working even if you never open Claude again.

### 5. Test the automation
Repository → **Actions** tab → select "Update KPI dashboard data" →
**Run workflow** (manual trigger, top right). Watch it run — it should
download the file, detect it changed (first run always counts as changed),
rebuild `dashboard_data.json`, and push the commit. Refresh your GitHub
Pages URL after ~1 minute and it'll be live.

From then on it also runs automatically every day at 06:00 UTC. To change
the time, edit the `cron: '0 6 * * *'` line in
`.github/workflows/update-data.yml` (the two numbers are minute and hour,
in UTC).

## Making it truly instant (optional upgrade)

The schedule above checks once a day. If you want the dashboard to update
within a minute or two of you replacing the file — not wait for the next
scheduled run — add a **Google Apps Script** bound to the file's folder:

1. In Drive, open **script.google.com** → New project.
2. Paste this, filling in your repo details and a GitHub token:

   ```javascript
   function checkForChange() {
     const fileId = 'YOUR_GDRIVE_FILE_ID';
     const file = DriveApp.getFileById(fileId);
     const lastModified = file.getLastUpdated().getTime();

     const props = PropertiesService.getScriptProperties();
     const prevModified = props.getProperty('lastModified');

     if (String(lastModified) === prevModified) return; // no change

     props.setProperty('lastModified', String(lastModified));

     UrlFetchApp.fetch('https://api.github.com/repos/YOUR_USERNAME/YOUR_REPO/dispatches', {
       method: 'post',
       contentType: 'application/json',
       headers: {
         Authorization: 'Bearer YOUR_GITHUB_TOKEN',   // a GitHub Personal Access Token, "repo" scope
         Accept: 'application/vnd.github+json'
       },
       payload: JSON.stringify({ event_type: 'gdrive-file-changed' })
     });
   }
   ```

3. In the Apps Script editor: **Triggers** (clock icon on the left) → **Add
   Trigger** → choose `checkForChange`, event source **Time-driven**,
   **Minutes timer**, every 5 or 10 minutes.
4. Add this to the workflow file's `on:` block so it also listens for that
   trigger:
   ```yaml
   on:
     schedule:
       - cron: '0 6 * * *'
     workflow_dispatch: {}
     repository_dispatch:
       types: [gdrive-file-changed]
   ```

This polls Drive every few minutes from Google's side (free, no extra
infrastructure) and only pings GitHub when something actually changed — so
in practice the dashboard updates within minutes of you replacing the file,
instead of waiting for the next 06:00 UTC run.

## A note on file size

The export is close to 90–100MB. Google Drive shows an extra "can't scan
this file for viruses" confirmation page for files over 100MB, which breaks
simple direct-download links — that's why the workflow uses the `gdown`
tool instead of a plain `curl` link; it handles that automatically. If you
ever swap `transform.py`/the workflow for your own script, keep using
`gdown` (or an equivalent) rather than a raw download link once the file
crosses that size.

## Updating the dashboard's look/logic itself

If you ever want to change the dashboard's design, filters, or charts (not
just the daily data), that still needs to go through Claude — ask for the
updated `index.html`, then re-upload just that one file to the repo.
`dashboard_data.json` keeps refreshing on its own regardless.
