# The Next Repertoire website: setup

Two files: `index.html` (the site) and `Code.gs` (the backend, a Google Apps Script bound to a Google Sheet).

## 1. Backend (about 10 minutes)

1. Create a new Google Sheet in Drive. Name it something like `Next Repertoire responses`. Share it with Ryan.
2. In the Sheet: Extensions > Apps Script. Delete the default code, paste the contents of `Code.gs`, save.
3. In `setAdminKey()`, change `"change-me"` to the facilitator passcode you want. Select `setAdminKey` in the function dropdown and click Run. Approve the permissions prompt. Then change the passcode back to `"change-me"` in the code (the real one is stored separately) or leave it; it is only read when you run that function.
4. Deploy > New deployment. Type: Web app. Execute as: Me. Who has access: Anyone. Deploy. Copy the web app URL (ends in `/exec`).
5. Test: open the URL in a browser. You should see `{"gates":{"unlocked":[],"current":"welcome"},"updated":"..."}`.

If you change `Code.gs` later, use Deploy > Manage deployments > edit > New version. The URL stays the same.

## 2. Site

1. Open `index.html` and set `var API_URL = "";` (near the top of the script, under "configuration") to the web app URL.
2. Host the file. Options:
   - GitHub Pages: new repository, upload `index.html`, Settings > Pages > deploy from main. URL is `https://<user>.github.io/<repo>/`.
   - Netlify: drag the folder onto app.netlify.com/drop.
3. Open the site. The top-right chip should read "Live". Go to `#admin` (link in the footer), enter the passcode, and open a section to confirm publishing works.

Until `API_URL` is set, the site runs in preview mode: all sections open, submissions off.

## How it works

- Gates (which sections are open, and the "Now" marker) live in the `config` tab of the Sheet. Participants' pages check every 8 seconds.
- Each Submit button appends a row to a tab named after the form: `pre`, `demographics`, `quadrant`, `interests`, `table_reports`, `commit`, `post`. Columns are created from the field ids on first write.
- Every participant gets a random anonymous id (`pid`) stored in their browser, so `pre` and `post` rows can be paired without identifying anyone. The `table` column comes from the table they picked in section 04.
- Result panels (interest poll tallies, snapshot counts, confidence means, table statements) appear on participant pages only when you switch them on in the facilitator panel. They refresh every 15 seconds. Aggregates come from `?action=summary`, which returns counts and table-level statements only, never individual free-text answers.
- The facilitator panel (`#admin`) shows counts, interest-poll tallies by table (for sorting the room), the latest table report per table for each stage, mean confidence pre vs post, and demographic counts. It refreshes every 20 seconds and on Refresh.
- Participants' own answers stay in their browser (localStorage) and can be exported as Markdown or printed.

## Limits

- Apps Script handles roughly 30 simultaneous requests. With 50 to 80 participants polling every 8 to 10 seconds that is fine; for a much larger room raise `POLL_MS`.
- The passcode only protects the facilitator panel and gate changes. Submissions are open to anyone with the site URL, which is what anonymous participation requires.
- Section ids and field ids are used as Sheet column names and localStorage keys. Renaming them after the workshop starts would split data.
