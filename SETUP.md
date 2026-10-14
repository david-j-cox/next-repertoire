# The Next Repertoire website: setup

Two files run the site: `index.html` (the site) and `Code.gs` (the backend, a Google Apps Script bound to a Google Sheet). Two optional helper scripts, `DryRun.gs` and `ResetData.gs`, go in the same Apps Script project (see Section 3).

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
- Three sections are split across two pages. The second page has no switch of its own; it opens and locks with the first.
  - 04 Choose (interest poll) and 04 My Table (seat, problem statement, mission cards). Submitting the poll moves people to My Table.
  - 11A Ideation Practice (Rep 1) and 11A Ideation Rep 2 (after Break 2).
  - 13A Chiron Field Lab (map and before-reveal record) and 13A Field Lab Reveal (reveal, revise, report). The two measure boxes appear on both pages and hold the same answer.
- Switches and room-result toggles are staged until you click Publish. Publish appears at the top of the panel and again under the room-result switches whenever there are unpublished changes.
- Each Submit button appends a row to a tab named after the form: `pre`, `demographics`, `quadrant`, `interests`, `table_reports`, `commit`, `post`. Columns are created from the field ids on first write.
- Every participant gets a random anonymous id (`pid`) stored in their browser, so `pre` and `post` rows can be paired without identifying anyone. The `table` column comes from the table they picked in section 04.
- Sorting the room (facilitator panel, Sort the Room): people go to a topic (A-F) by first, second, then third choice, and each topic splits into at most "Max tables per topic" physical tables of about "People per table" (A-1, A-2). Expected roles are dealt across the tables so each gets a mix. Results go to the `assignments` tab. Mark the table tents to match (A-1, A-2).
- Table sync: once someone has a table group (from the sort, or by picking table and table number in 04), every Table recorder box is shared with their tablemates. Edits appear on the others' screens within about 5 seconds. A box someone is typing in is never overwritten. Each person can switch sync off ("Writing on my own"). Table reports record `group` and `voice` (table or own). Shared entries are kept in the `group_notes` tab, one row per table.
- Mission cards in 04 stay hidden until the facilitator switches on "Mission Cards (04)".
- Each section header shows the matching deck slides (Oct 9 deck labels), and so does the facilitator panel.
- The Field Lab context-shift reveal (13A) is one of the facilitator switches. The shift text is held in `Code.gs` and served only after you open the reveal, so it is not in the page source beforehand. Participants see the shift matching the table they picked in section 04.
- Result panels (interest poll tallies, snapshot counts, confidence means, table statements) appear on participant pages only when you switch them on in the facilitator panel. They refresh every 15 seconds. Aggregates come from `?action=summary`, which returns counts and table-level statements only, never individual free-text answers.
- The facilitator panel (`#admin`) shows counts, interest-poll tallies by table (for sorting the room), the latest table report per table for each stage, mean confidence pre vs post, and demographic counts. It refreshes every 20 seconds and on Refresh.
- Participants' own answers stay in their browser (localStorage) and can be exported as Markdown or printed.

## 3. Helper scripts (optional)

Add each as a new script file in the Sheet's Apps Script project (Extensions > Apps Script > + > Script). No redeploy is needed; run them from the editor and read the Execution log.

- `DryRun.gs`, `dryRunSort()`: adds 14 fake interest-poll rows, runs the real sort with tables of 3, logs each seat, then removes the fake rows and clears `assignments`. Run it before participants sort, never after, because it clears real seats.
- `ResetData.gs`, `previewReset()` then `resetAll()`: clears every response row, `assignments`, and `group_notes` (keeping header rows), empties the backend cache, and turns off all room-result switches. Opened sections stay open. Use it before a run-through or the live session. After the workshop, save what you need first.

Test data also lives in the tester's browser. Clear it with the site icon left of the address bar > Cookies and site data > Manage on-device site data > delete github.io, or test in a private window.

## Limits

- Apps Script handles roughly 30 simultaneous requests. With 50 to 80 participants polling every 8 to 10 seconds that is fine; for a much larger room raise `POLL_MS`.
- The passcode only protects the facilitator panel and gate changes. Submissions are open to anyone with the site URL, which is what anonymous participation requires.
- Section ids and field ids are used as Sheet column names and localStorage keys. Renaming them after the workshop starts would split data.
