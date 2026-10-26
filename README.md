# Eagle Deep Dive Daily

A starting point for the Eagle Deep Dive Daily website. The project currently
uses plain HTML and has no build tools or runtime dependencies.

## Preview

Open `index.html` in a web browser. Add the site's content and assets there as
they are ready.

## Smore newsletters: Google Sheet + Apps Script setup

Flight Wire's **Read newsletter** button opens an on-site article. Instant search
includes its title, summary, and full body. The original Smore link remains available.
Only import public newsletters you have permission to republish.

### 1. Prepare the same spreadsheet

Keep your existing `submissions` tab and submission script unchanged. **Back up
your spreadsheet first.** The old link-per-row feed cannot supply full articles.
Rename the old `smore_feed_raw` and `smore_feed_public` tabs to names ending in
`_backup` (do not delete them). New tabs use **one row per newsletter issue**.

In **Extensions → Apps Script**, add a script file named `Smore` and paste the
contents of [`apps-script/Smore.gs`](apps-script/Smore.gs).
Remove/replace the old newsletter sync functions and **old `doGet`**; there must
be exactly one `doGet`. Keep your existing `doPost`, `cleanText_`, and submission
helpers. The new newsletter helpers are prefixed `smore` to avoid collisions.

Run `setupSmoreFeed` from the function dropdown and authorize spreadsheet and
external-fetch access. It creates these tabs/headers:

`smore_feed_raw!A1` and `smore_feed_public!A1`:

```text
issue_url	issue_title	published_at	summary	cover_image_url	article_body_text
```

`settings!A1`:

```text
key	value
```

If `settings` already has different headers, back it up and migrate its values
to this key/value layout before running setup. Keep other settings rows.
Setup adds this row if missing; edit its value to your public profile or a single
newsletter URL:

```text
smore_source_url	https://www.smore.com/u/maizemiddleschool
```

### 2. Import and check the content

1. Run `syncSmoreFeed`. It discovers **issue URLs only**, fetches each issue, and
   updates existing rows by canonical URL (tracking parameters/fragments removed).
   Links *inside* a newsletter stay in its body; they do not become separate issues.
2. Open **Executions** and inspect the sync log, then compare a
   `smore_feed_public` row with the original newsletter. Check headings, lists,
   dates, links, and the **end** of the article—not just the summary.
3. To import older issues not linked by the profile, paste their actual issue URLs
   into column A of `smore_feed_raw`, one URL per row, and sync again. Do not paste
   resource links or the profile URL into these rows.

**Source limitations:** live Smore markup could not be fetched in the development
environment. The importer supports JSON-LD `articleBody`, or server-rendered
`article`/`main`/recognized newsletter content containers. It strips page controls,
scripts, styles, and HTML, retaining plain-text paragraphs, lists, and HTTP(S)
link destinations. Client-rendered, private, blocked, or changed layouts may fail
or need parser adjustments; verify against your actual issue before enabling a
trigger. A successful fetch is not proof that every announcement was extracted.

Failures appear in **Executions** and preserve the last successful article.
No readable body means no new public row, rather than publishing a link card.
If extraction is unsupported, you can manually paste the six fields (including
the complete plain-text body) into `smore_feed_public` to publish immediately.
Leave `published_at` blank if unknown; otherwise use an ISO date such as
`2026-10-26T12:00:00Z`. Cells starting with `=` should be entered as plain text
(prefix with `'`). Bodies above 45,000 characters require manual splitting into
paragraph-sized chunks in a future storage extension; the importer rejects them
instead of silently truncating. A later successful sync replaces manual edits
for the same issue. Old link-only public rows must be migrated, not retained.

### 3. Deploy (or redeploy) the web app

1. **Deploy → Manage deployments → Edit** your existing web app.
   For a first deployment: **Deploy → New deployment → Web app**.
2. Choose **New version**, **Execute as: Me**, **Who has access: Anyone**.
   Only the public feed is returned by `doGet`; never expose submission rows.
3. Click **Deploy**. Updating the existing deployment preserves this URL:

```text
https://script.google.com/macros/s/AKfycbytRb0-UXs1XOt1XgETiVMJTBogJ3nCa3EZ-B3ZBQeAEBfkFBwZa9fMloGy-FvGOeNH/exec
```

Use the `/exec` web app URL, **not** the library URL or editor-only `/dev` URL.
If you created a different deployment URL, update `SMORE_FEED_API_URL` in
`index.html` (and the submission endpoint if they share the deployment).

### 4. Add the automatic trigger

In Apps Script **Triggers → Add Trigger**:

```text
Function: syncSmoreFeed
Deployment: Head
Event source: Time-driven
Type: Hour timer
Interval: Every 6 hours
```

Delete old newsletter-sync triggers to avoid duplicate writers. Do not delete
unrelated submission triggers. Larger archives may hit Apps Script execution or
URL-fetch quotas; check failed executions and reduce the refresh frequency.

### JSON contract / browser access

`GET /exec` must return JSON (not a login page), following Google's redirect:

```json
{
  "ok": true,
  "items": [{
    "issue_url": "https://www.smore.com/n/abcde-example",
    "issue_title": "Weekly school newsletter",
    "published_at": "2026-10-26T12:00:00Z",
    "summary": "This week's announcements.",
    "cover_image_url": "https://example.org/newsletter.jpg",
    "article_body_text": "First announcement.\n\nSecond announcement with full details."
  }]
}
```

`items` may be empty. All fields are strings; `published_at` and
`cover_image_url` may be empty. `article_body_text` is the full plain-text issue,
not its summary. Optional `article_body_html` is **ignored** by this client:
remote content is rendered as text, never executable HTML. Title and URL are
required. Legacy title/URL/summary-only rows remain readable as a clearly labeled
fallback but cannot provide body search until migrated.

The site uses a simple GET with no custom headers or credentials, avoiding a
CORS preflight. Apps Script must be publicly deployed; ContentService serves
JSON through Google's redirected content URL. Do not use `mode: 'no-cors'`
(it makes the response unreadable). If your Workspace disallows anonymous web
apps, an administrator must allow them or provide a same-origin JSON endpoint.
Browser JavaScript does not scrape Smore directly.

### Validation checklist

- [ ] Open `/exec` in an incognito window; confirm `ok: true`, one item per
      issue, and a complete `article_body_text`. Check no private sheet data leaks.
- [ ] Run sync twice; no duplicate issue rows. Edit a source issue and sync:
      the existing row updates. A failed fetch preserves the previous body.
- [ ] On the site, menu → **Flight Wire (Smore Board)** → **Read newsletter**.
      Verify paragraphs, title, date, cover, and **Open original Smore**.
- [ ] Search in **Instant** mode for a distinctive phrase near the *end* of the
      body that isn't in the title/summary. Select the result: it opens the issue.
- [ ] Close the reader; open a normal article. Classroom callouts/video controls,
      navigation, bookmarks, and lesson submissions still behave as before.
- [ ] If the feed is unavailable/empty, Flight Wire displays a clear message.
