# Eagle Deep Dive Daily

A starting point for the Eagle Deep Dive Daily website. The project currently
uses plain HTML and has no build tools or runtime dependencies.

## Preview

Open `index.html` in a web browser. Add the site's content and assets there as
they are ready.

## Unified Google Sheet CMS + Smore Sync Setup

The site is powered by a unified Google Apps Script (`apps-script/Smore.gs`) bound to your Google Sheet. It lets you edit **all site text, headings, articles, quick wins, directory links, staff lounge quotes, peer lesson studio cards, and EdTech tools** directly in your Google Sheet, while also syncing Smore newsletters and receiving form submissions (`Share a Tip`, `Post a Quote`, and `Peer Lesson Studio`) into `submissions` for review.

### 1. Paste the Unified Apps Script (`apps-script/Smore.gs`)

1. Open your Google Sheet and click **Extensions → Apps Script**.
2. Replace the older script code with the full contents of [`apps-script/Smore.gs`](apps-script/Smore.gs).
   *(Important: Google Apps Script only allows one `doGet` and one `doPost` in the project. `apps-script/Smore.gs` already includes the unified `doGet`, `doPost`, `syncSmoreFeed`, `setupSiteSheet`, and `publishApprovedSubmissions`.)*
3. Click **Save**.

### 2. Run `setupSiteSheet` Once to Auto-Create & Seed All Tabs

1. In the Apps Script toolbar dropdown at the top, select **`setupSiteSheet`** (or **`setupSmoreFeed`**) and click **Run** (or reload your Google Sheet and use the new **Eagle Site CMS → 1. Setup / Seed Site Tabs & Defaults** menu).
2. Authorize spreadsheet and external-fetch access when prompted.
3. The script will non-destructively ensure all required tabs and headers exist (matching columns by name and never clearing your existing rows) and will pre-populate any empty content tabs with the current site defaults so you can edit or delete them right away:

| Tab Name | Purpose | Required / Created Columns |
| --- | --- | --- |
| `settings` | Change any site heading, subtitle, badge, cover story ID, or video embed URL | `key`, `value`, `description` |
| `published_articles` | All main feed articles & playbooks (set `is_live` to `TRUE` or `FALSE`) | `article_id`, `title`, `category`, `summary`, `article_body_text`, `author`, `published_at`, `cover_image_url`, `is_live`, `takeaway`, `try_tomorrow`, `script`, `read_time`, `featured`, `has_video_demo`, `tags` |
| `quick_wins` | Context-aware Quick Wins banner cards per section (`all`, `Student Engagement`, `Classroom Management`, `Teacher Support`, `Tech & AI`, `saved`, `view-directory`, `view-bulletin`, `view-studio`, `view-academy`) | `tip_id`, `section`, `category_label`, `title`, `tip`, `is_live` |
| `directory_links` | Heart of MMS Building Directory cards | `resource_id`, `title`, `category`, `description`, `url`, `tags`, `is_live` |
| `lounge_quotes` | "Overheard at MMS" Staff Lounge quotes in the sidebar | `quote_id`, `quote`, `attribution`, `is_live` |
| `peer_studio` | Published Peer Lesson Studio cards | `studio_id`, `title`, `subject`, `dilemma`, `warm_feedback`, `cool_feedback`, `doc_url`, `is_live` |
| `tech_tools` | EdTech, AI & Classroom Video Lab software cards | `tool_id`, `badge`, `title`, `link_label`, `url`, `description`, `teacher_workflow`, `student_workflow`, `is_live` |
| `smore_feed_raw` / `smore_feed_public` | Automated Smore Flight Wire newsletter archive | `issue_url`, `issue_title`, `published_at`, `summary`, `cover_image_url`, `article_body_text` |
| `submissions` / `review_queue` | Incoming staff form submissions (`status = pending`) | `id`, `submitted_at`, `submitter_name`, `submitter_email`, `submission_type`, `title`, `content`, `links`, `status`, `category`, `reviewed_by`, `reviewed_at` |
| `audit_log` | Automatic log of submissions, approvals, and Smore syncs | `timestamp`, `action`, `source_tab`, `item_id`, `title`, `details` |

*(Note: If your sheet still has the tab spelled `puplished_articles`, the script automatically detects and works with both `published_articles` and `puplished_articles`.)*

### 3. Approving Submissions with One Click

When teachers submit a form on the site (**Share a Tip**, **Post a Classroom Quote**, or **Post a Lesson Draft** in Peer Lesson Studio):
1. A new row is added to `submissions` with `status = pending`.
2. To publish it to the live site, change its `status` cell in `submissions` (or `review_queue`) to **`approved`**.
3. In your Google Sheet top menu bar, click **Eagle Site CMS → 2. Publish Approved Submissions**.
4. The script automatically copies the item into the matching live tab (`published_articles`, `lounge_quotes`, or `peer_studio` with `is_live = TRUE`), updates the submission status to `published`, and records an entry in `audit_log`.

**Headers and migration.** Columns are matched **by name**, so order does not
matter and extra columns are kept. For example, the older layout
`id	issue_url	issue_title	published_at	summary	cover_image_url	is_live	last_synced_at`
works as-is. If a required header is missing, setup/sync writes it into the empty
cell(s) right of the last header; existing cells, rows, and columns are never
moved or cleared. `doGet` does not add headers; it returns `ok: false` naming the
missing header until `setupSmoreFeed` / `setupSiteSheet` is run. Optional `last_synced_at` is filled
on each successful refresh. Optional `is_live`: `FALSE` hides an issue; blank or
`TRUE` shows it.

`settings` uses column A for keys and column B for values (a `key	value	description`
header is written into an empty tab). Setup adds all default site text keys and
this row unless a `smore_source_url` (or legacy `source_url`) row already exists:

```text
smore_source_url	https://www.smore.com/u/maizemiddleschool
```

**Accepted source/issue URL formats.** Hosts must be `www.smore.com`,
`smore.com`, or `app.smore.com`.

| Use | Format | Example |
| --- | --- | --- |
| Profile (discovers its issues) | `/u/<profile>` | `https://www.smore.com/u/maizemiddleschool` |
| Single issue | `/n/<id>` or `/n/<id>-<slug>` | `https://app.smore.com/n/cp9a2` |
| Single issue (older style) | `/<id>-<slug>` or `/<id>` | `https://www.smore.com/abc12-fall-news` |

Issue URLs are normalized to `https://www.smore.com/<path>` with query
strings, fragments, and trailing slashes removed. This normalized URL is the
dedupe key. If `www` does not serve an `/n/` issue, it is fetched from
`app.smore.com`. Rejected: `cdn.smore.com` (images/files), `out.smore.com`
(tracking redirects), profile/utility pages (for example `/pricing`, `/login`,
`/educators`), asset paths with file extensions, and non-Smore hosts.
When the source is a single issue, links **inside** it are never treated as issues.

### 2. Import and check the content

1. Run `syncSmoreFeed` (function dropdown → **Run**). It discovers **issue URLs
   only**, fetches each issue, and updates the matching row by normalized URL,
   appending a row for new issues. Links *inside* a newsletter stay in its body;
   they do not become separate issues.
2. Open **Executions** and inspect the sync log (`refreshed` lists issue URLs;
   all must be `https://www.smore.com/...` issue pages). Compare a
   `smore_feed_public` row with the original newsletter. Check its non-empty
   `article_body_text`, headings, lists, dates, links, and the **end** of the
   article, not just the summary.
3. To import older issues not linked by the profile, paste their actual issue URLs
   into the `issue_url` column of `smore_feed_raw`, one URL per row, and sync again. Do not paste
   resource links or the profile URL into these rows.

**Source limitations:** live Smore markup could not be fetched in the development
environment. The importer supports JSON-LD `articleBody`, or server-rendered
`article`/`main`/recognized newsletter content containers, falling back to the
page `<body>`. It strips page controls (nav, header, footer, buttons, forms,
icon ligatures such as `zoom_out_map`), scripts, styles, and HTML, retaining
plain-text paragraphs, lists, and link destinations for text links. Image-only
links and `cdn`/`out.smore.com` URLs are left out. Bodies under 200 characters
are treated as a parse failure. Client-rendered, private, blocked, or changed layouts may fail
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
for the same issue.

### Troubleshooting: bad links in Flight Wire

Symptoms: `/exec` items have `issue_url` values like `https://cdn.smore.com/...jpeg`
or `https://out.smore.com/render/...`, titles like `zoom_out_map`, or no
`article_body_text`.

1. **An old script is still deployed.** If `/exec` lacks `article_body_text`,
   the live deployment runs the old sync/`doGet`, which scraped every link on
   the source page. Remove the old sync functions (`extractSmoreItems_`,
   `ensurePublicHeader_`, `ensureRawHeader_`, etc.) and old `doGet`, paste
   `apps-script/Smore.gs`, keep `doPost`, and redeploy a **New version** (step 3).
   Delete old sync triggers.
2. **Leftover legacy rows.** Rows whose `issue_url` is not an accepted issue
   URL, or that lack a title or `article_body_text`, are never returned by
   `doGet` and are ignored by sync. They are not deleted automatically; once the
   new rows look right, you may delete them by hand.
3. **Wrong source.** Use a profile or issue URL from the table above, not an
   image, redirect, or app editor link.
4. **Empty body.** The issue may be private, client-rendered, or changed layout;
   see Executions. Previous content is kept; you can paste plain text manually.

The site also ignores feed rows whose URL is not on `smore.com`,
`www.smore.com`, or `app.smore.com`.

### 3. Deploy (or redeploy) the web app

1. **Deploy → Manage deployments → Edit** your existing web app.
   For a first deployment: **Deploy → New deployment → Web app**.
2. Choose **New version**, **Execute as: Me**, **Who has access: Anyone**.
   Only the public feed is returned by `doGet`; never expose submission rows.
3. Click **Deploy**. Updating the existing deployment preserves this URL:

```text
https://script.google.com/macros/s/AKfycbyAxb9nhP6H34pWi5icBPtWLoNSenmHalzwbkQr81p5WScUl1ez1Ym9RoglEk2RWWQb/exec
```

Use the currently active `/exec` web app URL, **not** the library URL or
editor-only `/dev` URL. When the feed and submissions share this deployment,
`SMORE_FEED_API_URL` in `index.html` reuses `SUBMISSION_API_URL`. Published
articles also use that URL with `?feed=articles`. If you create a different
deployment, update `SUBMISSION_API_URL` only.

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
- [ ] Submit a lesson draft from **The Peer Lesson Studio**. Confirm the success
      message and verify a new row appears in the spreadsheet's `submissions`
      tab with `submission_type=lesson_draft` and `status=pending`.
- [ ] Load Flight Wire and confirm the Smore feed displays its published issues.

## Functional flow: private drafts → published articles

1. **Submit:** The Peer Lesson Studio sends a draft to the existing `doPost`,
   which must save it in `submissions` with `status=pending`. Success means
   **saved for review**, not published; no draft card is added to the page.
   Keep the spreadsheet private. The public GET routes never read `submissions`.
2. **Review and publish manually:** No reviewer UI is included yet. In your
   private spreadsheet, review the draft and copy only approved, public text
   into `published_articles`. Do not copy submitter emails or private draft links.
   Create the tab if needed, with these headers in row 1 (tab-separated):

   ```text
   article_id	title	category	summary	article_body_text	author	published_at	cover_image_url	is_live
   ```

   Headers are matched by name; order and extra private columns do not matter.
   Only these nine named fields can be returned publicly. Required fields are
   `article_id` (unique and stable), `title`, `category`, `article_body_text`
   (full plain text), and `is_live`. Other fields may be blank. Use one of the
   existing categories: `Student Engagement`, `Classroom Management`,
   `Teacher Support`, or `Tech & AI`. Use an ISO date for `published_at` and an
   HTTP(S) image URL if supplied. Enter text beginning with `=` as plain text
   (prefix with `'`).
3. **Public display:** Set `is_live` to `TRUE` (a checked checkbox also works).
   On the next page load, the site fetches `/exec?feed=articles` and includes
   valid live rows in article listings, the reader, bookmarks, and Instant search,
   including full-body search. False/blank live flags and incomplete rows are
   excluded. Set `is_live=FALSE` and refresh to unpublish. Existing editorial
   articles and Flight Wire newsletters remain unchanged; submissions never
   supply public content. An absent `published_articles` tab returns an empty
   feed. There is no automatic draft-to-article copy or reviewer portal yet.

**Before testing:** Install the updated `apps-script/Smore.gs`, preserving your
existing `doPost` and its pending-row write, then redeploy a **New version of the
existing deployment** as described above. Publish the updated `index.html` too.
Spreadsheet edits after this setup do not require another deployment.

### Manual test checklist — complete before the next phase

- [ ] **A. Submit a draft:** In The Peer Lesson Studio, enter a distinctive title,
      department/grade, feedback focus, and valid draft link. Click **Submit for
      Review**. Verify a new `submissions` row has `submission_type=lesson_draft`
      and `status=pending`; confirm the page says it was saved for review.
- [ ] **B. Confirm privacy:** No immediate public-looking draft card is added.
      Refresh and search for the draft title: it must not appear.
- [ ] **C. Publish one test row:** In `published_articles`, enter
      `article_id=test-review-1`, `title=Review workflow test`,
      `category=Teacher Support`, `summary=Approved test article`,
      `article_body_text=This approved article contains the phrase silver eagle experiment.`,
      and `is_live=TRUE`. Leave optional fields blank if desired. Open the active
      deployment URL with `?feed=articles` in an incognito window; verify only
      live articles and the documented public fields are returned, not submissions
      or extra private columns.
- [ ] **D. Verify display and search:** Refresh the site. Find the article under
      Latest Articles and Teacher Support; open **Read Full Playbook** and check
      its complete body. In **Instant** search, find its title and then
      `silver eagle experiment`; select the result and verify the same reader.
      Bookmark it and verify it appears in Saved.
- [ ] **E. Unpublish:** Set that row's `is_live=FALSE`, refresh the site, and
      confirm it disappears from listings, Saved, and search. Also check a blank
      live flag does not publish it. Already-open pages must be refreshed.
- [ ] **Regression:** Open Flight Wire, read a newsletter, and search for a phrase
      in its full body. Existing editorial articles and navigation still work.
      A failed lesson submission should show an error and retain the form values.
