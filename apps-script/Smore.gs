/***************************************************************
 * Eagle Deep Dive Daily — Unified Google Sheet CMS & Smore Sync
 *
 * Replace all older script code in your bound Apps Script project
 * with this file (there must be only one `doGet` and one `doPost`).
 *
 * Run `setupSiteSheet` (or `setupSmoreFeed`) once from the toolbar
 * to create/update all tabs, headers, and default site content.
 ***************************************************************/

const SMORE_COLUMNS = [
  'issue_url', 'issue_title', 'published_at', 'summary',
  'cover_image_url', 'article_body_text'
];

// Utility pages on smore.com that look like issue slugs but are never newsletters.
const SMORE_RESERVED_PATHS = /^(?:about|account|app|blog|business|careers|contact|dashboard|discover|educators|explore|features|gallery|help|jobs|legal|library|login|logout|newsletters|press|pricing|privacy|schools|search|settings|signup|support|templates|terms)$/i;

const PUBLISHED_ARTICLE_COLUMNS = [
  'article_id', 'title', 'category', 'summary', 'article_body_text',
  'author', 'published_at', 'cover_image_url', 'is_live',
  'takeaway', 'try_tomorrow', 'script', 'read_time', 'featured', 'has_video_demo', 'tags'
];

const QUICK_WINS_COLUMNS = [
  'tip_id', 'section', 'category_label', 'title', 'tip', 'is_live'
];

const DIRECTORY_LINK_COLUMNS = [
  'resource_id', 'title', 'category', 'description', 'url', 'tags', 'is_live'
];

const LOUNGE_QUOTE_COLUMNS = [
  'quote_id', 'quote', 'attribution', 'is_live'
];

const PEER_STUDIO_COLUMNS = [
  'studio_id', 'title', 'subject', 'dilemma', 'warm_feedback', 'cool_feedback', 'doc_url', 'is_live'
];

const TECH_TOOL_COLUMNS = [
  'tool_id', 'badge', 'title', 'link_label', 'url', 'description', 'teacher_workflow', 'student_workflow', 'is_live'
];

const SUBMISSION_COLUMNS = [
  'id', 'submitted_at', 'submitter_name', 'submitter_email', 'submission_type',
  'title', 'content', 'links', 'status', 'category', 'reviewed_by', 'reviewed_at'
];

const AUDIT_LOG_COLUMNS = [
  'timestamp', 'action', 'source_tab', 'item_id', 'title', 'details'
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Eagle Site CMS')
    .addItem('1. Setup / Seed Site Tabs & Defaults', 'setupSiteSheet')
    .addItem('2. Publish Approved Submissions', 'publishApprovedSubmissions')
    .addSeparator()
    .addItem('3. Sync Smore Newsletters Now', 'syncSmoreFeed')
    .addToUi();
}

function setupSmoreFeed() {
  setupSiteSheet();
}

function setupSiteSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ['smore_feed_raw', 'smore_feed_public'].forEach(name => {
    smoreHeaders_(ss.getSheetByName(name) || ss.insertSheet(name), SMORE_COLUMNS, true);
  });

  const settings = ss.getSheetByName('settings') || ss.insertSheet('settings');
  if (!settings.getLastRow()) {
    settings.getRange(1, 1, 1, 3).setValues([['key', 'value', 'description']]);
  } else {
    smoreHeaders_(settings, ['key', 'value'], true);
  }
  seedDefaultSettings_(settings);

  const articlesSheet = getPublishedArticlesSheet_(ss) || ss.insertSheet('published_articles');
  const articleHeaders = smoreHeaders_(articlesSheet, PUBLISHED_ARTICLE_COLUMNS, true);
  if (articlesSheet.getLastRow() < 2) {
    seedSheetRows_(articlesSheet, articleHeaders, defaultArticlesSeed_());
  }

  const quickWinsSheet = ss.getSheetByName('quick_wins') || ss.insertSheet('quick_wins');
  const quickWinsHeaders = smoreHeaders_(quickWinsSheet, QUICK_WINS_COLUMNS, true);
  if (quickWinsSheet.getLastRow() < 2) {
    seedSheetRows_(quickWinsSheet, quickWinsHeaders, defaultQuickWinsSeed_());
  }

  const directorySheet = ss.getSheetByName('directory_links') || ss.insertSheet('directory_links');
  const directoryHeaders = smoreHeaders_(directorySheet, DIRECTORY_LINK_COLUMNS, true);
  if (directorySheet.getLastRow() < 2) {
    seedSheetRows_(directorySheet, directoryHeaders, defaultDirectorySeed_());
  }

  const quotesSheet = ss.getSheetByName('lounge_quotes') || ss.insertSheet('lounge_quotes');
  const quotesHeaders = smoreHeaders_(quotesSheet, LOUNGE_QUOTE_COLUMNS, true);
  if (quotesSheet.getLastRow() < 2) {
    seedSheetRows_(quotesSheet, quotesHeaders, defaultQuotesSeed_());
  }

  const studioSheet = ss.getSheetByName('peer_studio') || ss.insertSheet('peer_studio');
  const studioHeaders = smoreHeaders_(studioSheet, PEER_STUDIO_COLUMNS, true);
  if (studioSheet.getLastRow() < 2) {
    seedSheetRows_(studioSheet, studioHeaders, defaultStudioSeed_());
  }

  const techSheet = ss.getSheetByName('tech_tools') || ss.insertSheet('tech_tools');
  const techHeaders = smoreHeaders_(techSheet, TECH_TOOL_COLUMNS, true);
  if (techSheet.getLastRow() < 2) {
    seedSheetRows_(techSheet, techHeaders, defaultTechToolsSeed_());
  }

  const submissionsSheet = ss.getSheetByName('submissions') || ss.insertSheet('submissions');
  smoreHeaders_(submissionsSheet, SUBMISSION_COLUMNS, true);

  const reviewQueueSheet = ss.getSheetByName('review_queue') || ss.insertSheet('review_queue');
  smoreHeaders_(reviewQueueSheet, SUBMISSION_COLUMNS, true);

  const auditSheet = ss.getSheetByName('audit_log') || ss.insertSheet('audit_log');
  smoreHeaders_(auditSheet, AUDIT_LOG_COLUMNS, true);
}

function getPublishedArticlesSheet_(ss) {
  return ss.getSheetByName('published_articles') || ss.getSheetByName('puplished_articles');
}

function smoreSourceUrl_(settings) {
  if (!settings || settings.getLastRow() < 1) return '';
  const rows = settings.getDataRange().getValues();
  const row = rows.find(r => String(r[0]).trim() === 'smore_source_url') ||
    rows.find(r => String(r[0]).trim() === 'source_url');
  return String(row && row[1] || '').trim();
}

function syncSmoreFeed() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another newsletter sync is running');
  try {
    setupSiteSheet();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const raw = ss.getSheetByName('smore_feed_raw');
    const published = ss.getSheetByName('smore_feed_public');
    const rawHeaders = smoreHeaders_(raw, SMORE_COLUMNS, true);
    const publicHeaders = smoreHeaders_(published, SMORE_COLUMNS, true);
    const urls = new Set();
    smoreReadIssues_(published, publicHeaders).concat(smoreReadIssues_(raw, rawHeaders)).forEach(row => {
      const url = smoreIssueUrl_(row.issue.issue_url);
      if (url) urls.add(url);
    });
    try {
      smoreDiscover_(smoreSourceUrl_(ss.getSheetByName('settings'))).forEach(url => urls.add(url));
    } catch (err) {
      console.warn('Discovery failed; retrying known issue URLs. ' + err.message);
    }
    if (!urls.size) throw new Error('No issue URLs found. Add a public issue URL to the issue_url column of smore_feed_raw.');
    const now = new Date().toISOString();
    const refreshed = [];
    const pending = [];
    const failures = [];
    for (const url of urls) {
      try {
        refreshed.push(Object.assign(smoreParseIssue_(smoreFetchIssue_(url), url), { last_synced_at: now }));
      } catch (err) {
        // Only the URL is recorded; the last successful article is never replaced.
        pending.push({ issue_url: url });
        failures.push(url + ': ' + err.message);
      }
    }
    smoreWriteIssues_(raw, rawHeaders, refreshed.concat(pending));
    smoreWriteIssues_(published, publicHeaders, refreshed);
    appendAuditLog_(ss, 'smore_sync', 'smore_feed_public', 'batch', 'Smore Sync', 'Refreshed: ' + refreshed.length + ', Failures: ' + failures.length);
    console.log(JSON.stringify({ refreshed: refreshed.map(issue => issue.issue_url), failures }));
    if (failures.length) throw new Error('Some issues could not refresh; previous content retained. See execution log.');
  } finally {
    lock.releaseLock();
  }
}

function smoreDiscover_(sourceUrl) {
  // A single newsletter is used as-is: links inside it are content, not issues.
  const issue = smoreIssueUrl_(sourceUrl);
  if (issue) return [issue];
  if (!/^https:\/\/(?:www\.|app\.)?smore\.com\/u\/[a-z0-9_-]+\/?$/i.test(sourceUrl)) {
    throw new Error('smore_source_url must be a public Smore profile or issue URL');
  }
  const found = new Set();
  for (const match of smoreFetch_(sourceUrl).matchAll(/\bhref\s*=\s*["']([^"']+)["']/gi)) {
    const link = smoreDecode_(match[1]).trim();
    const url = smoreIssueUrl_(/^\/(?!\/)/.test(link) ? 'https://www.smore.com' + link : link);
    if (url) found.add(url);
  }
  return [...found];
}

// Single unified doGet handler for all feeds.
function doGet(e) {
  try {
    const feed = e && e.parameter && String(e.parameter.feed || '').trim().toLowerCase();
    if (feed === 'articles') {
      return smoreJson_({ ok: true, items: publishedArticles_() });
    }
    if (feed === 'site' || feed === 'all') {
      return smoreJson_({
        ok: true,
        settings: publicSettings_(),
        articles: publishedArticles_(),
        smore: publicSmoreItems_(false),
        quick_wins: publicTableItems_('quick_wins', QUICK_WINS_COLUMNS, ['title', 'tip']),
        directory: publicTableItems_('directory_links', DIRECTORY_LINK_COLUMNS, ['title', 'url']),
        quotes: publicTableItems_('lounge_quotes', LOUNGE_QUOTE_COLUMNS, ['quote']),
        studio: publicTableItems_('peer_studio', PEER_STUDIO_COLUMNS, ['title']),
        tech_tools: publicTableItems_('tech_tools', TECH_TOOL_COLUMNS, ['title'])
      });
    }
    return smoreJson_({ ok: true, items: publicSmoreItems_(true) });
  } catch (err) {
    return smoreJson_({ ok: false, error: String(err.message) });
  }
}

function publicSmoreItems_(strictSheet) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('smore_feed_public');
  if (!sheet) {
    if (strictSheet) throw new Error('Run setupSmoreFeed first');
    return [];
  }
  const seen = new Set();
  const headers = strictSheet
    ? smoreHeaders_(sheet, SMORE_COLUMNS, false)
    : sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim());
  if (!strictSheet && SMORE_COLUMNS.some(col => !headers.includes(col))) return [];
  return smoreReadIssues_(sheet, headers).map(row => row.issue)
    .filter(issue => {
      const url = smoreIssueUrl_(issue.issue_url);
      if (!url || seen.has(url) || !issue.issue_title || !issue.article_body_text ||
        String(issue.is_live || '').toLowerCase() === 'false') return false;
      seen.add(url);
      issue.issue_url = url;
      return true;
    })
    .map(issue => Object.fromEntries(SMORE_COLUMNS.map(key => [key, issue[key] || ''])));
}

function publishedArticles_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getPublishedArticlesSheet_(ss);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const [rawHeaders, ...rows] = sheet.getDataRange().getValues();
  const headers = rawHeaders.map(h => String(h || '').trim());
  const required = ['article_id', 'title', 'category', 'article_body_text', 'is_live'];
  if (!headers.length || required.some(key => !headers.includes(key))) {
    throw new Error('Missing required published_articles headers: ' + required.join(', '));
  }
  return rows.filter(row => {
    const val = row[headers.indexOf('is_live')];
    return val === true || String(val || '').trim().toLowerCase() === 'true';
  }).map(row => Object.fromEntries(PUBLISHED_ARTICLE_COLUMNS.map(key => {
    const idx = headers.indexOf(key);
    const value = idx >= 0 ? row[idx] : '';
    if (key === 'is_live') return [key, true];
    if (key === 'featured' || key === 'has_video_demo') {
      return [key, value === true || String(value || '').trim().toLowerCase() === 'true'];
    }
    return [key, value instanceof Date ? value.toISOString() : String(value == null ? '' : value).trim()];
  }))).filter(item => required.every(key => item[key]));
}

function publicSettings_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('settings');
  if (!sheet || sheet.getLastRow() < 1) return {};
  const values = sheet.getDataRange().getValues();
  const out = {};
  const startRow = String(values[0][0] || '').trim().toLowerCase() === 'key' ? 1 : 0;
  for (let r = startRow; r < values.length; r++) {
    const key = String(values[r][0] || '').trim();
    const rawVal = values[r][1];
    if (!key) continue;
    const val = rawVal instanceof Date ? rawVal.toISOString() : String(rawVal == null ? '' : rawVal).trim();
    out[key] = val;
  }
  return out;
}

function publicTableItems_(sheetName, columns, requiredKeys) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const [rawHeaders, ...rows] = sheet.getDataRange().getValues();
  const headers = rawHeaders.map(h => String(h || '').trim());
  if (!headers.length) return [];
  const liveIdx = headers.indexOf('is_live');
  return rows.filter(row => {
    if (liveIdx < 0) return true;
    const val = row[liveIdx];
    return val === true || String(val || '').trim().toLowerCase() === 'true';
  }).map(row => Object.fromEntries(columns.map(key => {
    const idx = headers.indexOf(key);
    const value = idx >= 0 ? row[idx] : '';
    if (key === 'is_live') return [key, true];
    return [key, value instanceof Date ? value.toISOString() : String(value == null ? '' : value).trim()];
  }))).filter(item => (requiredKeys || []).every(k => item[k]));
}

function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('submissions') || ss.insertSheet('submissions');
    const headers = smoreHeaders_(sheet, SUBMISSION_COLUMNS, true);

    let body = {};
    if (e && e.parameter && Object.keys(e.parameter).length > 0) {
      body = e.parameter;
    } else if (e && e.postData && e.postData.contents) {
      try {
        body = JSON.parse(e.postData.contents);
      } catch (err) {
        body = {};
      }
    }

    const submitterName = cleanSubmissionText_(body.submitter_name || '');
    const submitterEmail = cleanSubmissionText_(body.submitter_email || '');
    const submissionType = cleanSubmissionText_(body.submission_type || '');
    const title = cleanSubmissionText_(body.title || '');
    const content = String(body.content || '').replace(/<[^>]*>/g, ' ').replace(/\r\n?/g, '\n').trim();
    const links = cleanSubmissionText_(body.links || '');
    const category = cleanSubmissionText_(body.category || '');

    if (!submissionType || !title || !content) {
      return smoreJson_({ ok: false, error: 'Missing required fields: submission_type, title, content' });
    }

    const id = 'sub_' + Utilities.getUuid();
    const submittedAt = new Date().toISOString();

    const record = {
      id: id,
      submitted_at: submittedAt,
      submitter_name: submitterName,
      submitter_email: submitterEmail,
      submission_type: submissionType,
      title: title,
      content: content,
      links: links,
      status: 'pending',
      category: category,
      reviewed_by: '',
      reviewed_at: ''
    };

    appendRecordByHeaders_(sheet, headers, record);
    appendAuditLog_(ss, 'submission_received', 'submissions', id, title, 'Type: ' + submissionType);

    return smoreJson_({ ok: true, id: id });
  } catch (err) {
    return smoreJson_({ ok: false, error: String(err.message || err) });
  }
}

function cleanSubmissionText_(value) {
  return String(value || '')
    .replace(/<[^>]*>/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .trim();
}

function publishApprovedSubmissions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSiteSheet();
  const now = new Date().toISOString();
  let publishedCount = 0;

  ['submissions', 'review_queue'].forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet || sheet.getLastRow() < 2) return;
    const headers = smoreHeaders_(sheet, SUBMISSION_COLUMNS, true);
    const statusIdx = headers.indexOf('status');
    const reviewedAtIdx = headers.indexOf('reviewed_at');
    if (statusIdx < 0) return;

    const rows = smoreReadIssues_(sheet, headers);
    rows.forEach(r => {
      const status = String(r.issue.status || '').toLowerCase();
      if (status !== 'approved' && status !== 'publish') return;

      const type = String(r.issue.submission_type || 'article').toLowerCase();
      const id = r.issue.id || ('item_' + Utilities.getUuid().slice(0, 8));
      const title = r.issue.title || 'Staff Contribution';
      const content = r.issue.content || '';
      const author = r.issue.submitter_name || 'MMS Staff';
      const category = r.issue.category || 'Teacher Support';
      const links = r.issue.links || '';

      if (type === 'quote' || type === 'lounge_quote') {
        const qSheet = ss.getSheetByName('lounge_quotes');
        const qHeaders = smoreHeaders_(qSheet, LOUNGE_QUOTE_COLUMNS, true);
        appendRecordByHeaders_(qSheet, qHeaders, {
          quote_id: id,
          quote: content,
          attribution: author + (title ? ' (' + title + ')' : ''),
          is_live: 'TRUE'
        });
        appendAuditLog_(ss, 'published_quote', tabName, id, title, 'Published to lounge_quotes');
      } else if (type === 'lesson_draft' || type === 'studio') {
        const sSheet = ss.getSheetByName('peer_studio');
        const sHeaders = smoreHeaders_(sSheet, PEER_STUDIO_COLUMNS, true);
        appendRecordByHeaders_(sSheet, sHeaders, {
          studio_id: id,
          title: title,
          subject: author || 'MMS Classroom',
          dilemma: content,
          warm_feedback: 'Open the attached lesson link to share warm feedback with your colleague.',
          cool_feedback: 'What is one "I wonder if..." adjustment that could make this lesson even smoother?',
          doc_url: links,
          is_live: 'TRUE'
        });
        appendAuditLog_(ss, 'published_studio', tabName, id, title, 'Published to peer_studio');
      } else {
        const aSheet = getPublishedArticlesSheet_(ss);
        const aHeaders = smoreHeaders_(aSheet, PUBLISHED_ARTICLE_COLUMNS, true);
        appendRecordByHeaders_(aSheet, aHeaders, {
          article_id: id,
          title: title,
          category: category,
          summary: content.slice(0, 220),
          article_body_text: content,
          author: author,
          published_at: now,
          cover_image_url: '',
          is_live: 'TRUE',
          takeaway: content.slice(0, 220),
          try_tomorrow: content,
          script: '',
          read_time: '2 min read',
          featured: 'FALSE',
          has_video_demo: 'FALSE',
          tags: category.toLowerCase() + ', staff tip'
        });
        appendAuditLog_(ss, 'published_article', tabName, id, title, 'Published to published_articles');
      }

      sheet.getRange(r.row, statusIdx + 1).setValue('published');
      if (reviewedAtIdx >= 0) {
        sheet.getRange(r.row, reviewedAtIdx + 1).setValue(now);
      }
      publishedCount++;
    });
  });

  try {
    SpreadsheetApp.getUi().alert('Published ' + publishedCount + ' approved submission(s).');
  } catch (err) {
    console.log('Published ' + publishedCount + ' approved submission(s).');
  }
}

function appendRecordByHeaders_(sheet, headers, record) {
  const row = headers.map(key => {
    const val = record[key];
    if (val === undefined || val === null) return '';
    const text = String(val);
    return /^[=+@-]/.test(text) ? "'" + text : text;
  });
  sheet.appendRow(row);
}

function seedSheetRows_(sheet, headers, records) {
  if (!records || !records.length) return;
  const rows = records.map(rec => headers.map(key => {
    const val = rec[key];
    if (val === undefined || val === null) return '';
    const text = String(val);
    return /^[=+@-]/.test(text) ? "'" + text : text;
  }));
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function appendAuditLog_(ss, action, sourceTab, itemId, title, details) {
  try {
    const sheet = ss.getSheetByName('audit_log');
    if (!sheet) return;
    const headers = smoreHeaders_(sheet, AUDIT_LOG_COLUMNS, true);
    appendRecordByHeaders_(sheet, headers, {
      timestamp: new Date().toISOString(),
      action: action,
      source_tab: sourceTab,
      item_id: itemId,
      title: title,
      details: details
    });
  } catch (err) {
    // Ignore audit log errors so primary flow is unaffected
  }
}

function smoreJson_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function smoreHeaders_(sheet, columns, addMissing) {
  // Columns are matched by name. Extra columns (for example id, is_live, last_synced_at)
  // are kept; missing ones are only added to empty cells right of the header row.
  if (!sheet.getLastRow() && addMissing) sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  const width = Math.max(sheet.getLastColumn(), 1);
  const headers = sheet.getRange(1, 1, 1, width).getValues()[0].map(h => String(h).trim());
  const missing = columns.filter(column => !headers.includes(column));
  if (missing.length) {
    if (!addMissing) throw new Error('Missing headers in ' + sheet.getName() + ': ' + missing.join(', ') + '. Run setupSmoreFeed.');
    sheet.getRange(1, width + 1, 1, missing.length).setValues([missing]);
    headers.push(...missing);
  }
  return headers;
}

function smoreReadIssues_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues().map((values, i) => ({
    row: i + 2,
    values,
    issue: Object.fromEntries(headers.map((key, c) => [
      key, values[c] instanceof Date ? values[c].toISOString() : String(values[c] == null ? '' : values[c]).trim()
    ]))
  }));
}

function smoreWriteIssues_(sheet, headers, issues) {
  // Upsert by canonical issue URL. Other rows and unrelated columns are never cleared.
  const rows = smoreReadIssues_(sheet, headers);
  let nextRow = Math.max(sheet.getLastRow(), 1) + 1;
  issues.forEach(issue => {
    let matches = rows.filter(row => smoreIssueUrl_(row.issue.issue_url) === issue.issue_url);
    if (!matches.length) {
      matches = [{ row: nextRow++, values: headers.map(() => ''), issue: { issue_url: issue.issue_url } }];
      rows.push(matches[0]);
    }
    matches.forEach(row => {
      Object.keys(issue).forEach(key => {
        const c = headers.indexOf(key);
        if (c < 0) return;
        const text = String(issue[key] || '');
        row.values[c] = /^[=+@-]/.test(text) ? "'" + text : text;
      });
      sheet.getRange(row.row, 1, 1, headers.length).setValues([row.values]);
    });
  });
}

function smoreIssueUrl_(value) {
  // Only newsletter pages: rejects cdn.smore.com, out.smore.com, profiles, assets and site pages.
  const match = String(value || '').trim().match(
    /^https?:\/\/(?:www\.|app\.)?smore\.com\/((?:n\/)?[a-z0-9]{5,}(?:-[a-z0-9-]+)?)\/?(?:[?#].*)?$/i
  );
  if (!match || SMORE_RESERVED_PATHS.test(match[1].split('-')[0])) return '';
  return 'https://www.smore.com/' + match[1];
}

function smoreFetchIssue_(url) {
  try {
    return smoreFetch_(url);
  } catch (err) {
    // Newer issues may only be served from app.smore.com.
    if (!/^https:\/\/www\.smore\.com\/n\//.test(url)) throw err;
    return smoreFetch_(url.replace('://www.', '://app.'));
  }
}

function smoreFetch_(url) {
  // Never follow a remote redirect off Smore (or into a private network).
  for (let i = 0; i < 4; i++) {
    if (!/^https:\/\/(?:www\.|app\.)?smore\.com\//i.test(url)) throw new Error('Only HTTPS Smore page URLs are allowed');
    const response = UrlFetchApp.fetch(url, { followRedirects: false, muteHttpExceptions: true });
    const code = response.getResponseCode();
    if (code === 200) return response.getContentText();
    if ([301, 302, 303, 307, 308].includes(code)) {
      const headers = response.getAllHeaders();
      const locationKey = Object.keys(headers).find(key => key.toLowerCase() === 'location');
      const location = String(headers[locationKey] || '');
      url = location.startsWith('/') && !location.startsWith('//') ? 'https://www.smore.com' + location : location;
    } else {
      throw new Error('Smore returned HTTP ' + code);
    }
  }
  throw new Error('Too many redirects');
}

function smoreParseIssue_(html, url) {
  let structured = null;
  function visit(value) {
    if (!value || typeof value !== 'object') return;
    if (typeof value.articleBody === 'string') structured = value;
    Object.values(value).forEach(child => {
      if (child && typeof child === 'object') visit(child);
    });
  }
  for (const match of html.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script\s*>/gi)) {
    if (!/application\/ld\+json/i.test(match[1])) continue;
    try { visit(JSON.parse(match[2])); } catch (err) { /* Ignore unrelated malformed metadata. */ }
  }
  const meta = {};
  for (const match of html.matchAll(/<meta\b[^>]*>/gi)) {
    const attrs = smoreAttributes_(match[0]);
    meta[attrs.property || attrs.name] = smoreDecode_(attrs.content || '');
  }
  const titleTag = html.match(/<title\b[^>]*>([\s\S]*?)<\/title\s*>/i);
  const title = smoreText_(structured && structured.headline || meta['og:title'] || (titleTag && titleTag[1]) || '');
  const body = structured && structured.articleBody
    ? smoreText_(structured.articleBody) : smoreText_(smoreContentRegion_(html));
  if (!title || body.length < 200) throw new Error('No readable issue body found; source may require JavaScript or authentication');
  if (body.length > 45000) throw new Error('Issue exceeds the 45,000 character sheet-cell limit; split it manually before publishing');
  const dateValue = structured && structured.datePublished || meta['article:published_time'] || meta.date || '';
  const date = new Date(dateValue);
  const image = structured && structured.image;
  const cover = typeof image === 'string' ? image : image && image.url || meta['og:image'] || '';
  return {
    issue_url: url,
    issue_title: title,
    published_at: isNaN(date.getTime()) ? '' : date.toISOString(),
    summary: smoreText_(structured && structured.description || meta.description || meta['og:description'] || body.slice(0, 280)),
    cover_image_url: /^https?:\/\//i.test(cover) ? cover : '',
    article_body_text: body
  };
}

function smoreContentRegion_(html) {
  // Match a bounded article region, not all profile links or the entire page.
  const tags = /<\/?([a-z][a-z0-9]*)\b[^>]*>/gi;
  const clean = html.replace(/<(script|style|noscript)\b[^>]*>[\s\S]*?<\/\1\s*>/gi, '');
  let match;
  while ((match = tags.exec(clean))) {
    const attrs = smoreAttributes_(match[0]);
    if (match[0][1] === '/' || !(
      /^(article|main)$/i.test(match[1]) ||
      /(?:^|\s)(?:flyer-content|flyer-body|newsletter-content|flyer-components)(?:\s|$)/i.test((attrs.id || '') + ' ' + (attrs.class || ''))
    )) continue;
    const start = tags.lastIndex;
    const tag = match[1].toLowerCase();
    let depth = 1;
    let end;
    while ((end = tags.exec(clean))) {
      if (end[1].toLowerCase() !== tag) continue;
      depth += end[0][1] === '/' ? -1 : 1;
      if (!depth) return clean.slice(start, end.index);
    }
  }
  // Last resort: the page body; nav/header/footer/buttons are stripped by smoreText_.
  const page = clean.match(/<body\b[^>]*>([\s\S]*)<\/body\s*>/i);
  if (page) return page[1];
  throw new Error('No article content region found; inspect source or manually publish plain text');
}

function smoreAttributes_(tag) {
  const attrs = {};
  for (const match of tag.matchAll(/([\w:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g)) {
    attrs[match[1].toLowerCase()] = match[2] === undefined ? (match[3] === undefined ? match[4] : match[3]) : match[2];
  }
  return attrs;
}

function smoreText_(html) {
  return smoreDecode_(String(html)
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/<(script|style|noscript|nav|header|footer|button|form|iframe|svg)\b[^>]*>[\s\S]*?<\/\1\s*>/gi, '')
    // Icon-font ligatures such as "zoom_out_map" are controls, not text.
    .replace(/<(i|span)\b[^>]*class\s*=\s*["'][^"']*material-(?:icons|symbols)[^"']*["'][^>]*>[^<]*<\/\1\s*>/gi, '')
    .replace(/<a\b([^>]*)>([\s\S]*?)<\/a\s*>/gi, (_, attrs, label) => {
      const href = smoreDecode_(smoreAttributes_(attrs).href || '').trim();
      const text = smoreDecode_(label.replace(/<[^>]*>/g, '')).trim();
      // Image-only links and Smore CDN/redirect URLs add noise, so keep only the label.
      const keep = text && text !== href && /^https?:\/\/[^\s<>"]+$/i.test(href) &&
        !/^https?:\/\/(?:cdn|out)\.smore\.com\//i.test(href);
      return label + (keep ? ' (' + href + ')' : '');
    })
    .replace(/<br\b[^>]*>/gi, '\n')
    .replace(/<\/(?:p|div|h[1-6]|section|article|ul|ol|blockquote)\s*>/gi, '\n\n')
    .replace(/<li\b[^>]*>/gi, '\n• ')
    .replace(/<[^>]*>/g, ''))
    .replace(/\r/g, '').replace(/[ \t\u00a0]+/g, ' ')
    .replace(/ *\n */g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

function smoreDecode_(text) {
  const entities = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', ndash: '–', mdash: '—', rsquo: '’', lsquo: '‘', rdquo: '”', ldquo: '“', bull: '•', hellip: '…' };
  return String(text).replace(/&(#x[0-9a-f]+|#\d+|[a-z]+);/gi, (whole, code) => {
    if (code[0] !== '#') return entities[code.toLowerCase()] || whole;
    const point = code[1].toLowerCase() === 'x' ? parseInt(code.slice(2), 16) : Number(code.slice(1));
    return point > 0 && point <= 0x10ffff ? String.fromCodePoint(point) : '';
  });
}

/* ========================= Default Site Seeds ========================= */

function seedDefaultSettings_(settingsSheet) {
  const existing = new Set(
    settingsSheet.getDataRange().getValues().map(r => String(r[0] || '').trim())
  );
  const defaults = [
    ['smore_source_url', 'https://www.smore.com/u/maizemiddleschool', 'Public Smore profile or issue URL to sync'],
    ['site_title', 'Eagle Deep Dive Daily', 'Main publication title in the top header and footer'],
    ['site_badge', 'MMS Eagles', 'Pill badge next to the main title'],
    ['site_subtitle', 'Maize Middle School • Classroom Craft, Student Engagement, Teacher Support & Tech', 'Subtitle under the main title'],
    ['site_tagline', 'Built by teachers, for teachers', 'Tagline on the right of the topic bar'],
    ['hero_article_id', 'art-1', 'article_id of the featured cover story on the home feed'],
    ['hero_badge', 'Cover Story • Student Engagement', 'Badge shown on the cover story image'],
    ['hero_button_label', 'Read Full Story & Script', 'Button text on the cover story card'],
    ['feed_heading_all', 'Latest Articles & Classroom Playbooks', 'Main heading above the articles grid'],
    ['coach_heading', 'Instant Strategy Brainstorm', 'Sidebar AI Coach widget title'],
    ['coach_subtitle', 'Powered by Gemini Flash • Grounded in MMS Craft', 'Sidebar AI Coach widget subtitle'],
    ['coach_prompt_intro', 'Need a quick 5-minute sponge activity, a NotebookLM / OrbitNote workflow, or a de-escalation phrase?', 'Sidebar AI Coach description'],
    ['lounge_heading', 'Overheard at MMS', 'Sidebar quotes widget heading'],
    ['lounge_badge', 'Staff Lounge', 'Sidebar quotes widget badge'],
    ['directory_heading', 'Heart of MMS: Clean Building Directory', 'Directory page main heading'],
    ['directory_subtitle', 'Streamlined, color-coded access to internal USD 266 forms, maps, and operational links.', 'Directory page subtitle'],
    ['bulletin_heading', 'The Flight Wire (Smore Bulletin Archive)', 'Flight Wire page main heading'],
    ['bulletin_subtitle', 'Read administrative newsletters here and search their full text for past announcements.', 'Flight Wire page subtitle'],
    ['studio_heading', 'The Peer Lesson Studio', 'Peer Lesson Studio page main heading'],
    ['studio_subtitle', 'Share a draft lab, CER prompt, or activity for low-stakes Warm & Cool feedback from colleagues.', 'Peer Lesson Studio page subtitle'],
    ['academy_heading', 'EdTech, AI & Classroom Video Lab', 'EdTech Lab page main heading'],
    ['academy_badge', 'Field Toolkit', 'EdTech Lab page badge'],
    ['academy_subtitle', 'NotebookLM in Google Classroom, missing-student slide/video catch-up, content video creation, OrbitNote & essential software.', 'EdTech Lab page subtitle'],
    ['academy_video_url', '', 'Optional YouTube or Google Drive embed URL for the featured walkthrough video'],
    ['academy_video_badge', 'Staff Walkthrough Video • Andrew Hogan', 'Badge above the featured video player'],
    ['academy_video_title', 'Demonstrating NotebookLM in Google Classroom & Missing-Student Catch-Up', 'Default title in the video walkthrough box'],
    ['academy_video_desc', 'Click any chapter timestamp below to preview the workflow steps covered in the walkthrough video.', 'Default description in the video walkthrough box'],
    ['footer_left', 'Classroom Craft, Student Engagement, Teacher Support & Tech/AI', 'Footer description text'],
    ['footer_right', 'Maize Middle School (USD 266)', 'Footer school name text']
  ];
  const toAdd = defaults.filter(row => !existing.has(row[0]) && !(row[0] === 'smore_source_url' && existing.has('source_url')));
  if (toAdd.length) {
    const numCols = Math.max(settingsSheet.getLastColumn(), 3);
    const padded = toAdd.map(r => {
      const copy = r.slice(0, numCols);
      while (copy.length < numCols) copy.push('');
      return copy;
    });
    settingsSheet.getRange(settingsSheet.getLastRow() + 1, 1, padded.length, numCols).setValues(padded);
  }
}

function defaultArticlesSeed_() {
  return [
    {
      article_id: 'art-1',
      title: 'The ‘Low-Floor, High-Ceiling’ Hook: Getting Every 8th Grader Talking in the First 4 Minutes',
      category: 'Student Engagement',
      summary: 'How replacing traditional glossary bellringers with visual phenomena controversies and 90-second table debates transforms passive classes into active investigators.',
      article_body_text: 'The first four minutes of class set the metabolic rate for the entire period. When students walk in and see \'Define photosynthesis in your notebook,\' half the room stalls because they are afraid of being wrong, and the other half copies a glossary line without thinking.\n\nInstead, shift to a \'Low-Floor, High-Ceiling\' phenomenon hook. A low floor means any student—regardless of reading level or yesterday\'s attendance—can make an immediate observation. A high ceiling means the prompt naturally pushes advanced thinkers toward deeper conceptual connections.\n\nFor example, in 8th grade science, project two contrasting photos: a massive oak tree and a tiny packet of soil. Ask: \'Where did the 4,000 pounds of wood actually come from—the dirt, the water, or the thin air?\' Don\'t grade it. Give tables 90 seconds to argue and lock in a table vote.',
      author: 'Science & STEM Team',
      published_at: '2026-10-21',
      cover_image_url: '',
      is_live: 'TRUE',
      takeaway: 'Students shut down when a warm-up feels like a pop quiz. Opening with a counter-intuitive visual mystery where everyone can hazard a guess gets 100% of voices active before formal vocabulary hits the board.',
      try_tomorrow: 'Put up a split image or a \'Which One Doesn\'t Belong?\' prompt with zero academic jargon. Give tables 60 seconds to agree on a defense, then call a table spokesperson to share what their group debated.',
      script: 'Take 30 seconds of silent look-time at the screen. You don\'t need any textbook words yet—just notice what seems weird or surprising. Now, turn to your shoulder partner: Person A, you have 45 seconds to convince Person B of your hunch.',
      read_time: '4 min read',
      featured: 'TRUE',
      has_video_demo: 'FALSE',
      tags: 'bellringers, engagement, discussion, science, inquiry, warm-ups'
    },
    {
      article_id: 'art-7',
      title: 'NotebookLM in Google Classroom: Curriculum Planning & Catch-Up Slides/Videos for Missing Students',
      category: 'Tech & AI',
      summary: 'How to use NotebookLM to map units, organize Google Classroom resources, and generate rapid slide decks and audio/video overviews so absent students never fall behind.',
      article_body_text: 'Unlike open-ended AI chatbots that pull from the random internet, Google\'s NotebookLM is strictly grounded in the documents, Google Slides, PDFs, and YouTube links you upload. That makes it a game-changer for three everyday middle school workflows: Google Classroom resource hubs, curriculum planning, and supporting absent students.\n\n1. Curriculum Planning & Google Classroom Alignment: Drop your Kansas state standards PDF, pacing guide, and past unit labs into a single notebook. Ask NotebookLM to spot vocabulary gaps, sequence daily bellringers, or draft differentiated reading passages that you can export directly to Google Docs and assign in Google Classroom.\n\n2. Slide & Video/Audio Generation for Missing Students: When a student misses a pivotal lab or demonstration, use NotebookLM to generate a concise 5-slide \'What You Missed\' summary outline (ready to paste into Google Slides or Canva) alongside an AI Audio/Video Overview that walks the student through the core phenomenon in conversational language.\n\n3. Student-Facing Study Guides: Post the exported study guide, FAQ, and audio/video overview into your Google Classroom stream before a unit test so students have a 24/7 study guide built strictly from your classroom notes.',
      author: 'Andrew Hogan & EdTech Lab',
      published_at: '2026-10-21',
      cover_image_url: '',
      is_live: 'TRUE',
      takeaway: 'Catching up absent students shouldn\'t eat up 20 minutes of your plan hour. By keeping a unit NotebookLM loaded with your slides, lab docs, and standards, you can generate a 5-slide catch-up outline and a 3-minute audio/video overview to post straight into Google Classroom.',
      try_tomorrow: 'Create one \'Unit Master Notebook\' in NotebookLM and upload this week\'s slides and guided notes. Generate a \'Missing Student Catch-Up Guide\' and post it under a pinned \'Absent? Start Here\' topic in Google Classroom.',
      script: 'NotebookLM Prompt for Absent Students: \'Using only Tuesday\'s lab slides and guided notes, create a 1-page Missing Student Catch-Up Sheet with: 1) The Big Question we investigated, 2) A 5-slide summary outline of key findings, 3) Three vocabulary terms in plain 8th-grade language, and 4) Two check-for-understanding questions they can answer to earn lab makeup credit.\'',
      read_time: '4 min read + Video Demo',
      featured: 'FALSE',
      has_video_demo: 'TRUE',
      tags: 'tech, ai, notebooklm, google classroom, missing students, absent, slides, video, curriculum planning'
    },
    {
      article_id: 'art-2',
      title: 'The 90-Second Transition Reset: Reclaiming Focus Without Raising Your Voice',
      category: 'Classroom Management',
      summary: 'Practical, non-confrontational routines that bridge chaotic hallway passing periods into calm, productive workshop time.',
      article_body_text: 'Middle schoolers carry hallway adrenaline three minutes into every period. When we try to talk over that energy from the front of the room, we accidentally train students that the first two requests don\'t count.\n\nThe most effective reset uses spatial anchoring and visual offloading. Keep a consistent spot on the screen that shows three things only: What stays on your desk, what goes away, and your 2-minute starting task.\n\nWhen transitioning out of group lab work back to whole-class debrief, avoid open-ended countdowns. Give a specific physical anchor: \'In 15 seconds, be Seated, Markers capped, Eyes on the graph.\'',
      author: 'Classroom Craft Collective',
      published_at: '2026-10-19',
      cover_image_url: '',
      is_live: 'TRUE',
      takeaway: 'Competing with middle school volume drains teacher energy by 5th hour. Anchoring transitions to predictable visual cues and physical proximity resets the room faster than verbal warnings.',
      try_tomorrow: 'Stand at the threshold during passing period to greet students by name, and point silently to a single \'Do Immediately\' visual box on the board. Start class from the back or side of the room rather than behind the podium.',
      script: 'Pause where you are. Cap the marker until it clicks. Look at the blue box on the screen—in 60 seconds, your table needs one sticky note on the front board.',
      read_time: '3 min read',
      featured: 'FALSE',
      has_video_demo: 'FALSE',
      tags: 'classroom management, transitions, routines, proximity, de-escalation'
    },
    {
      article_id: 'art-8',
      title: 'Content Videos Dual-Track: Making Quick Teacher Videos & Having Students Record Explainers',
      category: 'Tech & AI',
      summary: 'How 3-minute teacher screencasts clone your support during labs—and how student-created 90-second videos build deeper mastery than worksheets.',
      article_body_text: 'Video in the middle school classroom works best when it is short, authentic, and low-production. Whether you are recording a micro-lesson for your students or handing the camera to them, the secret is keeping the runtime under 3 minutes.\n\nTrack 1 — Teacher-Created Content Videos: Instead of repeating the same multi-step directions 18 times during a lab or math workshop, record a 2-minute screen capture (using Screencastify, Loom, or Canva\'s built-in recorder) over your Google Slides. Embed it in Google Classroom or Edpuzzle. Struggling learners can pause and rewind without embarrassment, and absent students get your exact explanation.\n\nTrack 2 — Student-Created Explainer Videos: When students have to teach a concept on camera in 90 seconds, they can\'t hide behind vague copy-pasted sentences. Have pairs create a \'Paper Slide Video\' or Chromebook screen recording where they point to a diagram, explain the cause-and-effect relationship using 3 target vocabulary words, and bust one common myth.\n\nThe Golden Rule for Student Videos: Never let students open the camera app first. Require a 5-minute \'3-Box Storyboard\' on paper (Box 1: The Hook Question, Box 2: The Visual Diagram/Evidence, Box 3: The Real-World Connection). Once you initial their storyboard, they get 10 minutes to record in one take.',
      author: 'MMS Media & Instructional Team',
      published_at: '2026-10-18',
      cover_image_url: '',
      is_live: 'TRUE',
      takeaway: 'Middle schoolers tune out generic 15-minute YouTube lectures, but they lock in when they hear their own teacher\'s voice in a 3-minute screencast—and comprehension skyrockets when students script and film their own 90-second concept explainers.',
      try_tomorrow: 'Track A (Teacher Video): Record a single-take 2-minute screencast walking through your next lab diagram so students can re-watch at their own pace. Track B (Student Video): Give pairs a \'3-Box Paper Storyboard\' that they must get stamped before recording a 60-second explanation.',
      script: 'Student Video Challenge Script: \'Your pair has 10 minutes to record a 60-to-90 second One-Take Explainer. You must show your annotated diagram on screen, use all three unit vocabulary words accurately, and explain why the common misconception is wrong. No fancy editing—clarity beats special effects!\'',
      read_time: '4 min read',
      featured: 'FALSE',
      has_video_demo: 'FALSE',
      tags: 'tech, ai, video, student videos, screencast, edpuzzle, canva, engagement'
    },
    {
      article_id: 'art-3',
      title: 'The Friday 3:15 Triage: Leaving Work at Work With Monday Already Staged',
      category: 'Teacher Support',
      summary: 'A 15-minute end-of-week ritual designed to protect your weekend bandwidth while eliminating Sunday-night lesson anxiety.',
      article_body_text: 'Teacher burnout rarely comes from the 45 minutes we spend teaching a period; it comes from the cognitive friction of carrying 140 open loops home over the weekend.\n\nInstead of leaving Friday afternoon in a blur or staying until 5:30 PM, run the \'5-5-5 Reset.\' First, do a 3-pile sort on any paper collected that day: Recycle (practice that served its purpose), Quick Scan (check-plus/check/reteach), and Key Assessment.\n\nSecond, stage Monday morning\'s very first bellringer on your desktop and place Monday\'s copies on the front corner of your desk. Walking in Monday morning to a prepped 1st hour is the best gift you can give your future self.',
      author: 'Staff Wellbeing & Sustainability',
      published_at: '2026-10-17',
      cover_image_url: '',
      is_live: 'TRUE',
      takeaway: 'Taking ungraded stacks home in a tote bag rarely leads to productive grading—it just creates weekend guilt. A ruthless 15-minute Friday afternoon triage closes open loops before you walk to your car.',
      try_tomorrow: 'Set a 15-minute timer after the final bell on Friday: 5 minutes to clear your physical desk surface, 5 minutes to stage Monday\'s 1st-hour slide/copies, and 5 minutes to reply to two urgent emails.',
      script: 'Friday 3:15 Checklist: 1) Monday Slide 1 open on desktop, 2) Whiteboard date updated, 3) Emergency sub folder visible on desk corner, 4) Take-home bag empty.',
      read_time: '3 min read',
      featured: 'FALSE',
      has_video_demo: 'FALSE',
      tags: 'teacher support, wellbeing, burnout, grading, boundaries, friday reset'
    },
    {
      article_id: 'art-9',
      title: 'OrbitNote & Essential Classroom Software: Interactive PDFs, Read-Alouds & Time-Saving Tech',
      category: 'Tech & AI',
      summary: 'Unlocking static PDFs in Google Classroom with OrbitNote\'s text-to-speech, color-coded vocabulary highlights, and voice notes—plus our top teacher software picks.',
      article_body_text: 'We all have great curriculum resources locked inside PDFs—lab procedures, historical primary sources, math reference sheets, and science articles. But on a student Chromebook, a plain PDF is tough to read and even harder to annotate.\n\n1. Why OrbitNote Shines in Google Classroom: When students open an assigned PDF with OrbitNote, they immediately get Read&Write accessibility tools built right in: natural-sounding Text-to-Speech with dual-color word tracking, a Picture Dictionary for unfamiliar Tier 2/3 words, and a Screen Mask (reading ruler) that helps students with visual fatigue stay on the right line.\n\n2. High-Leverage Student Annotation Moves in OrbitNote: Beyond typing directly onto worksheets, students can use freehand drawing to label science diagrams or graph trends, use the 4-color highlighters to code Claim-Evidence-Reasoning (and click \'Collect Highlights\' to pull all their highlights into a clean Google Doc), or leave a 20-second Voice Comment explaining their thinking if typing is a barrier.\n\n3. Pairing OrbitNote with Other High-Utility Teacher Software: Combine OrbitNote for accessible PDFs with NotebookLM (for unit synthesis & catch-up slides), Canva for Education (for visual slide decks & student video templates), and Edpuzzle (for embedding quick comprehension questions inside teacher-made videos).',
      author: 'Assistive & Classroom Tech Team',
      published_at: '2026-10-15',
      cover_image_url: '',
      is_live: 'TRUE',
      takeaway: 'Static PDFs frustrate struggling readers and stall Chromebook workflows. Assigning PDFs through OrbitNote in Google Classroom gives every student text-to-speech read-alouds, visual dictionaries, diagram markup, and voice-recorded answers.',
      try_tomorrow: 'Push tomorrow\'s reading or lab PDF through the OrbitNote Google Classroom integration. Ask students to use the Yellow highlighter for the main claim, Green for evidence, and click \'Collect Highlights\' to auto-generate a study sheet.',
      script: 'OrbitNote Classroom Routine: \'Open today\'s article in OrbitNote from Google Classroom. Plug in headphones if you\'d like to use the Play button to listen along. Highlight the author\'s claim in Yellow, two pieces of data in Green, and drop one pushpin text comment next to the graph.\'',
      read_time: '5 min read',
      featured: 'FALSE',
      has_video_demo: 'FALSE',
      tags: 'tech, ai, orbitnote, orbit note, pdf, accessibility, google classroom, software, read aloud'
    },
    {
      article_id: 'art-4',
      title: 'Killing the Dictionary Copy-Paste: Morphological Root Walls & Visual Frayer Tweaks',
      category: 'Student Engagement',
      summary: 'Lightweight, 5-minute disciplinary literacy routines that build real Tier 2 & Tier 3 vocabulary stamina across every subject.',
      article_body_text: 'Across science, social studies, math, and ELA, academic vocabulary is the single biggest bottleneck for comprehension. Yet traditional vocabulary lists treat every word as an isolated island to memorize.\n\nWhen we teach high-leverage roots (like photo-, therm-, chron-, morph-, equi-), we hand students a master key that unlocks dozens of unfamiliar terms on state assessments and lab instructions.\n\nCombine root dissection with the \'Modified Frayer Sketch\': give students 3 minutes to draw what a concept looks like in action and write the #1 misconception a 6th grader might have about it.',
      author: 'Disciplinary Literacy Team',
      published_at: '2026-10-14',
      cover_image_url: '',
      is_live: 'TRUE',
      takeaway: 'When students copy glossary definitions, retention drops near zero within 48 hours. Teaching word roots and requiring sketched \'non-examples\' forces genuine conceptual processing.',
      try_tomorrow: 'Swap out two quadrants of the classic Frayer model: replace \'Definition\' with \'Sketch It (Zero Words Allowed)\' and \'Non-Examples\' with \'The Common Trap (What people confuse this with).\'',
      script: 'Don\'t tell me what \'equilibrium\' means from Google. Draw a 10-second stick-figure sketch of what equilibrium looks like in a hallway, and write one thing that looks like equilibrium but isn\'t.',
      read_time: '5 min read',
      featured: 'FALSE',
      has_video_demo: 'FALSE',
      tags: 'literacy, vocabulary, frayer, roots, wigs, engagement'
    },
    {
      article_id: 'art-5',
      title: 'De-Escalating the ‘I’m Not Doing This’ Shutdown in Under 60 Seconds',
      category: 'Classroom Management',
      summary: 'How to side-step power struggles with reluctant middle schoolers using micro-choices, private crouch-ins, and low-stakes entry ramps.',
      article_body_text: 'In 7th and 8th grade, social status is everything. When a student refuses to start a lab write-up or math set, it\'s frequently masking reading fatigue or fear of looking incapable in front of peers.\n\nThe key rule of thumb: Praise in public, redirect in whisper, and always walk away after giving a micro-choice. Lingering over a student\'s desk waiting for immediate compliance creates a stage where they feel forced to hold their ground.\n\nOffer a two-option ramp: \'Do you want to jot bullet points for question 1, or sketch the diagram for question 2 first? I\'ll swing back in three minutes to see which one you picked.\'',
      author: 'MTSS & Student Support',
      published_at: '2026-10-11',
      cover_image_url: '',
      is_live: 'TRUE',
      takeaway: 'Publicly confronting task refusal in middle school almost always triggers a face-saving power struggle. A quiet 10-second crouch-in offering a micro-choice preserves dignity and restarts momentum.',
      try_tomorrow: 'When a student pushes a paper away, drop down beside the desk privately, validate the friction in 5 words, offer a \'Start with #2 or #4?\' choice, and immediately walk away to give them space to comply without an audience.',
      script: 'Hey, I know this graph looks dense right now. Do you want to tackle just the X-axis label first, or partner with Sam for the first two rows? Your call—I\'ll check back in two minutes.',
      read_time: '4 min read',
      featured: 'FALSE',
      has_video_demo: 'FALSE',
      tags: 'de-escalation, behavior, management, reluctant learners, mtss'
    },
    {
      article_id: 'art-6',
      title: 'The 90-Second Feedback Loop: Grading Less While Giving Better Direction',
      category: 'Teacher Support',
      summary: 'Why whole-class \'Spotlight & Fix\' slides beat spending four hours writing the same margin comment on 130 student papers.',
      article_body_text: 'We\'ve all spent a Sunday afternoon writing \'Need to connect your evidence back to your claim!\' on 85 consecutive lab reports—only to watch students glance at the score and shove the paper into their binder.\n\nWhole-Class Feedback flips the ratio: spend 12 minutes flipping through the stack with a single notepad. Jot down 3 student names whose work nailed the target, and 2 common traps that tripped up the class.\n\nOpen the next class by celebrating the exemplar moves and projecting a composite \'Almost There\' response. Have students do the cognitive work of fixing the gap live in class.',
      author: 'Instructional Coaching Hub',
      published_at: '2026-10-08',
      cover_image_url: '',
      is_live: 'TRUE',
      takeaway: 'Students rarely read paragraphs of red ink returned a week later. Scanning work for the top two building-wide patterns and opening the next day with an anonymous \'Upgrade This Response\' warmup saves hours.',
      try_tomorrow: 'Instead of writing comments on every CER paragraph tonight, tally the top 3 errors on a sticky note. Tomorrow, project two anonymous student sentences and let pairs spend 3 minutes upgrading them.',
      script: 'I read through everyone\'s lab claims yesterday. Here are two anonymized sentences from our class that have great evidence, but they\'re missing the \'because\' bridge. With your partner, rewrite Sentence A on your mini-whiteboard.',
      read_time: '4 min read',
      featured: 'FALSE',
      has_video_demo: 'FALSE',
      tags: 'grading, feedback, time savers, teacher support, cer'
    }
  ];
}

function defaultQuickWinsSeed_() {
  return [
    { tip_id: 'qw-1', section: 'all', category_label: 'Student Engagement', title: 'The ‘Wrong Answer Only’ Brainstorm', tip: 'Lower participation anxiety by asking tables for the most plausible incorrect answer to a complex question first—then have them explain why it\'s a tempting trap.', is_live: 'TRUE' },
    { tip_id: 'qw-2', section: 'all', category_label: 'Tech & AI (NotebookLM)', title: 'Instant ‘Missed Class’ Catch-Up Link', tip: 'Drop today\'s slides into your unit NotebookLM to auto-generate a 5-point summary and audio overview—paste it into Google Classroom for absent students in 60 seconds.', is_live: 'TRUE' },
    { tip_id: 'qw-3', section: 'all', category_label: 'Classroom Management', title: 'The Threshold Greeting & Point', tip: 'Greet students at the door by name, then gesture directly toward the board\'s \'First 2 Minutes\' box. It cuts transition chatter in half before the bell rings.', is_live: 'TRUE' },
    { tip_id: 'qw-4', section: 'Student Engagement', category_label: 'Bellringer Twist', title: 'The ‘Wrong Answer Only’ Brainstorm', tip: 'Ask tables to invent the most believable wrong answer to today\'s question. Analyzing a misconception hooks reluctant speakers without fear of being wrong.', is_live: 'TRUE' },
    { tip_id: 'qw-5', section: 'Student Engagement', category_label: 'Total Participation', title: 'The 60-Second Stand & Vote Spectrum', tip: 'Have students stand and point toward Wall A or Wall B to vote on a phenomenon prediction before any direct instruction begins—instant kinetic buy-in.', is_live: 'TRUE' },
    { tip_id: 'qw-6', section: 'Student Engagement', category_label: 'Vocabulary Hook', title: 'Sketch It With Zero Words', tip: 'Give students 45 seconds on mini-whiteboards to draw a stick-figure scene of a Tier 3 term in action before showing the formal textbook definition.', is_live: 'TRUE' },
    { tip_id: 'qw-7', section: 'Classroom Management', category_label: 'Passing Period', title: 'The Threshold Greeting & Point', tip: 'Greet students at the door by name, then gesture silently toward the board\'s \'Do Immediately\' box so expectations start at the doorframe.', is_live: 'TRUE' },
    { tip_id: 'qw-8', section: 'Classroom Management', category_label: 'De-Escalation', title: 'The 10-Second Crouch & Walk-Away', tip: 'Drop to eye level beside a reluctant student\'s desk, offer a \'Start with #1 or #3?\' choice in a whisper, and walk away immediately so they save face.', is_live: 'TRUE' },
    { tip_id: 'qw-9', section: 'Classroom Management', category_label: 'Lab & Group Reset', title: 'The ‘Marker Click’ Physical Anchor', tip: 'Replace countdown yelling with a concrete physical cue: \'Pause where you are, cap your marker until it clicks, and eyes on the blue graph.\'', is_live: 'TRUE' },
    { tip_id: 'qw-10', section: 'Teacher Support', category_label: 'Grading Triage', title: 'The 3-Pile Exit Ticket Sort', tip: 'Skip logging every formative slip into Skyward. Sort papers into \'Got It\', \'Almost\', and \'Reteach\' in 90 seconds to guide tomorrow\'s warm-up.', is_live: 'TRUE' },
    { tip_id: 'qw-11', section: 'Teacher Support', category_label: 'Focus Armor', title: 'The Closed-Tab Passing Period Rule', tip: 'Keep Outlook closed during instructional periods. Reading a stressful email during a 4-minute passing period hijacks your working memory for the next class.', is_live: 'TRUE' },
    { tip_id: 'qw-12', section: 'Teacher Support', category_label: 'Friday Ritual', title: 'Stage Monday\'s 1st Hour Before 3:30', tip: 'Spend 5 minutes Friday afternoon opening Monday\'s bellringer slide and laying out 1st-hour copies so you walk out with zero Sunday-night dread.', is_live: 'TRUE' },
    { tip_id: 'qw-13', section: 'Tech & AI', category_label: 'NotebookLM + Google Classroom', title: 'Absent-Student Slide & Audio Catch-Up', tip: 'Upload your unit slides and lab sheet into NotebookLM to generate a 5-slide catch-up outline and audio overview—drop both into Google Classroom for missing students.', is_live: 'TRUE' },
    { tip_id: 'qw-14', section: 'Tech & AI', category_label: 'OrbitNote Workflow', title: '2-Color Highlight & ‘Collect’ Routine', tip: 'Assign PDFs via OrbitNote in Google Classroom: have students highlight Claims in Yellow and Evidence in Green, then click \'Collect Highlights\' into a clean study doc.', is_live: 'TRUE' },
    { tip_id: 'qw-15', section: 'Tech & AI', category_label: 'Classroom Videos', title: 'The 3-Box Storyboard Before Recording', tip: 'Before students record a 90-second concept video, require a 3-box paper sketch (Hook -> Visual Proof -> Vocab Tie-In) so they nail it in one take.', is_live: 'TRUE' }
  ];
}

function defaultDirectorySeed_() {
  return [
    { resource_id: 'res-1', title: 'Sub Coverage & Absence Request', category: 'Daily Operations', description: 'Frontline Aesop login, district sub request guidelines, and emergency roster submission.', url: 'https://drive.google.com/open?id=sample_sub_request', tags: 'sub, sick, absence, coverage, frontline, leave', is_live: 'TRUE' },
    { resource_id: 'res-2', title: 'Crisis Response & Evacuation Maps', category: 'Emergency & Logistics', description: 'Standard Response Protocol (SRP), severe weather safe rooms, and fire drill routes.', url: 'https://drive.google.com/open?id=sample_crisis_plan', tags: 'emergency, fire, tornado, lockdown, srp, maps', is_live: 'TRUE' },
    { resource_id: 'res-3', title: 'Student MTSS & Tier 2 Referral', category: 'Student Services & MTSS', description: 'Academic and behavioral intervention intake forms, SIT team notes, and progress monitoring.', url: 'https://drive.google.com/open?id=sample_mtss_referral', tags: 'mtss, tier 2, tier 3, sit, referral, intervention', is_live: 'TRUE' },
    { resource_id: 'res-4', title: 'MMS Staff Directory & Room Map', category: 'Directories & Handbooks', description: 'Building extensions, teacher classroom numbers, grade-level leads, and duty rosters.', url: 'https://drive.google.com/open?id=sample_directory', tags: 'phone, extension, room, roster, directory, biggs, hogan', is_live: 'TRUE' },
    { resource_id: 'res-5', title: 'Discipline Referral (Major / Minor)', category: 'Emergency & Logistics', description: 'Skyward discipline tracking protocol, administrative triage, and parent contact logs.', url: 'https://drive.google.com/open?id=sample_discipline', tags: 'discipline, referral, skyward, incident, admin', is_live: 'TRUE' },
    { resource_id: 'res-6', title: 'Facility Maintenance & Work Orders', category: 'Daily Operations', description: 'Custodial requests, HVAC adjustments, key card access, and classroom repairs.', url: 'https://drive.google.com/open?id=sample_custodial', tags: 'custodial, maintenance, repair, hvac, work order', is_live: 'TRUE' },
    { resource_id: 'res-7', title: 'Nurse Pass & Medical Protocols', category: 'Student Services & MTSS', description: 'Clinic guidelines, medication distribution times, and emergency allergy action plans.', url: 'https://drive.google.com/open?id=sample_nurse_clinic', tags: 'nurse, clinic, medication, allergy', is_live: 'TRUE' },
    { resource_id: 'res-8', title: 'Certified Staff Handbook (USD 266)', category: 'Directories & Handbooks', description: 'Contract hours, personal leave policies, evaluation rubrics, and professional expectations.', url: 'https://drive.google.com/open?id=sample_handbook', tags: 'handbook, contract, policy, dress code', is_live: 'TRUE' }
  ];
}

function defaultQuotesSeed_() {
  return [
    { quote_id: 'q-1', quote: 'Mr. H, if the Earth is spinning at 1,000 miles an hour, why don’t my shoes wear out faster when I stand still?', attribution: '8th Grade Science Debrief', is_live: 'TRUE' },
    { quote_id: 'q-2', quote: 'I didn’t lose my homework, it’s just currently in an unverified location in my locker.', attribution: '7th Hour Hallway', is_live: 'TRUE' }
  ];
}

function defaultStudioSeed_() {
  return [
    {
      studio_id: 'studio-1',
      title: '8th Grade Plate Tectonics CER Lab Prompt',
      subject: '8th Science',
      dilemma: 'Opening Hook & Student Buy-In',
      warm_feedback: 'Opening with the sonar map mystery hooked even reluctant readers immediately.',
      cool_feedback: 'I wonder if trimming the background paragraph to 3 bullet points would get them into the data faster.',
      doc_url: '',
      is_live: 'TRUE'
    }
  ];
}

function defaultTechToolsSeed_() {
  return [
    {
      tool_id: 'tool-1',
      badge: 'PDF & Accessibility Star',
      title: 'OrbitNote (by Texthelp)',
      link_label: 'OrbitNote',
      url: 'https://orbitnote.texthelp.com/',
      description: 'Transforms static PDFs in Google Classroom into interactive, accessible workspaces with Read&Write tools built in.',
      teacher_workflow: 'Assign PDFs directly via Google Classroom; leave voice comments for fast grading; use freehand drawing to model math/science diagrams live.',
      student_workflow: 'Use Text-to-Speech read-aloud with dual-color word tracking, Picture Dictionary for Tier 3 terms, 4-color highlighters + "Collect Highlights" into a study doc, and Voice Notes.',
      is_live: 'TRUE'
    },
    {
      tool_id: 'tool-2',
      badge: 'AI Curriculum Co-Pilot',
      title: 'Google NotebookLM',
      link_label: 'NotebookLM',
      url: 'https://notebooklm.google.com/',
      description: 'Source-grounded AI notebook that only uses the Google Docs, Slides, PDFs, and links you upload—zero outside hallucinations.',
      teacher_workflow: 'Generate unit pacing guides, leveled reading passages, DOK assessment audits, and 5-slide catch-up outlines for missing students.',
      student_workflow: 'Listen to the 2-host Audio Overview podcast at stations or after an absence, and review teacher-verified unit FAQs and study guides in Google Classroom.',
      is_live: 'TRUE'
    },
    {
      tool_id: 'tool-3',
      badge: 'Visuals & Video Studio',
      title: 'Canva for Education & Slides',
      link_label: 'Canva Edu',
      url: 'https://www.canva.com/education/',
      description: 'All-in-one visual deck builder and browser-based video recorder (100% free for K-12 schools).',
      teacher_workflow: 'Paste NotebookLM slide outlines into visual templates and use "Present & Record" to make a 2-minute video link for Google Classroom.',
      student_workflow: 'Collaborate in pairs on infographics, Frayer visual posters, or 90-second voice-over explainer videos without complex editing software.',
      is_live: 'TRUE'
    },
    {
      tool_id: 'tool-4',
      badge: 'Interactive Video',
      title: 'Edpuzzle & Screencastify',
      link_label: 'Edpuzzle',
      url: 'https://edpuzzle.com/',
      description: 'Turn passive video watching into active check-for-understanding pauses synced directly with Google Classroom grades.',
      teacher_workflow: 'Trim any content clip to 3 minutes, embed 3 pause-and-predict questions, and turn on "Prevent Skipping" for station accountability.',
      student_workflow: 'Re-watch the exact 20-second segment before a question if they get stuck, building self-paced comprehension.',
      is_live: 'TRUE'
    },
    {
      tool_id: 'tool-5',
      badge: 'Chrome Extension',
      title: 'Brisk Teaching & Diffit',
      link_label: 'Brisk / Diffit',
      url: 'https://www.briskteaching.com/',
      description: 'Instant reading-level adaptation and Google Docs feedback helpers that live right inside your Chrome browser.',
      teacher_workflow: 'Convert any web article into a 6th, 7th, or 8th-grade Google Doc or Google Slides deck in 10 seconds, then push it to OrbitNote.',
      student_workflow: 'Access core grade-level concepts with scaffolded sentence stems and built-in vocabulary banks.',
      is_live: 'TRUE'
    }
  ];
}
