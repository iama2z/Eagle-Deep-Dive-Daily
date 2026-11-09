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
  'resource_id', 'title', 'category', 'description', 'url', 'tags', 'update_status', 'is_live'
];

const LOUNGE_QUOTE_COLUMNS = [
  'quote_id', 'quote', 'attribution', 'is_live', 'post_type', 'appreciate', 'empathy', 'respect'
];

const LOUNGE_REACTIONS = ['appreciate', 'empathy', 'respect'];

const PEER_STUDIO_COLUMNS = [
  'studio_id', 'title', 'subject', 'dilemma', 'warm_feedback', 'cool_feedback', 'doc_url', 'is_live'
];

const TECH_TOOL_COLUMNS = [
  'tool_id', 'badge', 'title', 'link_label', 'url', 'description', 'teacher_workflow', 'student_workflow', 'is_live'
];

const IMAGE_LIBRARY_COLUMNS = ['image_id', 'category', 'label', 'image_url', 'is_live'];

// Page structure tabs. The site code defines the available components (kind, block_type, style);
// these tabs define which ones appear, their text, order and visibility.
// kind: blocks | videos | tools.  layout: stack | columns.
const PAGE_SECTION_COLUMNS = [
  'section_id', 'page', 'kind', 'layout', 'title', 'intro', 'badge', 'sort_order', 'is_live'
];

// block_type: text | list | callout | prompt.  style: neutral | red | green (list: bullets | numbered).
// For list blocks, put one item per line in body. For prompt blocks, body is the copyable text.
const CONTENT_BLOCK_COLUMNS = [
  'block_id', 'section_id', 'block_type', 'title', 'body', 'description', 'url', 'link_label',
  'style', 'sort_order', 'is_live'
];

const VIDEO_COLUMNS = [
  'video_id', 'section_id', 'badge', 'title', 'description', 'url', 'link_label', 'sort_order', 'is_live'
];

// start_time accepts m:ss, h:mm:ss or plain seconds.
const VIDEO_CHAPTER_COLUMNS = [
  'chapter_id', 'video_id', 'title', 'description', 'start_time', 'sort_order', 'is_live'
];

const SUBMISSION_COLUMNS = [
  'id', 'submitted_at', 'submitter_name', 'submitter_email', 'submission_type',
  'title', 'content', 'links', 'status', 'category', 'reviewed_by', 'reviewed_at', 'cover_image_url'
];

const SUBMISSION_STATUSES = ['pending', 'approved', 'published', 'delete'];

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
    .addItem('4. Sync School News & Social', 'syncSchoolUpdates')
    .addSeparator()
    .addItem('5. Unpublish Sample Rows', 'unpublishSampleRows')
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
  const articlesSheetWasEmpty = !articlesSheet.getLastRow();
  const articleHeaders = smoreHeaders_(articlesSheet, PUBLISHED_ARTICLE_COLUMNS, true);
  if (articlesSheetWasEmpty) {
    seedSheetRows_(articlesSheet, articleHeaders, unpublishedSamples_(defaultArticlesSeed_()));
  }

  const quickWinsSheet = ss.getSheetByName('quick_wins') || ss.insertSheet('quick_wins');
  const quickWinsSheetWasEmpty = !quickWinsSheet.getLastRow();
  const quickWinsHeaders = smoreHeaders_(quickWinsSheet, QUICK_WINS_COLUMNS, true);
  if (quickWinsSheetWasEmpty) {
    seedSheetRows_(quickWinsSheet, quickWinsHeaders, unpublishedSamples_(defaultQuickWinsSeed_()));
  }

  const directorySheet = ss.getSheetByName('directory_links') || ss.insertSheet('directory_links');
  const directorySheetWasEmpty = !directorySheet.getLastRow();
  const directoryHeaders = smoreHeaders_(directorySheet, DIRECTORY_LINK_COLUMNS, true);
  if (directorySheetWasEmpty) {
    seedSheetRows_(directorySheet, directoryHeaders, unpublishedSamples_(defaultDirectorySeed_()));
  }

  const quotesSheet = ss.getSheetByName('lounge_quotes') || ss.insertSheet('lounge_quotes');
  const quotesSheetWasEmpty = !quotesSheet.getLastRow();
  const quotesHeaders = smoreHeaders_(quotesSheet, LOUNGE_QUOTE_COLUMNS, true);
  if (quotesSheetWasEmpty) {
    seedSheetRows_(quotesSheet, quotesHeaders, unpublishedSamples_(defaultQuotesSeed_()));
  }

  const studioSheet = ss.getSheetByName('peer_studio') || ss.insertSheet('peer_studio');
  const studioSheetWasEmpty = !studioSheet.getLastRow();
  const studioHeaders = smoreHeaders_(studioSheet, PEER_STUDIO_COLUMNS, true);
  if (studioSheetWasEmpty) {
    seedSheetRows_(studioSheet, studioHeaders, unpublishedSamples_(defaultStudioSeed_()));
  }

  const techSheet = ss.getSheetByName('tech_tools') || ss.insertSheet('tech_tools');
  const techSheetWasEmpty = !techSheet.getLastRow();
  const techHeaders = smoreHeaders_(techSheet, TECH_TOOL_COLUMNS, true);
  if (techSheetWasEmpty) {
    seedSheetRows_(techSheet, techHeaders, unpublishedSamples_(defaultTechToolsSeed_()));
  }

  [
    ['page_sections', PAGE_SECTION_COLUMNS, defaultPageSectionsSeed_()],
    ['content_blocks', CONTENT_BLOCK_COLUMNS, unpublishedSamples_(defaultContentBlocksSeed_())],
    ['videos', VIDEO_COLUMNS, unpublishedSamples_(defaultVideosSeed_())],
    ['video_chapters', VIDEO_CHAPTER_COLUMNS, unpublishedSamples_(defaultVideoChaptersSeed_())]
  ].forEach(([name, columns, seed]) => {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    const wasEmpty = !sheet.getLastRow();
    const headers = smoreHeaders_(sheet, columns, true);
    if (wasEmpty) seedSheetRows_(sheet, headers, seed);
  });

  const imageSheet = ss.getSheetByName('image_library') || ss.insertSheet('image_library');
  smoreHeaders_(imageSheet, IMAGE_LIBRARY_COLUMNS, true);

  const submissionsSheet = ss.getSheetByName('submissions') || ss.insertSheet('submissions');
  applyStatusDropdown_(submissionsSheet, smoreHeaders_(submissionsSheet, SUBMISSION_COLUMNS, true));

  const reviewQueueSheet = ss.getSheetByName('review_queue') || ss.insertSheet('review_queue');
  applyStatusDropdown_(reviewQueueSheet, smoreHeaders_(reviewQueueSheet, SUBMISSION_COLUMNS, true));

  const auditSheet = ss.getSheetByName('audit_log') || ss.insertSheet('audit_log');
  smoreHeaders_(auditSheet, AUDIT_LOG_COLUMNS, true);
}

function applyStatusDropdown_(sheet, headers) {
  const col = headers.indexOf('status') + 1;
  if (!col) return;
  const rows = Math.max(sheet.getMaxRows() - 1, 1000);
  if (sheet.getMaxRows() < rows + 1) sheet.insertRowsAfter(sheet.getMaxRows(), rows + 1 - sheet.getMaxRows());
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(SUBMISSION_STATUSES, true)
    .setAllowInvalid(false)
    .setHelpText('pending = waiting, approved = publish on next run, published = live, delete = remove from site and sheet')
    .build();
  sheet.getRange(2, col, rows, 1).setDataValidation(rule);
}

function liveTarget_(ss, type) {
  if (type === 'quote' || type === 'lounge_quote') {
    const sheet = ss.getSheetByName('lounge_quotes');
    return { sheet, headers: smoreHeaders_(sheet, LOUNGE_QUOTE_COLUMNS, true), idKey: 'quote_id' };
  }
  if (type === 'lesson_draft' || type === 'studio') {
    const sheet = ss.getSheetByName('peer_studio');
    return { sheet, headers: smoreHeaders_(sheet, PEER_STUDIO_COLUMNS, true), idKey: 'studio_id' };
  }
  const sheet = getPublishedArticlesSheet_(ss);
  return { sheet, headers: smoreHeaders_(sheet, PUBLISHED_ARTICLE_COLUMNS, true), idKey: 'article_id' };
}

function liveRowsById_(target, id) {
  if (!id) return [];
  return smoreReadIssues_(target.sheet, target.headers).filter(r => r.issue[target.idKey] === id);
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
    syncSmoreFeedUnlocked_();
  } finally {
    lock.releaseLock();
  }
}

function requestPublicSmoreSync_(customSourceUrl) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) {
    return { ok: false, code: 'busy', error: 'Another newsletter sync is running. Please try again shortly.', retry_after_seconds: 10 };
  }
  try {
    let customUrl = '';
    if (customSourceUrl) {
      customUrl = smoreNormalizeSource_(customSourceUrl);
      if (!customUrl) {
        return { ok: false, code: 'invalid_url', error: 'Enter a public Smore newsletter (smore.com/n/...) or profile (smore.com/u/...) URL.', retry_after_seconds: 0 };
      }
    }
    let items;
    if (customUrl) {
      // Saved to the sheets so the archive persists across reloads.
      saveSmoreSourceIssues_(customUrl);
      items = publicSmoreItems_(true);
    } else {
      syncSmoreFeedUnlocked_();
      items = publicSmoreItems_(true);
    }
    return {
      ok: true, action: 'sync_smore', items: items, preview: false,
      retry_after_seconds: 0
    };
  } catch (err) {
    console.error('Public Smore sync failed: ' + err.message);
    return {
      ok: false, code: 'sync_failed', error: String(err.message),
      retry_after_seconds: 0
    };
  } finally {
    lock.releaseLock();
  }
}

function smoreNormalizeSource_(value) {
  const text = String(value || '').trim();
  if (text.length > 300) return '';
  const issue = smoreIssueUrl_(text);
  if (issue) return issue;
  const match = text.match(/^https:\/\/(?:www\.|app\.)?smore\.com\/u\/([a-z0-9_-]+)\/?$/i);
  return match ? 'https://www.smore.com/u/' + match[1] : '';
}

function saveSmoreSourceIssues_(sourceUrl) {
  setupSiteSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('smore_feed_raw');
  const published = ss.getSheetByName('smore_feed_public');
  const rawHeaders = smoreHeaders_(raw, SMORE_COLUMNS, true);
  const publicHeaders = smoreHeaders_(published, SMORE_COLUMNS, true);
  const now = new Date().toISOString();
  const saved = [];
  const failures = [];
  smoreDiscover_(sourceUrl).slice(0, 10).forEach(url => {
    try {
      saved.push(Object.assign(smoreParseIssue_(smoreFetchIssue_(url), url), { last_synced_at: now }));
    } catch (err) {
      failures.push(url + ': ' + err.message);
    }
  });
  if (!saved.length) throw new Error('No readable newsletters found at that URL.' + (failures.length ? ' ' + failures[0] : ''));
  smoreWriteIssues_(raw, rawHeaders, saved);
  smoreWriteIssues_(published, publicHeaders, saved);
  appendAuditLog_(ss, 'smore_add', 'smore_feed_public', 'batch', 'Smore Add', 'Saved: ' + saved.length + ', Failures: ' + failures.length);
}

function smorePreviewItems_(sourceUrl) {
  const urls = smoreDiscover_(sourceUrl).slice(0, 10);
  const previews = [];
  const failures = [];
  urls.forEach(url => {
    try {
      previews.push(smoreParseIssue_(smoreFetchIssue_(url), url));
    } catch (err) {
      failures.push(url + ': ' + err.message);
    }
  });
  if (!previews.length) throw new Error('No readable newsletters found at that URL.' + (failures.length ? ' ' + failures[0] : ''));
  const byUrl = new Map(publicSmoreItems_(false).map(item => [item.issue_url, item]));
  previews.forEach(issue => {
    const url = smoreIssueUrl_(issue.issue_url);
    byUrl.set(url, Object.fromEntries(SMORE_COLUMNS.map(key => [key, key === 'issue_url' ? url : (issue[key] || '')])));
  });
  return [...byUrl.values()];
}

function syncSmoreFeedUnlocked_() {
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
    if (feed === 'school') {
      return smoreJson_({ ok: true, updates: publicSchoolUpdates_() });
    }
    if (feed === 'articles') {
      return smoreJson_({ ok: true, items: publishedArticles_() });
    }
    if (feed === 'site' || feed === 'all') {
      ensureLoungeQuoteIds_();
      return smoreJson_({
        ok: true,
        settings: publicSettings_(),
        articles: publishedArticles_(),
        smore: publicSmoreItems_(false),
        quick_wins: publicTableItems_('quick_wins', QUICK_WINS_COLUMNS, ['title', 'tip']),
        directory: publicTableItems_('directory_links', DIRECTORY_LINK_COLUMNS, ['title', 'url']),
        quotes: publicTableItems_('lounge_quotes', LOUNGE_QUOTE_COLUMNS, ['quote']),
        studio: publicTableItems_('peer_studio', PEER_STUDIO_COLUMNS, ['title']),
        tech_tools: publicTableItems_('tech_tools', TECH_TOOL_COLUMNS, ['title']),
        image_library: publicTableItems_('image_library', IMAGE_LIBRARY_COLUMNS, ['image_url'], true),
        page_sections: publicTableItems_('page_sections', PAGE_SECTION_COLUMNS, ['section_id'], false, true),
        content_blocks: publicTableItems_('content_blocks', CONTENT_BLOCK_COLUMNS, ['section_id', 'block_type'], false, true),
        videos: publicTableItems_('videos', VIDEO_COLUMNS, ['video_id', 'url'], false, true),
        video_chapters: publicTableItems_('video_chapters', VIDEO_CHAPTER_COLUMNS, ['video_id', 'title'], false, true)
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
  const required = ['article_id', 'title', 'category', 'is_live'];
  if (!headers.length || required.some(key => !headers.includes(key))) {
    throw new Error('Missing required published_articles headers: ' + required.join(', '));
  }
  if (!headers.includes('article_body_text') && !headers.includes('body')) {
    throw new Error('Missing required published_articles header: article_body_text or body');
  }
  return rows.filter(row => {
    const val = row[headers.indexOf('is_live')];
    return val === true || String(val || '').trim().toLowerCase() === 'true';
  }).map(row => Object.fromEntries(PUBLISHED_ARTICLE_COLUMNS.map(key => {
    const idx = headers.indexOf(key);
    let value = idx >= 0 ? row[idx] : '';
    if ((key === 'article_body_text' || key === 'cover_image_url') && !String(value ?? '').trim()) {
      const aliasIdx = headers.indexOf(key === 'article_body_text' ? 'body' : 'image_url');
      if (aliasIdx >= 0) value = row[aliasIdx];
    }
    if (key === 'is_live') return [key, true];
    if (key === 'featured' || key === 'has_video_demo') {
      return [key, value === true || String(value || '').trim().toLowerCase() === 'true'];
    }
    return [key, value instanceof Date ? value.toISOString() : String(value == null ? '' : value).trim()];
  }))).filter(item => required.every(key => item[key]) && item.article_body_text);
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

// requireTab: a missing tab or is_live header is a setup error (the site shows a load failure),
// not an intentionally empty list.
function publicTableItems_(sheetName, columns, requiredKeys, blankIsLive = false, requireTab = false) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    if (requireTab) throw new Error('Missing tab "' + sheetName + '". Run Setup / Seed Site Tabs.');
    return [];
  }
  if (sheet.getLastRow() < 2) return [];
  const [rawHeaders, ...rows] = sheet.getDataRange().getValues();
  const headers = rawHeaders.map(h => String(h || '').trim());
  if (!headers.length) return [];
  const liveIdx = headers.indexOf('is_live');
  if (requireTab && liveIdx < 0) throw new Error('Tab "' + sheetName + '" needs an is_live column.');
  return rows.filter(row => {
    if (liveIdx < 0) return true;
    const val = row[liveIdx];
    const live = String(val ?? '').trim().toLowerCase();
    return val === true || live === 'true' || (blankIsLive && live === '');
  }).map(row => Object.fromEntries(columns.map(key => {
    const idx = headers.indexOf(key);
    const value = idx >= 0 ? row[idx] : '';
    if (key === 'is_live') return [key, true];
    return [key, value instanceof Date ? value.toISOString() : String(value == null ? '' : value).trim()];
  }))).filter(item => (requiredKeys || []).every(k => item[k]));
}

// Persist IDs before publishing manually entered posts to the feed.
function ensureLoungeQuoteIds_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('lounge_quotes');
  if (!sheet || sheet.getLastRow() < 2) return;
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const headers = smoreHeaders_(sheet, LOUNGE_QUOTE_COLUMNS, true);
    const idCol = headers.indexOf('quote_id') + 1;
    const quoteIdx = headers.indexOf('quote');
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][quoteIdx] || '').trim() && !String(rows[i][idCol - 1] || '').trim()) {
        sheet.getRange(i + 1, idCol).setValue('q_' + Utilities.getUuid());
      }
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function addLoungeReaction_(quoteId, reaction, undo) {
  quoteId = String(quoteId || '').trim();
  reaction = String(reaction || '').trim().toLowerCase();
  if (!quoteId || LOUNGE_REACTIONS.indexOf(reaction) < 0) {
    return { ok: false, error: 'Invalid reaction.' };
  }
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('lounge_quotes');
  if (!sheet) return { ok: false, error: 'Staff Lounge is not set up yet.' };
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const headers = smoreHeaders_(sheet, LOUNGE_QUOTE_COLUMNS, true);
    const idCol = headers.indexOf('quote_id') + 1;
    const countCol = headers.indexOf(reaction) + 1;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return { ok: false, error: 'Post not found.' };
    const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]).trim() === quoteId) {
        const cell = sheet.getRange(i + 2, countCol);
        const next = Math.max(0, (Number(cell.getValue()) || 0) + (undo ? -1 : 1));
        cell.setValue(next);
        SpreadsheetApp.flush();
        return { ok: true, quote_id: quoteId, reaction: reaction, count: next };
      }
    }
    return { ok: false, error: 'Post not found.' };
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  try {
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

    if (body.action === 'sync_smore') {
      return smoreJson_(requestPublicSmoreSync_(body.source_url));
    }
    if (body.action === 'react') {
      return smoreJson_(addLoungeReaction_(body.quote_id, body.reaction, body.undo === 'true'));
    }
    if (String(body.action || '').indexOf('admin_') === 0) {
      return smoreJson_(handleAdminAction_(body));
    }
    if (body.action) {
      return smoreJson_({ ok: false, error: 'Unknown action' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('submissions') || ss.insertSheet('submissions');
    const headers = smoreHeaders_(sheet, SUBMISSION_COLUMNS, true);

    const submitterName = cleanSubmissionText_(body.submitter_name || '');
    const submitterEmail = cleanSubmissionText_(body.submitter_email || '');
    const submissionType = cleanSubmissionText_(body.submission_type || '');
    const title = cleanSubmissionText_(body.title || '');
    const content = String(body.content || '').replace(/<[^>]*>/g, ' ').replace(/\r\n?/g, '\n').trim();
    const links = cleanSubmissionText_(body.links || '');
    const category = cleanSubmissionText_(body.category || '');
    const coverImageUrl = String(body.cover_image_url || '').trim();
    if (coverImageUrl && !/^https?:\/\/[^\s/?#]+(?:[/?#][^\s]*)?$/i.test(coverImageUrl)) {
      return smoreJson_({ ok: false, error: 'Cover image link must be an http:// or https:// URL.' });
    }

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
      cover_image_url: coverImageUrl,
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

// ---------------------------------------------------------------
// Admin panel API (doPost actions beginning with `admin_`).
// Set the passcode in Project Settings → Script properties as ADMIN_PASSCODE.
// ---------------------------------------------------------------

const ADMIN_TABLES = {
  articles: { sheet: 'published_articles', columns: PUBLISHED_ARTICLE_COLUMNS, idKey: 'article_id', prefix: 'art' },
  quick_wins: { sheet: 'quick_wins', columns: QUICK_WINS_COLUMNS, idKey: 'tip_id', prefix: 'tip' },
  directory: { sheet: 'directory_links', columns: DIRECTORY_LINK_COLUMNS, idKey: 'resource_id', prefix: 'res' },
  quotes: { sheet: 'lounge_quotes', columns: LOUNGE_QUOTE_COLUMNS, idKey: 'quote_id', prefix: 'quote' },
  studio: { sheet: 'peer_studio', columns: PEER_STUDIO_COLUMNS, idKey: 'studio_id', prefix: 'studio' },
  tech_tools: { sheet: 'tech_tools', columns: TECH_TOOL_COLUMNS, idKey: 'tool_id', prefix: 'tool' },
  image_library: { sheet: 'image_library', columns: IMAGE_LIBRARY_COLUMNS, idKey: 'image_id', prefix: 'img' },
  page_sections: { sheet: 'page_sections', columns: PAGE_SECTION_COLUMNS, idKey: 'section_id', prefix: 'sec', customId: true },
  content_blocks: { sheet: 'content_blocks', columns: CONTENT_BLOCK_COLUMNS, idKey: 'block_id', prefix: 'blk' },
  videos: { sheet: 'videos', columns: VIDEO_COLUMNS, idKey: 'video_id', prefix: 'vid', customId: true },
  video_chapters: { sheet: 'video_chapters', columns: VIDEO_CHAPTER_COLUMNS, idKey: 'chapter_id', prefix: 'chap' }
};

const ADMIN_MAX_FAILURES = 5;
const ADMIN_LOCKOUT_SECONDS = 15 * 60;

function adminCheckPasscode_(passcode) {
  const expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSCODE');
  if (!expected) throw new Error('Admin access is not set up. Add an ADMIN_PASSCODE script property.');
  const cache = CacheService.getScriptCache();
  const failures = Number(cache.get('admin_failures') || 0);
  if (failures >= ADMIN_MAX_FAILURES) throw new Error('Too many incorrect passcodes. Try again in 15 minutes.');
  const given = String(passcode || '');
  let diff = given.length === expected.length ? 0 : 1;
  for (let i = 0; i < expected.length; i++) diff |= (given.charCodeAt(i) || 0) ^ expected.charCodeAt(i);
  if (diff) {
    cache.put('admin_failures', String(failures + 1), ADMIN_LOCKOUT_SECONDS);
    throw new Error('Incorrect passcode.');
  }
  cache.remove('admin_failures');
}

function handleAdminAction_(body) {
  try {
    adminCheckPasscode_(body.passcode);
    const action = String(body.action);
    if (action === 'admin_login') return { ok: true };
    if (action === 'admin_load') return adminLoad_();

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(20000)) throw new Error('The sheet is busy. Please try again.');
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (action === 'admin_save_setting') return adminSaveSetting_(ss, body.key, body.value);
      if (action === 'admin_save_row') return adminSaveRow_(ss, body.table, body.id, JSON.parse(body.fields || '{}'), false);
      if (action === 'admin_add_row') return adminSaveRow_(ss, body.table, '', JSON.parse(body.fields || '{}'), true);
      if (action === 'admin_delete_row') return adminDeleteRow_(ss, body.table, body.id);
      if (action === 'admin_set_status') return adminSetStatus_(ss, body.tab, body.id, body.status);
    } finally {
      lock.releaseLock();
    }
    return { ok: false, error: 'Unknown admin action' };
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

function adminLoad_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = { ok: true, settings: [], tables: {}, submissions: [] };

  const settings = ss.getSheetByName('settings');
  if (settings) seedDefaultSettings_(settings);
  if (settings && settings.getLastRow() > 1) {
    const rows = settings.getDataRange().getValues();
    const descIdx = String(rows[0][2] || '').trim().toLowerCase() === 'description' ? 2 : -1;
    for (let r = 1; r < rows.length; r++) {
      const key = String(rows[r][0] || '').trim();
      if (!key) continue;
      const val = rows[r][1];
      out.settings.push({
        key,
        value: val instanceof Date ? val.toISOString() : String(val == null ? '' : val),
        description: descIdx >= 0 ? String(rows[r][descIdx] || '') : ''
      });
    }
  }

  Object.keys(ADMIN_TABLES).forEach(name => {
    const cfg = ADMIN_TABLES[name];
    const sheet = name === 'articles' ? getPublishedArticlesSheet_(ss) : ss.getSheetByName(cfg.sheet);
    out.tables[name] = sheet ? smoreReadIssues_(sheet, smoreHeaders_(sheet, cfg.columns, true)).map(r => {
      const item = {};
      cfg.columns.forEach(c => { item[c] = r.issue[c]; });
      return item;
    }) : [];
  });

  ['submissions', 'review_queue'].forEach(tab => {
    const sheet = ss.getSheetByName(tab);
    if (!sheet) return;
    smoreReadIssues_(sheet, smoreHeaders_(sheet, SUBMISSION_COLUMNS, true)).forEach(r => {
      const item = { tab };
      SUBMISSION_COLUMNS.forEach(c => { item[c] = r.issue[c]; });
      out.submissions.push(item);
    });
  });
  out.submissions.sort((a, b) => String(b.submitted_at).localeCompare(String(a.submitted_at)));
  return out;
}

const SITE_CATEGORIES = ['Student Engagement', 'Classroom Management', 'Teacher Support', 'Tech & AI'];

// Maps typed topics such as "student engagement" or "Tech and AI" to the site's topic names.
function canonicalCategory_(value) {
  const key = String(value || '').toLowerCase().replace(/&/g, 'and').replace(/[^a-z0-9]+/g, '');
  const map = {
    studentengagement: 'Student Engagement', engagement: 'Student Engagement',
    classroommanagement: 'Classroom Management', management: 'Classroom Management',
    teachersupport: 'Teacher Support', teachersupportandwellbeing: 'Teacher Support', wellbeing: 'Teacher Support',
    techandai: 'Tech & AI', techai: 'Tech & AI', techandaitips: 'Tech & AI', tech: 'Tech & AI', ai: 'Tech & AI'
  };
  return map[key] || String(value || '').trim();
}

function adminSafeCell_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+@-]/.test(text) ? "'" + text : text;
}

function adminSaveSetting_(ss, key, value) {
  key = String(key || '').trim();
  const sheet = ss.getSheetByName('settings');
  if (!sheet || !key) throw new Error('Unknown setting.');
  const keys = sheet.getLastRow() ? sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues() : [];
  const idx = keys.findIndex((row, i) => i > 0 && String(row[0] || '').trim() === key);
  if (idx < 0) throw new Error('Unknown setting: ' + key);
  sheet.getRange(idx + 1, 2).setValue(adminSafeCell_(value));
  appendAuditLog_(ss, 'admin_setting_edit', 'settings', key, key, 'Updated from admin panel');
  return { ok: true };
}

function adminTable_(ss, name) {
  const cfg = ADMIN_TABLES[name];
  if (!cfg) throw new Error('Unknown section.');
  const sheet = name === 'articles' ? getPublishedArticlesSheet_(ss) : ss.getSheetByName(cfg.sheet);
  if (!sheet) throw new Error('Sheet tab "' + cfg.sheet + '" is missing. Run setupSiteSheet.');
  return { cfg, sheet, headers: smoreHeaders_(sheet, cfg.columns, true) };
}

function adminSaveRow_(ss, name, id, fields, isNew) {
  const t = adminTable_(ss, name);
  const editable = t.cfg.columns.filter(c => c !== t.cfg.idKey);
  if (isNew) {
    const record = {};
    editable.forEach(c => { if (fields[c] !== undefined) record[c] = fields[c]; });
    if (name === 'articles' && record.category !== undefined) record.category = canonicalCategory_(record.category);
    if (record.is_live === undefined && t.cfg.columns.indexOf('is_live') >= 0) record.is_live = 'TRUE';
    // Sections and videos are referenced by other tabs, so the admin may choose their ID.
    const wanted = t.cfg.customId ? String(fields[t.cfg.idKey] || '').trim().replace(/[^A-Za-z0-9_-]/g, '-') : '';
    if (wanted && liveRowsById_({ sheet: t.sheet, headers: t.headers, idKey: t.cfg.idKey }, wanted).length) {
      throw new Error('The ID "' + wanted + '" is already used. Choose a different one.');
    }
    record[t.cfg.idKey] = wanted || t.cfg.prefix + '_' + Utilities.getUuid().slice(0, 8);
    appendRecordByHeaders_(t.sheet, t.headers, record);
    appendAuditLog_(ss, 'admin_add', t.cfg.sheet, record[t.cfg.idKey], record.title || '', 'Added from admin panel');
    return { ok: true, id: record[t.cfg.idKey] };
  }
  const matches = liveRowsById_({ sheet: t.sheet, headers: t.headers, idKey: t.cfg.idKey }, String(id || ''));
  if (!matches.length) throw new Error('That item no longer exists. Reload the panel.');
  if (name === 'articles' && fields.category !== undefined) fields.category = canonicalCategory_(fields.category);
  matches.forEach(m => {
    editable.forEach(c => {
      if (fields[c] === undefined) return;
      const col = t.headers.indexOf(c) + 1;
      if (col) t.sheet.getRange(m.row, col).setValue(adminSafeCell_(fields[c]));
    });
  });
  appendAuditLog_(ss, 'admin_edit', t.cfg.sheet, id, fields.title || '', 'Edited from admin panel');
  return { ok: true, id: id };
}

function adminDeleteRow_(ss, name, id) {
  const t = adminTable_(ss, name);
  const matches = liveRowsById_({ sheet: t.sheet, headers: t.headers, idKey: t.cfg.idKey }, String(id || ''));
  if (!matches.length) throw new Error('That item no longer exists. Reload the panel.');
  matches.map(m => m.row).sort((a, b) => b - a).forEach(n => t.sheet.deleteRow(n));
  appendAuditLog_(ss, 'admin_delete', t.cfg.sheet, id, '', 'Deleted from admin panel');
  return { ok: true };
}

function adminSetStatus_(ss, tab, id, status) {
  status = String(status || '').toLowerCase();
  if (['pending', 'approved', 'delete'].indexOf(status) < 0) throw new Error('Invalid status.');
  if (['submissions', 'review_queue'].indexOf(tab) < 0) throw new Error('Unknown submissions tab.');
  const sheet = ss.getSheetByName(tab);
  if (!sheet) throw new Error('Missing tab: ' + tab);
  const headers = smoreHeaders_(sheet, SUBMISSION_COLUMNS, true);
  const matches = smoreReadIssues_(sheet, headers).filter(r => r.issue.id === String(id || ''));
  if (!matches.length) throw new Error('Submission not found. Reload the panel.');
  const statusCol = headers.indexOf('status') + 1;
  const reviewedByCol = headers.indexOf('reviewed_by') + 1;
  matches.forEach(m => {
    sheet.getRange(m.row, statusCol).setValue(status);
    if (reviewedByCol) sheet.getRange(m.row, reviewedByCol).setValue('admin panel');
  });
  appendAuditLog_(ss, 'admin_status', tab, id, matches[0].issue.title, 'Status set to ' + status);
  publishApprovedSubmissions();
  return { ok: true };
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
  let hiddenCount = 0;
  let deletedCount = 0;
  const statusCounts = {};

  ['submissions', 'review_queue'].forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet || sheet.getLastRow() < 2) return;
    const headers = smoreHeaders_(sheet, SUBMISSION_COLUMNS, true);
    const statusIdx = headers.indexOf('status');
    const reviewedAtIdx = headers.indexOf('reviewed_at');
    if (statusIdx < 0) return;

    const rows = smoreReadIssues_(sheet, headers);
    const rowsToDelete = [];
    rows.forEach(r => {
      const status = String(r.issue.status || '').toLowerCase();
      const statusKey = tabName + ': ' + (status || '(blank)');
      statusCounts[statusKey] = (statusCounts[statusKey] || 0) + 1;
      const type = String(r.issue.submission_type || 'article').toLowerCase();

      if (status === 'delete' || status === 'pending') {
        if (!r.issue.id) {
          if (status === 'delete') rowsToDelete.push(r.row);
          return;
        }
        const target = liveTarget_(ss, type);
        if (!target.sheet) return;
        const live = liveRowsById_(target, r.issue.id);
        if (status === 'delete') {
          live.map(l => l.row).sort((a, b) => b - a).forEach(n => target.sheet.deleteRow(n));
          if (live.length) appendAuditLog_(ss, 'deleted_live_item', tabName, r.issue.id, r.issue.title, 'Removed from live tab');
          rowsToDelete.push(r.row);
          deletedCount++;
        } else {
          const liveCol = target.headers.indexOf('is_live') + 1;
          live.forEach(l => {
            if (liveCol && String(l.issue.is_live).toLowerCase() !== 'false') {
              target.sheet.getRange(l.row, liveCol).setValue('FALSE');
              hiddenCount++;
            }
          });
        }
        return;
      }

      if (status !== 'approved' && status !== 'publish') return;

      const id = r.issue.id || ('item_' + Utilities.getUuid().slice(0, 8));
      const existing = r.issue.id ? liveRowsById_(liveTarget_(ss, type), id) : [];
      if (existing.length) {
        const target = liveTarget_(ss, type);
        const liveCol = target.headers.indexOf('is_live') + 1;
        if (liveCol) existing.forEach(l => target.sheet.getRange(l.row, liveCol).setValue('TRUE'));
        sheet.getRange(r.row, statusIdx + 1).setValue('published');
        if (reviewedAtIdx >= 0) sheet.getRange(r.row, reviewedAtIdx + 1).setValue(now);
        publishedCount++;
        return;
      }
      const title = r.issue.title || 'Staff Contribution';
      const content = r.issue.content || '';
      const author = r.issue.submitter_name || 'MMS Staff';
      const category = canonicalCategory_(r.issue.category) || 'Teacher Support';
      const links = r.issue.links || '';

      if (type === 'quote' || type === 'lounge_quote') {
        const qSheet = ss.getSheetByName('lounge_quotes');
        const qHeaders = smoreHeaders_(qSheet, LOUNGE_QUOTE_COLUMNS, true);
        appendRecordByHeaders_(qSheet, qHeaders, {
          quote_id: id,
          quote: content,
          attribution: author + (title && title !== 'Staff Lounge Post' ? ' (' + title + ')' : ''),
          is_live: 'TRUE',
          post_type: ['win', 'quote', 'question'].indexOf(String(category).toLowerCase()) >= 0 ? String(category).toLowerCase() : 'quote',
          appreciate: 0,
          empathy: 0,
          respect: 0
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
          cover_image_url: r.issue.cover_image_url || '',
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
    rowsToDelete.sort((a, b) => b - a).forEach(n => sheet.deleteRow(n));
  });

  try {
    let msg = 'Published ' + publishedCount + ' approved submission(s).';
    if (hiddenCount) msg += '\nHid ' + hiddenCount + ' item(s) set back to pending.';
    if (deletedCount) msg += '\nDeleted ' + deletedCount + ' item(s) marked delete.';
    if (!publishedCount && !hiddenCount && !deletedCount) {
      const found = Object.keys(statusCounts).map(k => k + ' x' + statusCounts[k]).join('\n');
      msg += '\n\nOnly rows with status "approved" are published. Rows found:\n' + (found || 'none in submissions or review_queue');
    }
    SpreadsheetApp.getUi().alert(msg);
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
    ['smore_source_url', 'https://app.smore.com/n/rsn15', 'Public Smore profile or issue URL to sync'],
    ['site_title', 'Eagle Deep Dive Daily', 'Main publication title in the top header and footer'],
    ['site_badge', 'MMS Eagles', 'Pill badge next to the main title'],
    ['site_subtitle', 'Maize Middle School • Classroom Craft, Student Engagement, Teacher Support & Tech', 'Subtitle under the main title'],
    ['site_tagline', 'Built by teachers, for teachers', 'Tagline on the right of the topic bar'],
    ['hero_article_id', 'art-1', 'article_id of the featured cover story on the home feed'],
    ['hero_badge', 'Cover Story • Student Engagement', 'Badge shown on the cover story image'],
    ['hero_button_label', 'Read Full Story & Script', 'Button text on the cover story card'],
    ['feed_heading_all', 'Latest Articles & Classroom Playbooks', 'Main heading above the articles grid'],
    ['article_search_heading', 'Find an Article', 'Sidebar article search title'],
    ['article_search_subtitle', 'Search published articles by keyword', 'Sidebar article search subtitle'],
    ['article_search_intro', 'Enter a topic or a few keywords, or choose a shortcut below. Minor typos are OK.', 'Sidebar article search description'],
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
    ['academy_nav_title', 'EdTech, AI & Video Lab', 'Name of the Tech & AI page in the menu'],
    ['academy_nav_desc', '', 'Short description under the Tech & AI menu name'],
    ['page_empty_message', 'Nothing has been published on this page yet.', 'Shown when a page has no published sections. Leave blank to show nothing.'],
    ['reader_video_badge', 'Companion video', 'Badge on articles marked has_video_demo when a video is published'],
    ['reader_video_button', 'Open the video page', 'Button on that callout'],
    ['reader_video_note', '', 'Optional note on that callout'],
    ['footer_left', 'Classroom Craft, Student Engagement, Teacher Support & Tech/AI', 'Footer description text'],
    ['footer_right', 'Maize Middle School (USD 266)', 'Footer school name text'],
    ['loading_messages', 'Burnout prevention protocol enabled.\nEmergency coffee reserves detected.\nThe copier has agreed to cooperate. For now.\nGrading pile temporarily classified as “later.”\nOne more deep breath has been scheduled.\nTiny classroom wins loading. No committee required.\nYour inbox has been asked to calm down.\nClassroom chaos contained. Please do not startle it.\nSurvival mode activated.\nThe laminator is warming up. Everyone stay calm.\nTeacher powers activate!\nHave you tried turning it off and on again?\nMaybe there\'s snacks in the lounge?\n6-7 has been reclaimed.\nLow key lit unc bootloader\nSearching for peak rizz\nSearching databases for "aura"\nNo crash outs detected.', 'Messages shown while the site loads, one per line. New messages appear from a visitor\'s next visit.'],
    ['menu_title', 'Site Navigation Menu', 'Small heading at the top of the site menu'],
    ['menu_badge', 'Eagles Hub', 'Red badge in the site menu header'],
    ['menu_home_title', 'Home: Main Editorial Feed', 'Menu item: Home title'],
    ['menu_home_desc', 'All featured stories, playbooks & quick wins', 'Menu item: Home description'],
    ['menu_section_heading', 'Building Wings & Special Sections', 'Menu heading above the special page links'],
    ['menu_directory_title', 'Heart of MMS Directory', 'Menu item: Directory title'],
    ['menu_directory_desc', 'Clean, searchable building links & forms', 'Menu item: Directory description'],
    ['menu_bulletin_title', 'Flight Wire (School Updates)', 'Menu item: Flight Wire title'],
    ['menu_bulletin_desc', 'News, social posts, calendars, and newsletter archive', 'Menu item: Flight Wire description'],
    ['menu_studio_title', 'The Peer Lesson Studio', 'Menu item: Peer Lesson Studio title'],
    ['menu_studio_desc', 'Warm & Cool feedback lab for lesson drafts', 'Menu item: Peer Lesson Studio description'],
    ['menu_topics_heading', 'Editorial Topics', 'Menu heading above the topic shortcuts'],
    ['menu_topic_engagement', 'Engagement', 'Menu topic shortcut: Student Engagement'],
    ['menu_topic_management', 'Management', 'Menu topic shortcut: Classroom Management'],
    ['menu_topic_support', 'Teacher Support', 'Menu topic shortcut: Teacher Support'],
    ['menu_topic_tech', 'Tech & AI', 'Menu topic shortcut: Tech & AI'],
    ['menu_share_tip', 'Share a Tip', 'Menu button: Share a Tip'],
    ['menu_settings', 'Settings', 'Menu button: Settings'],
    ['search_button_label', 'Search', 'Label on the top search bar button'],
    ['topbar_all', 'All Featured & Latest', 'Topic bar: All'],
    ['topbar_engagement', 'Student Engagement', 'Topic bar: Student Engagement'],
    ['topbar_management', 'Classroom Management', 'Topic bar: Classroom Management'],
    ['topbar_support', 'Teacher Support & Wellbeing', 'Topic bar: Teacher Support'],
    ['topbar_tech', 'Tech & AI Tips', 'Topic bar: Tech & AI'],
    ['feed_saved_note', 'Saved articles are stored in this browser. Clearing browser data/cache will remove your saved items.', 'Note shown on the Saved Reading List'],
    ['article_search_button', 'Search Articles', 'Sidebar article search button'],
    ['lounge_view_all_button', 'Visit the Staff Lounge →', 'Sidebar button linking to the Staff Lounge'],
    ['lounge_add_button', '+ Post a Classroom Quote or Win', 'Sidebar button to post to the lounge'],
    ['back_to_feed_label', 'Back to Editorial Feed', 'Back-to-feed button label (used on every page)'],
    ['directory_footnote', 'Some resources require your USD 266 work account.', 'Note under the directory'],
    ['lounge_page_heading', 'Staff Lounge', 'Staff Lounge page heading'],
    ['lounge_page_subtitle', 'Share a classroom win, a funny quote, or a question. React to a colleague\'s post with appreciation, empathy, or respect.', 'Staff Lounge page description'],
    ['lounge_new_post_button', '+ New Post', 'Staff Lounge: new post button'],
    ['lounge_filter_all', 'All', 'Staff Lounge filter: All'],
    ['lounge_filter_wins', 'Wins', 'Staff Lounge filter: Wins'],
    ['lounge_filter_quotes', 'Quotes', 'Staff Lounge filter: Quotes'],
    ['lounge_filter_questions', 'Questions', 'Staff Lounge filter: Questions'],
    ['flight_page_heading', 'Flight Wire: School Updates', 'Flight Wire page heading'],
    ['flight_page_subtitle', 'MMS updates first, with district posts and Smore newsletters in their own tabs.', 'Flight Wire page description'],
    ['flight_tab_mms', 'MMS Updates', 'Flight Wire tab: MMS'],
    ['flight_tab_district', 'District Updates', 'Flight Wire tab: District'],
    ['flight_tab_smore', 'Smore Archive', 'Flight Wire tab: Smore'],
    ['flight_older_note', 'Updates older than seven days are tucked away; items without a date remain visible.', 'Flight Wire note about older updates'],
    ['flight_news_heading', 'News', 'Flight Wire: News heading'],
    ['flight_news_note', 'Headlines link to the complete articles on the school website.', 'Flight Wire: News note'],
    ['flight_social_heading', 'Social', 'Flight Wire: Social heading'],
    ['flight_calendars_heading', 'Calendars', 'Flight Wire: Calendars heading'],
    ['flight_calendars_note', 'These open the official calendars; events are not imported into this site.', 'Flight Wire: Calendars note'],
    ['flight_district_news_heading', 'District News', 'Flight Wire: District News heading'],
    ['flight_district_social_heading', 'District Social', 'Flight Wire: District Social heading'],
    ['flight_district_note', 'District and other account updates are kept here rather than mixed into MMS Updates.', 'Flight Wire: District note'],
    ['studio_form_heading', 'Post a Lesson Draft', 'Peer Lesson Studio form heading'],
    ['studio_label_title', 'Lesson / Activity Title', 'Peer Lesson Studio form label'],
    ['studio_label_subject', 'Department & Grade', 'Peer Lesson Studio form label'],
    ['studio_label_dilemma', 'Feedback Focus Dilemma', 'Peer Lesson Studio form label'],
    ['studio_label_link', 'Google Doc / Slides Link', 'Peer Lesson Studio form label'],
    ['studio_submit_button', 'Submit for Review', 'Peer Lesson Studio submit button'],
    ['article_try_tomorrow_label', 'Try This Tomorrow in Class', 'Article reader heading: Try This Tomorrow'],
    ['article_script_label', 'Ready-to-Use Classroom Script / Prompt', 'Article reader heading: Script / Prompt'],
    ['article_share_note', 'Found this helpful? Share it with your PLC or pin it to your reading list.', 'Article reader footer note'],
    ['article_done_button', 'Done Reading', 'Article reader done button'],
    ['tip_modal_heading', 'Share a Classroom Tip or Win', 'Share a Tip form heading'],
    ['tip_modal_subtitle', 'Contribute a strategy, Tech/AI workflow, engagement hook, or student quote', 'Share a Tip form description'],
    ['tip_submit_button', 'Submit for Review', 'Share a Tip submit button'],
    ['lounge_modal_heading', 'Post to the Staff Lounge', 'Lounge post form heading'],
    ['lounge_modal_subtitle', 'Share a classroom win, a quote you overheard, or a question. Posts are reviewed before they appear.', 'Lounge post form description'],
    ['lounge_submit_button', 'Post', 'Lounge post submit button']
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

function unpublishedSamples_(records) {
  return records.map(record => Object.assign({}, record, { is_live: 'FALSE' }));
}

// Built-in demonstration rows were seeded as live in earlier versions. Hide them without deleting.
function unpublishSampleRows() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.alert(
    'Unpublish sample rows',
    'Sets is_live to FALSE on rows that still use the original sample IDs (for example art-1, qw-1, tool-1). ' +
    'Rows you edited but kept the same ID are hidden too; set is_live back to TRUE to republish them. Continue?',
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targets = [
    [getPublishedArticlesSheet_(ss), 'article_id', defaultArticlesSeed_()],
    [ss.getSheetByName('quick_wins'), 'tip_id', defaultQuickWinsSeed_()],
    [ss.getSheetByName('directory_links'), 'resource_id', defaultDirectorySeed_()],
    [ss.getSheetByName('lounge_quotes'), 'quote_id', defaultQuotesSeed_()],
    [ss.getSheetByName('peer_studio'), 'studio_id', defaultStudioSeed_()],
    [ss.getSheetByName('tech_tools'), 'tool_id', defaultTechToolsSeed_()]
  ];
  const summary = [];
  targets.forEach(([sheet, idKey, seed]) => {
    if (!sheet || sheet.getLastRow() < 2) return;
    const ids = new Set(seed.map(record => String(record[idKey])));
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim());
    const idCol = headers.indexOf(idKey);
    const liveCol = headers.indexOf('is_live');
    if (idCol < 0 || liveCol < 0) return;
    let changed = 0;
    for (let r = 1; r < values.length; r++) {
      const live = values[r][liveCol];
      if (ids.has(String(values[r][idCol]).trim()) && (live === true || String(live).trim().toLowerCase() === 'true')) {
        sheet.getRange(r + 1, liveCol + 1).setValue('FALSE');
        changed++;
      }
    }
    if (changed) {
      summary.push(sheet.getName() + ': ' + changed);
      appendAuditLog_(ss, 'unpublish_samples', sheet.getName(), '', '', changed + ' sample rows set to FALSE');
    }
  });
  ui.alert(summary.length ? 'Unpublished rows\n' + summary.join('\n') : 'No live sample rows found.');
}

// Tech & AI page structure. Sections are live; they stay invisible until they have published content.
function defaultPageSectionsSeed_() {
  return [
    { section_id: 'tech-video', page: 'tech_ai', kind: 'videos', layout: 'stack', title: '', intro: '', badge: '', sort_order: 10, is_live: 'TRUE' },
    { section_id: 'tech-prompts', page: 'tech_ai', kind: 'blocks', layout: 'columns', title: 'Copy-Paste NotebookLM Workflows', intro: 'Upload your unit slides, lab sheets, or standards PDF into NotebookLM, then copy a prompt.', badge: 'Prompt Recipes', sort_order: 20, is_live: 'TRUE' },
    { section_id: 'tech-playbook', page: 'tech_ai', kind: 'blocks', layout: 'columns', title: 'Classroom Video Playbook', intro: 'Ideas for teacher-made and student-created videos.', badge: 'Under 3 Minutes', sort_order: 30, is_live: 'TRUE' },
    { section_id: 'tech-tools', page: 'tech_ai', kind: 'tools', layout: 'stack', title: 'Essential Classroom Tech Toolkit', intro: 'Tools that work smoothly on school Chromebooks.', badge: '', sort_order: 40, is_live: 'TRUE' }
  ];
}

function defaultContentBlocksSeed_() {
  const prompt = (id, order, title, description, body) => ({
    block_id: id, section_id: 'tech-prompts', block_type: 'prompt', title: title,
    description: description, link_label: 'Copy', style: 'neutral', body: body, sort_order: order
  });
  return [
    prompt('prompt-catch-up', 10, 'Missing-Student Catch-Up Slide & Video Script',
      'Generates a 5-slide summary, 90-sec narration script, and Google Classroom makeup check.',
      'Using only the uploaded lesson slides and lab handout from today, create a \'What You Missed\' catch-up package for an absent 8th grader: 1) A 5-slide outline they can read in 2 minutes, 2) A 90-second video/audio narration script summarizing the main phenomenon, and 3) Two quick check questions they can submit in Google Classroom for makeup credit.'),
    prompt('prompt-study-guide', 20, 'Google Classroom Study Guide & FAQ Builder',
      'Builds a clean unit review doc ready to post to your Google Classroom stream.',
      'Create a student-friendly Google Classroom Study Guide from our uploaded unit notes. Include: 1) Top 5 Big Ideas in plain middle-school language, 2) A glossary of 6 Tier 3 vocabulary terms with everyday analogies, and 3) 4 common misconceptions students should watch out for on Friday\'s quiz.'),
    prompt('prompt-tiered-reading', 30, 'Tiered Reading Passages (Lexile 700 / 950 / 1150)',
      'Differentiates complex texts while keeping the same discussion questions.',
      'Using only the uploaded unit standards and background article, write 3 versions of an engaging 250-word reading passage explaining this phenomenon at Lexile 700, 950, and 1150, followed by 3 identical discussion questions so all students can join the same table debate.'),
    prompt('prompt-listening-station', 40, 'Audio Overview Listening Station Sheet',
      'Pairs with NotebookLM\'s 2-host podcast generator for active listening stations.',
      'Create a 1-page student station guide with 4 key vocabulary look-fors, 2 diagram sketching boxes, and 3 pause-and-predict questions to accompany the NotebookLM Audio Overview for this unit.'),
    { block_id: 'playbook-teacher', section_id: 'tech-playbook', block_type: 'list', title: 'Making Short, Content-Focused Videos for Your Students',
      body: 'The Lab Setup & Safety Loop: record a 90-second walkthrough of equipment setup and post it in Google Classroom.\nThe Worked-Example Replay: record yourself solving one tough problem so students or parents can replay it at home.\nMissing-Student Slide Narration: record a 2-minute voiceover over a catch-up outline for absent students.',
      description: '', url: '', link_label: '', style: 'numbered', sort_order: 10 },
    { block_id: 'playbook-student', section_id: 'tech-playbook', block_type: 'list', title: 'Having Students Make Videos to Prove Mastery',
      body: 'Box 1 (0-15s) The Hook & Claim: state the mystery question and your team\'s claim.\nBox 2 (15-60s) Visual Diagram Proof: point to your model or graph and explain why using 3 unit vocabulary words.\nBox 3 (60-90s) Bust the Myth: explain the #1 common trap and why the evidence disproves it.',
      description: '', url: '', link_label: '', style: 'bullets', sort_order: 20 },
    { block_id: 'playbook-tip', section_id: 'tech-playbook', block_type: 'callout', title: 'Pro Tip',
      body: 'Never re-record over a tiny verbal stumble. Authentic one-take videos feel like real teaching.',
      description: '', url: '', link_label: '', style: 'green', sort_order: 30 },
    { block_id: 'playbook-student-rubric', section_id: 'tech-playbook', block_type: 'prompt', title: 'Student Video Directions & 3-Box Rubric',
      body: 'Student 90-Second Video Rubric & Directions: 1) Complete your 3-Box Paper Storyboard first and get a teacher stamp. 2) Record a 60-90 second One-Take Video (Paper Slide or Chromebook Screencast). 3) Must include: Visual diagram on screen, all 3 target vocabulary words used in context, and 1 busted misconception.',
      description: '', url: '', link_label: 'Copy directions', style: 'neutral', sort_order: 40 }
  ];
}

// Sample video: add a real YouTube/Vimeo/Loom/Drive URL, then set is_live to TRUE on the video and its chapters.
function defaultVideosSeed_() {
  return [{
    video_id: 'sample-video', section_id: 'tech-video', badge: 'Staff Walkthrough Video', title: 'Sample video title',
    description: 'Replace this row with a real video. With no live video row there is no video panel.',
    url: '', link_label: '', sort_order: 10
  }];
}

function defaultVideoChaptersSeed_() {
  return [
    { chapter_id: 'sample-chapter-1', video_id: 'sample-video', title: 'First chapter', description: 'What this part of the video covers.', start_time: '0:00', sort_order: 10 },
    { chapter_id: 'sample-chapter-2', video_id: 'sample-video', title: 'Second chapter', description: 'Use the real timestamp from the video.', start_time: '1:25', sort_order: 20 }
  ];
}
