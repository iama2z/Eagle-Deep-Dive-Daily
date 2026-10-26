// Bound to the newsletter spreadsheet. Keep the existing submissions doPost.
const SMORE_COLUMNS = [
  'issue_url', 'issue_title', 'published_at', 'summary',
  'cover_image_url', 'article_body_text'
];

// Utility pages on smore.com that look like issue slugs but are never newsletters.
const SMORE_RESERVED_PATHS = /^(?:about|account|app|blog|business|careers|contact|dashboard|discover|educators|explore|features|gallery|help|jobs|legal|library|login|logout|newsletters|press|pricing|privacy|schools|search|settings|signup|support|templates|terms)$/i;

function setupSmoreFeed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['smore_feed_raw', 'smore_feed_public'].forEach(name => {
    smoreHeaders_(ss.getSheetByName(name) || ss.insertSheet(name), SMORE_COLUMNS, true);
  });
  const settings = ss.getSheetByName('settings') || ss.insertSheet('settings');
  if (!settings.getLastRow()) settings.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
  if (!smoreSourceUrl_(settings)) {
    settings.appendRow(['smore_source_url', 'https://www.smore.com/u/maizemiddleschool']);
  }
}

function smoreSourceUrl_(settings) {
  // `source_url` is the key used by the older sync script; keep honoring it.
  const rows = settings.getDataRange().getValues();
  const row = rows.find(r => String(r[0]).trim() === 'smore_source_url') ||
    rows.find(r => String(r[0]).trim() === 'source_url');
  return String(row && row[1] || '').trim();
}

function syncSmoreFeed() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another newsletter sync is running');
  try {
    setupSmoreFeed();
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

// Replace the existing doGet; do not add a second doGet.
function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.feed === 'articles') {
      return smoreJson_({ ok: true, items: publishedArticles_() });
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('smore_feed_public');
    if (!sheet) throw new Error('Run setupSmoreFeed first');
    const seen = new Set();
    const items = smoreReadIssues_(sheet, smoreHeaders_(sheet, SMORE_COLUMNS, false)).map(row => row.issue)
      .filter(issue => {
        const url = smoreIssueUrl_(issue.issue_url);
        if (!url || seen.has(url) || !issue.issue_title || !issue.article_body_text ||
          String(issue.is_live || '').toLowerCase() === 'false') return false;
        seen.add(url);
        issue.issue_url = url;
        return true;
      })
      .map(issue => Object.fromEntries(SMORE_COLUMNS.map(key => [key, issue[key] || ''])));
    return smoreJson_({ ok: true, items });
  } catch (err) {
    return smoreJson_({ ok: false, error: String(err.message) });
  }
}

function publishedArticles_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('published_articles');
  if (!sheet) return [];
  const [headers, ...rows] = sheet.getDataRange().getValues();
  const columns = [
    'article_id', 'title', 'category', 'summary', 'article_body_text',
    'author', 'published_at', 'cover_image_url', 'is_live'
  ];
  const required = ['article_id', 'title', 'category', 'article_body_text', 'is_live'];
  if (!headers || required.some(key => !headers.includes(key))) {
    throw new Error('Missing required published_articles headers: ' + required.join(', '));
  }
  return rows.filter(row =>
    String(row[headers.indexOf('is_live')]).trim().toLowerCase() === 'true'
  ).map(row => Object.fromEntries(columns.map(key => {
    const value = row[headers.indexOf(key)];
    return [key, key === 'is_live' ? true :
      value instanceof Date ? value.toISOString() : String(value == null ? '' : value).trim()];
  }))).filter(item => required.every(key => item[key]));
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
