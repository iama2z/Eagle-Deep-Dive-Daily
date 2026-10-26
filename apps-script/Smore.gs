// Bound to the newsletter spreadsheet. Keep the existing submissions doPost.
const SMORE_COLUMNS = [
  'issue_url', 'issue_title', 'published_at', 'summary',
  'cover_image_url', 'article_body_text'
];

function setupSmoreFeed() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ['smore_feed_raw', 'smore_feed_public'].forEach(name => {
    const sheet = ss.getSheetByName(name) || ss.insertSheet(name);
    smoreCheckColumns_(sheet, SMORE_COLUMNS);
  });
  const settings = ss.getSheetByName('settings') || ss.insertSheet('settings');
  smoreCheckColumns_(settings, ['key', 'value']);
  const rows = settings.getDataRange().getValues();
  if (!rows.some(row => row[0] === 'smore_source_url')) {
    settings.appendRow(['smore_source_url', 'https://www.smore.com/u/maizemiddleschool']);
  }
}

function syncSmoreFeed() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another newsletter sync is running');
  try {
    setupSmoreFeed();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const raw = ss.getSheetByName('smore_feed_raw');
    const published = ss.getSheetByName('smore_feed_public');
    const issues = new Map();
    const publicIssues = new Map();
    [published, raw].forEach(sheet => {
      smoreReadIssues_(sheet).forEach(issue => {
        const url = smoreIssueUrl_(issue.issue_url);
        if (!url) return;
        issue.issue_url = url;
        if (sheet === published) publicIssues.set(url, issue);
        issues.set(url, issue);
      });
    });
    const settings = ss.getSheetByName('settings').getDataRange().getValues();
    const source = settings.find(row => row[0] === 'smore_source_url');
    try {
      const sourceUrl = String(source && source[1] || '').trim();
      if (smoreIssueUrl_(sourceUrl)) {
        if (!issues.has(smoreIssueUrl_(sourceUrl))) issues.set(smoreIssueUrl_(sourceUrl), {});
      } else {
        if (!/^https:\/\/(?:www\.)?smore\.com\/u\/[a-z0-9_-]+\/?$/i.test(sourceUrl)) {
          throw new Error('smore_source_url must be a public Smore profile or issue URL');
        }
        const html = smoreFetch_(sourceUrl);
        for (const match of html.matchAll(/\bhref\s*=\s*["']([^"']+)["']/gi)) {
          const link = smoreDecode_(match[1]);
          const url = smoreIssueUrl_(link.startsWith('/') ? 'https://www.smore.com' + link : link);
          if (url && !issues.has(url)) issues.set(url, {});
        }
      }
    } catch (err) {
      console.warn('Discovery failed; retrying known issue URLs. ' + err.message);
    }
    if (!issues.size) throw new Error('No issue URLs found. Add a public issue URL to smore_feed_raw column A.');
    let refreshed = 0;
    const failures = [];
    for (const [url, previous] of issues) {
      try {
        const issue = smoreParseIssue_(smoreFetch_(url), url);
        issues.set(url, issue);
        publicIssues.set(url, issue);
        refreshed++;
      } catch (err) {
        // Retain the last successful article, never replace it with a link card.
        issues.set(url, Object.assign({}, previous, { issue_url: url }));
        failures.push(url + ': ' + err.message);
      }
    }
    smoreWriteIssues_(raw, [...issues.values()]);
    smoreWriteIssues_(published, [...publicIssues.values()]);
    console.log(JSON.stringify({ refreshed, retained: publicIssues.size, failures }));
    if (failures.length) throw new Error('Some issues could not refresh; previous content retained. See execution log.');
  } finally {
    lock.releaseLock();
  }
}

// Replace the existing doGet; do not add a second doGet.
function doGet() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('smore_feed_public');
    if (!sheet) throw new Error('Run setupSmoreFeed first');
    smoreCheckColumns_(sheet, SMORE_COLUMNS);
    const items = smoreReadIssues_(sheet).filter(issue =>
      smoreIssueUrl_(issue.issue_url) && issue.issue_title && issue.article_body_text
    );
    return smoreJson_({ ok: true, items });
  } catch (err) {
    return smoreJson_({ ok: false, error: String(err.message) });
  }
}

function smoreJson_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function smoreCheckColumns_(sheet, columns) {
  if (!sheet.getLastRow()) sheet.getRange(1, 1, 1, columns.length).setValues([columns]);
  const actual = sheet.getRange(1, 1, 1, columns.length).getValues()[0];
  if (actual.join('\t') !== columns.join('\t')) {
    throw new Error('Unexpected headers in ' + sheet.getName() + '; back up and migrate this tab first');
  }
}

function smoreReadIssues_(sheet) {
  return sheet.getDataRange().getValues().slice(1).map(row =>
    Object.fromEntries(SMORE_COLUMNS.map((key, i) => [
      key, row[i] instanceof Date ? row[i].toISOString() : String(row[i] == null ? '' : row[i])
    ]))
  );
}

function smoreWriteIssues_(sheet, issues) {
  // A single range update avoids clearing good data before its replacement is ready.
  const rows = issues.map(issue => SMORE_COLUMNS.map(key => {
    const text = String(issue[key] || '');
    return /^[=+@-]/.test(text) ? "'" + text : text;
  }));
  while (rows.length < sheet.getLastRow() - 1) rows.push(SMORE_COLUMNS.map(() => ''));
  if (rows.length) sheet.getRange(2, 1, rows.length, SMORE_COLUMNS.length).setValues(rows);
}

function smoreIssueUrl_(value) {
  const match = String(value || '').trim().match(
    /^https:\/\/(?:www\.)?smore\.com\/((?:n\/)?[a-z0-9]{5,}(?:-[a-z0-9-]+)?)\/?(?:[?#].*)?$/i
  );
  if (!match || /^(?:about|login|signup|pricing|educators|newsletters|privacy|contact|terms)$/i.test(match[1])) return '';
  return 'https://www.smore.com/' + match[1];
}

function smoreFetch_(url) {
  // Never follow a remote redirect off Smore (or into a private network).
  for (let i = 0; i < 4; i++) {
    if (!/^https:\/\/(?:www\.)?smore\.com\//i.test(url)) throw new Error('Only HTTPS Smore URLs are allowed');
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
  if (!title || body.length < 80) throw new Error('No readable issue body found; source may require JavaScript or authentication');
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
    .replace(/<(script|style|noscript|nav|footer|button|iframe|svg)\b[^>]*>[\s\S]*?<\/\1\s*>/gi, '')
    .replace(/<a\b([^>]*)>([\s\S]*?)<\/a\s*>/gi, (_, attrs, label) => {
      const href = smoreDecode_(smoreAttributes_(attrs).href || '');
      return label + (/^https?:\/\/[^\s<>"]+$/i.test(href) ? ' (' + href + ')' : '');
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
