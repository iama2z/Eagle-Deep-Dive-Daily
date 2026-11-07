const SCHOOL_UPDATE_COLUMNS = ['id', 'title', 'text', 'url', 'published_at', 'source'];
const SCHOOL_UPDATE_SOURCES = {
  news: 'https://mms.usd266.com/about-us/news',
  social: 'https://www.juicer.io/api/feeds/378904?filter=783230,778997'
};

function syncSchoolUpdates() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) throw new Error('Another site sync is running.');
  const failures = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(SCHOOL_UPDATE_SOURCES).forEach(source => {
      try {
        const response = UrlFetchApp.fetch(SCHOOL_UPDATE_SOURCES[source], {
          muteHttpExceptions: true, followRedirects: true
        });
        if (response.getResponseCode() !== 200) throw new Error('HTTP ' + response.getResponseCode());
        const items = source === 'news'
          ? schoolNewsItems_(response.getContentText())
          : schoolSocialItems_(JSON.parse(response.getContentText()));
        if (!items.length) throw new Error('No recognizable entries found; previous entries retained.');
        const sheet = ss.getSheetByName('school_' + source) || ss.insertSheet('school_' + source);
        const headers = smoreHeaders_(sheet, SCHOOL_UPDATE_COLUMNS, true);
        const rows = sheet.getLastRow() > 1
          ? sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues() : [];
        const ids = new Map(rows.map((row, index) => [String(row[headers.indexOf('id')]), index + 2]));
        items.forEach(item => {
          const rowNumber = ids.get(item.id) || sheet.getLastRow() + 1;
          const row = ids.has(item.id) ? rows[rowNumber - 2].slice() : new Array(headers.length).fill('');
          SCHOOL_UPDATE_COLUMNS.forEach(key => {
            const value = String(item[key] || '');
            row[headers.indexOf(key)] = /^[=+\-@]/.test(value) ? "'" + value : value;
          });
          sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
        });
        console.log(source + ': imported ' + items.length + ' entries.');
      } catch (err) {
        console.error(source + ': ' + err.message);
        failures.push(source + ': ' + err.message);
      }
    });
    if (failures.length) throw new Error(failures.join('; '));
  } finally {
    lock.releaseLock();
  }
}

function schoolUpdateText_(html) {
  return String(html || '').replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '')
    .replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '')
    .replace(/<(?:br|\/p|\/div)\b[^>]*>/gi, '\n')
    .replace(/<[^>]*>/g, '').replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&').replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'").replace(/&lt;/gi, '<').replace(/&gt;/gi, '>')
    .replace(/&#(x[0-9a-f]+|\d+);/gi, (match, code) => {
      const number = code[0].toLowerCase() === 'x' ? parseInt(code.slice(1), 16) : Number(code);
      return number > 0 && number <= 0x10ffff ? String.fromCodePoint(number) : match;
    }).replace(/[ \t]+/g, ' ').replace(/\n\s*\n/g, '\n\n').trim();
}

function schoolNewsItems_(html) {
  const items = new Map();
  for (const match of html.matchAll(/<a\b[^>]*href=["']([^"']+)["'][^>]*>([\s\S]*?)<\/a>/gi)) {
    const href = match[1].replace(/&amp;/g, '&');
    const url = href.startsWith('/') ? 'https://mms.usd266.com' + href : href;
    if (!/^https:\/\/mms\.usd266\.com\/about-us\/news\/news-details\/~board\/(?:maize-middle-school-news|district-news)\/post\/[a-z0-9-]+\/?$/i.test(url)) continue;
    const title = schoolUpdateText_(match[2]);
    if (!title || title.length > 400) continue;
    items.set(url, { id: url, title: title, text: '', url: url, published_at: '',
      source: url.includes('/district-news/') ? 'District News' : 'MMS News' });
  }
  return [...items.values()];
}

function schoolSocialItems_(data) {
  if (!data || !data.posts || !Array.isArray(data.posts.items)) throw new Error('Missing Juicer posts.items.');
  return data.posts.items.filter(post => post && post.source &&
    [783230, 778997].includes(Number(post.source.id)) && !post.deleted_at).map(post => {
    const text = schoolUpdateText_(post.unformatted_message || post.message || post.description);
    const date = Date.parse(post.external_created_at);
    return { id: String(post.id || ''), title: String(post.source.name || 'School Social'),
      text: text, url: /^https?:\/\//i.test(String(post.external || '')) ? String(post.external) : '',
      published_at: Number.isFinite(date) ? new Date(date).toISOString() : '',
      source: String(post.source.name || 'School Social') };
  }).filter(item => item.id && item.text);
}

function publicSchoolUpdates_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = { news: [], social: [] };
  Object.keys(result).forEach(source => {
    const sheet = ss.getSheetByName('school_' + source);
    if (!sheet || sheet.getLastRow() < 2) return;
    const headers = smoreHeaders_(sheet, SCHOOL_UPDATE_COLUMNS, false);
    result[source] = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues()
      .map(row => Object.fromEntries(SCHOOL_UPDATE_COLUMNS.map(key => {
        const value = row[headers.indexOf(key)];
        return [key, value instanceof Date ? value.toISOString() : String(value || '')];
      }))).filter(item => item.id && item.title)
      .sort((a, b) => b.published_at.localeCompare(a.published_at)).slice(0, 100);
  });
  return result;
}
