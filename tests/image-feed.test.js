const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const page = fs.readFileSync('index.html', 'utf8');
const imageCode = page.slice(page.indexOf('    const IMG_ENGAGEMENT'), page.indexOf('    // Revolving Search Suggestions'));
const articleCode = page.slice(page.indexOf('    function normalizePublishedArticle'), page.indexOf('    async function loadPublishedArticles'));

function frontend() {
  const context = vm.createContext({ URL, console, Event: class {},
    formatDateLabel: () => '', initFuse: () => {}, renderArticlesGrid: () => {},
    document: { getElementById: () => null }, activeCategory: 'all' });
  vm.runInContext(`${imageCode}
    let siteSettings = {};
    const DEFAULT_ARTICLES_SNAPSHOT = [{id: 'default'}, {id: 'other'}];
    const APP_DATA = {articles: DEFAULT_ARTICLES_SNAPSHOT.slice()};
    ${articleCode}
    globalThis.api = {
      resolveImageUrl, normalizeImageSource, normalizePublishedArticle, applyPublishedArticles, setArticleImage,
      setLibrary: rows => { approvedImageLibrary = rows; },
      setSettings: settings => { siteSettings = settings; },
      articles: () => APP_DATA.articles
    };`, context);
  return context.api;
}

test('resolves library, Drive, relative, deterministic and default images', () => {
  const api = frontend();
  api.setLibrary([
    { id: 'one', label: 'First', category: 'Teacher Support', url: 'https://example.com/one.jpg' },
    { id: 'two', label: 'Second', category: 'teacher support', url: 'https://example.com/two.jpg' }
  ]);
  assert.equal(api.resolveImageUrl(' FIRST '), 'https://example.com/one.jpg');
  assert.equal(api.resolveImageUrl('https://drive.google.com/file/d/abc-123/view'), 'https://lh3.googleusercontent.com/d/abc-123');
  assert.equal(api.resolveImageUrl('https://drive.google.com/open?id=abc_123'), 'https://lh3.googleusercontent.com/d/abc_123');
  assert.equal(api.resolveImageUrl('./images/photo.jpg'), './images/photo.jpg');
  assert.equal(api.resolveImageUrl('', 'Teacher Support', 'A'), 'https://example.com/two.jpg');
  assert.equal(api.normalizeImageSource('not-a-url'), '');
  api.setSettings({ default_img_management: 'https://example.com/default.jpg' });
  assert.equal(api.resolveImageUrl('', 'Classroom Management'), 'https://example.com/default.jpg');
});

test('live sheet articles replace placeholders unless explicitly enabled', () => {
  const api = frontend();
  const row = { article_id: 'default', title: 'Title', category: 'Teacher Support',
    body: 'Article text', image_url: 'https://example.com/photo.jpg', is_live: true };
  const item = api.normalizePublishedArticle(row);
  assert.equal(item.image, row.image_url);
  assert.equal(item.subtitle, 'Article text…');
  assert.equal(item.takeaway, 'Article text');
  api.applyPublishedArticles([row]);
  assert.deepEqual(Array.from(api.articles(), article => article.id), ['default']);
  api.setSettings({ show_placeholder_articles: ' true ' });
  api.applyPublishedArticles([row]);
  assert.deepEqual(Array.from(api.articles(), article => article.id), ['default', 'other']);
});

test('reused hero images regain their error fallback', () => {
  const api = frontend();
  const image = {};
  api.setArticleImage(image, 'https://example.com/broken.jpg');
  image.onerror();
  assert.match(image.src, /^https:\/\/images\.unsplash\.com\//);
  api.setArticleImage(image, 'https://example.com/another.jpg');
  assert.equal(typeof image.onerror, 'function');
  image.onerror();
  assert.match(image.src, /^https:\/\/images\.unsplash\.com\//);
});

test('Apps Script exposes body and image aliases and blank live library rows', () => {
  const sheets = {
    published_articles: [
      ['article_id', 'title', 'category', 'article_body_text', 'body', 'cover_image_url', 'image_url', 'is_live'],
      ['a', 'Title', 'Teacher Support', '', 'Body', '', 'https://example.com/a.jpg', true]
    ],
    image_library: [
      ['image_id', 'category', 'label', 'image_url', 'is_live'],
      ['a', 'Teacher Support', 'A', 'https://example.com/a.jpg', ''],
      ['b', 'Teacher Support', 'B', 'https://example.com/b.jpg', false]
    ]
  };
  const context = vm.createContext({ SpreadsheetApp: { getActiveSpreadsheet: () => ({
    getSheetByName: name => sheets[name] && ({
      getLastRow: () => sheets[name].length,
      getDataRange: () => ({ getValues: () => sheets[name] })
    })
  }) } });
  vm.runInContext(fs.readFileSync('apps-script/Smore.gs', 'utf8') + `
    globalThis.articles = publishedArticles_();
    globalThis.images = publicTableItems_('image_library', IMAGE_LIBRARY_COLUMNS, ['image_url'], true);`, context);
  assert.equal(context.articles[0].article_body_text, 'Body');
  assert.equal(context.articles[0].cover_image_url, 'https://example.com/a.jpg');
  assert.deepEqual(Array.from(context.images, image => image.image_id), ['a']);
});
