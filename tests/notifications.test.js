const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const page = fs.readFileSync('index.html', 'utf8');
const code = page.slice(page.indexOf('// NOTIFICATIONS START'), page.indexOf('// NOTIFICATIONS END'));
const ctx = vm.createContext({});
vm.runInContext(code, ctx);

const quote = (id, r) => ({ id, quote: `Quote ${id}`, reactions: { appreciate: 0, empathy: 0, respect: 0, ...r } });
const art = id => ({ id, title: `Title ${id}`, source: 'published_articles' });

test('reports new articles, new lounge posts and reaction increases only', () => {
  const prev = ctx.buildNotifySnapshot([art('a1')], [quote('q1', { respect: 2 }), quote('q2', { empathy: 5 })]);
  const cur = ctx.buildNotifySnapshot(
    [art('a1'), art('a2'), { id: 'static', title: 'x', source: 'static' }],
    [quote('q1', { respect: 4 }), quote('q2', { empathy: 3 }), quote('q3')]
  );
  const texts = Array.from(ctx.diffNotifySnapshots(prev, cur, 1), n => n.text);
  assert.deepEqual(texts, ['New article: Title a2', '👏 +2 on “Quote q1”', 'New in the Staff Lounge: “Quote q3”']);
});

test('admin submission alerts cover only unseen pending rows', () => {
  const subs = [
    { id: 's1', title: 'Old', status: 'pending' },
    { id: 's2', title: 'Fresh', status: 'pending' },
    { id: 's3', title: 'Done', status: 'approved' }
  ];
  const texts = Array.from(ctx.diffSubmissions({ s1: 1 }, subs, 1), n => n.text);
  assert.deepEqual(texts, ['New submission awaiting review: Fresh']);
});

test('notification controls live outside the hidden legacy header', () => {
  const header = page.match(/<header\b[\s\S]*?<\/header>/)[0];
  assert.ok(!header.includes('id="notifyBellBtn"'));
  const actions = page.slice(page.indexOf('<div id="siteActions">'), page.indexOf('<div id="miniBar">'));
  for (const id of ['notifyBellBtn', 'notifyBadge', 'notifyPanel', 'notifyList', 'notifyClearBtn', 'contributeFab']) {
    assert.ok(actions.includes(`id="${id}"`), `${id} belongs in the persistent actions`);
  }
});

function notificationBrowser(stored = new Map()) {
  const elements = {};
  for (const id of ['notifyBellBtn', 'notifyBadge', 'notifyPanel', 'notifyList', 'notifyClearBtn']) {
    const classes = new Set(['hidden']);
    elements[id] = {
      textContent: '', innerHTML: '', attributes: {}, listeners: {},
      classList: {
        contains: name => classes.has(name),
        toggle: (name, force) => force ? classes.add(name) : classes.delete(name)
      },
      setAttribute(name, value) { this.attributes[name] = value; },
      addEventListener(name, listener) { this.listeners[name] = listener; },
      contains: target => target === elements[id]
    };
  }
  const opened = [];
  const context = vm.createContext({
    APP_DATA: { articles: [art('a1')], loungeQuotes: [] },
    getStored: (key, fallback) => stored.has(key) ? JSON.parse(stored.get(key)) : fallback,
    setStored: (key, value) => stored.set(key, JSON.stringify(value)),
    escapeHtml: text => String(text),
    openArticleReader: id => opened.push(id),
    document: { getElementById: id => elements[id], addEventListener: () => {} }
  });
  vm.runInContext(code, context);
  context.initNotifications();
  return { context, elements, opened };
}

test('return visits show an unread badge, retain it on refresh, and open the article', () => {
  const stored = new Map();
  const first = notificationBrowser(stored);
  first.context.checkForUpdates();
  assert.equal(first.elements.notifyBadge.classList.contains('hidden'), true);

  const next = notificationBrowser(stored);
  next.context.APP_DATA.articles.push(art('a2'));
  next.context.checkForUpdates();
  assert.equal(next.elements.notifyBadge.textContent, '1');
  assert.equal(next.elements.notifyBadge.classList.contains('hidden'), false);
  assert.match(next.elements.notifyList.innerHTML, /New article: Title a2/);

  const refreshed = notificationBrowser(stored);
  refreshed.context.APP_DATA.articles.push(art('a2'));
  refreshed.context.checkForUpdates();
  assert.equal(refreshed.elements.notifyBadge.textContent, '1');
  refreshed.elements.notifyBellBtn.listeners.click({ stopPropagation() {} });
  assert.equal(refreshed.elements.notifyPanel.classList.contains('hidden'), false);
  assert.equal(refreshed.elements.notifyBellBtn.attributes['aria-expanded'], 'true');
  refreshed.elements.notifyList.listeners.click({
    target: { closest: () => ({ dataset: { notifyIndex: '0' } }) }
  });
  assert.deepEqual(refreshed.opened, ['a2']);
  assert.equal(refreshed.elements.notifyBadge.classList.contains('hidden'), true);
  assert.equal(refreshed.elements.notifyPanel.classList.contains('hidden'), true);
});

test('marking all read clears persisted notifications', () => {
  const stored = new Map();
  const { context, elements } = notificationBrowser(stored);
  context.checkForUpdates();
  context.APP_DATA.articles.push(art('a2'));
  context.checkForUpdates();
  elements.notifyClearBtn.listeners.click();
  assert.equal(elements.notifyBadge.classList.contains('hidden'), true);
  assert.match(elements.notifyList.innerHTML, /caught up/);
  const refreshed = notificationBrowser(stored);
  assert.equal(refreshed.elements.notifyBadge.classList.contains('hidden'), true);
});
