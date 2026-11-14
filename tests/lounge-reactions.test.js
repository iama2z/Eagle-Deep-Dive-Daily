const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function backend() {
  const rows = [
    ['respect', 'quote', 'quote_id', 'is_live', 'appreciate', 'empathy', 'attribution', 'post_type'],
    [0, 'Manually entered', '', true, 0, 0, 'Staff', 'win'],
    [2, 'Existing post', 'existing', true, 4, 1, 'Staff', 'quote'],
    ['', '', '', '', '', '', '', '']
  ];
  const sheet = {
    getLastRow: () => rows.length,
    getLastColumn: () => rows[0].length,
    getDataRange: () => ({ getValues: () => rows.map(row => row.slice()) }),
    getRange: (r, c, height = 1, width = 1) => ({
      getValues: () => rows.slice(r - 1, r - 1 + height).map(row => row.slice(c - 1, c - 1 + width)),
      getValue: () => rows[r - 1][c - 1],
      setValue: value => { rows[r - 1][c - 1] = value; }
    })
  };
  let locked = false;
  let ids = 0;
  const context = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }), flush: () => {} },
    LockService: { getScriptLock: () => ({
      waitLock: () => { assert.equal(locked, false); locked = true; },
      releaseLock: () => { locked = false; }
    }) },
    Utilities: { getUuid: () => `uuid-${++ids}` }
  });
  vm.runInContext(fs.readFileSync('apps-script/Smore.gs', 'utf8'), context);
  return { context, rows, isLocked: () => locked };
}

test('manual posts gain stable IDs and reactions persist across feed reads', () => {
  const { context: api, rows, isLocked } = backend();
  api.ensureLoungeQuoteIds_();
  assert.equal(rows[1][2], 'q_uuid-1');
  assert.equal(rows[2][2], 'existing');
  assert.equal(rows[2][4], 4);
  assert.equal(rows[3][2], '');
  api.ensureLoungeQuoteIds_();
  assert.equal(rows[1][2], 'q_uuid-1');
  assert.equal(api.addLoungeReaction_('q_uuid-1', 'appreciate', false).count, 1);
  const feed = () => api.publicTableItems_('lounge_quotes', ['quote_id', 'quote', 'is_live', 'appreciate'], ['quote']);
  assert.equal(feed()[0].appreciate, '1');
  assert.equal(api.addLoungeReaction_('q_uuid-1', 'appreciate', false).count, 2);
  assert.equal(feed()[0].appreciate, '2');
  assert.equal(api.addLoungeReaction_('q_uuid-1', 'appreciate', true).count, 1);
  assert.equal(api.addLoungeReaction_('missing', 'appreciate', false).ok, false);
  assert.equal(api.addLoungeReaction_('existing', 'invalid', false).ok, false);
  assert.equal(isLocked(), false);
});

test('staff lounge post types are available, labeled, and preserved in feed data', () => {
  const { context: api } = backend();
  const types = vm.runInContext('LOUNGE_POST_TYPES', api);
  const page = fs.readFileSync('index.html', 'utf8');
  const selectStart = page.indexOf('<select id="loungePostType"');
  const selectEnd = page.indexOf('</select>', selectStart);
  const options = [...page.slice(selectStart, selectEnd).matchAll(/<option value="([^"]+)">([^<]+)<\/option>/g)]
    .map(([, value, label]) => [value, label]);
  const labelsStart = page.indexOf('    const LOUNGE_TYPE_LABELS');
  const labelsEnd = page.indexOf('    let loungeFilter', labelsStart);
  const mapperStart = page.indexOf('    function mapLoungeQuotes');
  const mapperEnd = page.indexOf('    async function refreshLounge', mapperStart);
  const frontend = vm.createContext({});
  vm.runInContext(`
    ${page.slice(labelsStart, labelsEnd)}
    ${page.slice(mapperStart, mapperEnd)}
    globalThis.labels = LOUNGE_TYPE_LABELS;
    globalThis.mapQuotes = mapLoungeQuotes;
  `, frontend);

  assert.deepEqual(Array.from(types), ['win', 'quote', 'question', 'random_delight', 'venting', 'shout_out']);
  assert.deepEqual(options, [
    ['win', 'Win'],
    ['quote', 'Quote'],
    ['question', 'Question'],
    ['random_delight', 'Random Delight'],
    ['venting', 'Venting'],
    ['shout_out', 'Shout Out']
  ]);
  assert.deepEqual(
    Array.from(types, type => [type, frontend.labels[type]]),
    options
  );
  assert.deepEqual(
    Array.from(frontend.mapQuotes(types.map(post_type => ({ post_type, quote: post_type }))), post => post.type),
    Array.from(types)
  );
  assert.equal(frontend.labels.shout_out, 'Shout Out');
  assert.equal(frontend.mapQuotes([{ post_type: 'legacy', quote: 'Old post' }])[0].type, 'quote');
});

function frontend(fetch) {
  const page = fs.readFileSync('index.html', 'utf8');
  const code = page.slice(page.indexOf('    const LOUNGE_REACTIONS'), page.indexOf('    function renderStudioItems'));
  const stored = new Map();
  const status = { textContent: '' };
  const context = vm.createContext({
    URLSearchParams, fetch, console: { warn: () => {} },
    document: { getElementById: id => id === 'loungeReactionStatus' ? status : null },
    localStorage: { getItem: key => stored.get(key), setItem: (key, value) => stored.set(key, value) }
  });
  vm.runInContext(`
    const SUBMISSION_API_URL = 'https://example.com/exec';
    const APP_DATA = { loungeQuotes: [{ id: 'q', reactions: { appreciate: 0, empathy: 0, respect: 0 } }] };
    ${code}
    globalThis.post = APP_DATA.loungeQuotes[0];
  `, context);
  return { api: context, status };
}

test('client uses the saved server count and blocks overlapping post reactions', async () => {
  let finish;
  let calls = 0;
  const { api } = frontend(() => {
    calls++;
    return new Promise(resolve => { finish = resolve; });
  });
  const request = api.toggleLoungeReaction('q', 'appreciate');
  assert.equal(api.post.reactions.appreciate, 1);
  await api.toggleLoungeReaction('q', 'appreciate');
  assert.equal(calls, 1);
  finish({ ok: true, json: async () => ({ ok: true, count: 8 }) });
  await request;
  assert.equal(api.post.reactions.appreciate, 8);
  assert.equal(api.loadMyReactions().q[0], 'appreciate');
});

test('failed saves restore the exact count and display the backend error', async () => {
  const { api, status } = frontend(async () => ({
    ok: true, json: async () => ({ ok: false, error: 'Post not found.' })
  }));
  await api.toggleLoungeReaction('q', 'appreciate');
  assert.equal(api.post.reactions.appreciate, 0);
  assert.equal(api.loadMyReactions().q.length, 0);
  assert.equal(status.textContent, 'Post not found.');
  await api.toggleLoungeReaction('q', 'appreciate');
  assert.equal(api.post.reactions.appreciate, 0);
});
