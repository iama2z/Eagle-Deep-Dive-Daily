const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

// ---------------------------------------------------------------
// Backend (apps-script/Smore.gs): addStudioComment_ and reactStudioComment_
// ---------------------------------------------------------------

function addCommentBackend() {
  const saved = [];
  let locked = false;
  const studioSheet = {};
  const commentsSheet = {};
  const context = vm.createContext({
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: name => (name === 'peer_studio' ? studioSheet : name === 'peer_studio_comments' ? commentsSheet : null),
        insertSheet: () => commentsSheet
      }),
      flush: () => {}
    },
    LockService: { getScriptLock: () => ({
      waitLock: () => { assert.equal(locked, false); locked = true; },
      releaseLock: () => { locked = false; }
    }) },
    Utilities: { getUuid: () => 'uuid-1' }
  });
  vm.runInContext(fs.readFileSync('apps-script/Smore.gs', 'utf8'), context);
  Object.assign(context, {
    smoreHeaders_: () => Array.from(PEER_STUDIO_COMMENT_COLUMNS_STUB),
    appendRecordByHeaders_: (sheet, headers, record) => saved.push(record),
    appendAuditLog_() {},
    liveRowsById_: (target, id) => (id === 'studio-1' ? [{ row: 2, issue: { studio_id: 'studio-1', title: 'Test Lesson', is_live: 'TRUE' } }] : [])
  });
  return { context, saved, isLocked: () => locked };
}
const PEER_STUDIO_COMMENT_COLUMNS_STUB = ['comment_id', 'studio_id', 'author', 'tag', 'text', 'agree_count', 'disagree_count', 'created_at', 'is_live'];

test('addStudioComment_ validates required fields and the submission must still exist', () => {
  const { context: api, saved } = addCommentBackend();
  assert.equal(api.addStudioComment_({ studio_id: '', author: 'A', tag: 'warm', text: 't' }).ok, false);
  assert.equal(api.addStudioComment_({ studio_id: 'studio-1', author: '', tag: 'warm', text: 't' }).ok, false);
  assert.equal(api.addStudioComment_({ studio_id: 'studio-1', author: 'A', tag: 'spicy', text: 't' }).ok, false);
  assert.equal(api.addStudioComment_({ studio_id: 'studio-1', author: 'A', tag: 'warm', text: '' }).ok, false);
  assert.equal(api.addStudioComment_({ studio_id: 'missing', author: 'A', tag: 'warm', text: 't' }).ok, false);
  assert.equal(saved.length, 0);
});

test('addStudioComment_ appends a new comment with zeroed vote counts', () => {
  const { context: api, saved, isLocked } = addCommentBackend();
  const result = api.addStudioComment_({ studio_id: 'studio-1', author: '  8th Science  ', tag: 'cool', text: 'Nice pacing!' });
  assert.equal(result.ok, true);
  assert.equal(result.comment.studio_id, 'studio-1');
  assert.equal(result.comment.author, '8th Science');
  assert.equal(result.comment.tag, 'cool');
  assert.equal(result.comment.text, 'Nice pacing!');
  assert.equal(result.comment.agree_count, 0);
  assert.equal(result.comment.disagree_count, 0);
  assert.equal(saved.length, 1);
  assert.equal(isLocked(), false);
});

function reactionBackend() {
  const rows = [
    ['comment_id', 'studio_id', 'author', 'tag', 'text', 'agree_count', 'disagree_count', 'created_at', 'is_live'],
    ['c1', 'studio-1', 'A', 'warm', 'Great hook', 0, 0, '', 'TRUE']
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
  const context = vm.createContext({
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: () => sheet }), flush: () => {} },
    LockService: { getScriptLock: () => ({
      waitLock: () => { assert.equal(locked, false); locked = true; },
      releaseLock: () => { locked = false; }
    }) }
  });
  vm.runInContext(fs.readFileSync('apps-script/Smore.gs', 'utf8'), context);
  return { context, rows, isLocked: () => locked };
}

test('reactStudioComment_ increments, undoes, and validates agree/disagree votes', () => {
  const { context: api, rows, isLocked } = reactionBackend();
  assert.equal(api.reactStudioComment_('c1', 'agree', false).count, 1);
  assert.equal(rows[1][5], 1);
  assert.equal(api.reactStudioComment_('c1', 'agree', false).count, 2);
  assert.equal(api.reactStudioComment_('c1', 'agree', true).count, 1);
  assert.equal(api.reactStudioComment_('c1', 'disagree', false).count, 1);
  assert.equal(rows[1][6], 1);
  assert.equal(api.reactStudioComment_('missing', 'agree', false).ok, false);
  assert.equal(api.reactStudioComment_('c1', 'unsure', false).ok, false);
  assert.equal(isLocked(), false);
});

// ---------------------------------------------------------------
// Frontend (index.html): tag tallies and the agree/disagree toggle
// ---------------------------------------------------------------

function frontend(fetch) {
  const page = fs.readFileSync('index.html', 'utf8');
  const code = page.slice(page.indexOf('    const studioListState'), page.indexOf('    function renderTechTools'));
  const stored = new Map();
  const status = { textContent: '' };
  const context = vm.createContext({
    URLSearchParams, fetch, console: { warn: () => {} }, escapeHtml: s => s, formatDateLabel: s => s,
    document: { getElementById: id => (id === 'studioForumStatus' ? status : null) },
    localStorage: { getItem: key => stored.get(key), setItem: (key, value) => stored.set(key, value) }
  });
  vm.runInContext(`
    const SUBMISSION_API_URL = 'https://example.com/exec';
    const APP_DATA = { studioComments: [{ id: 'c1', studioId: 's1', tag: 'warm', agree: 0, disagree: 0 }] };
    ${code}
    globalThis.comment = APP_DATA.studioComments[0];
    globalThis.APP_DATA = APP_DATA;
  `, context);
  context.renderStudioForum = () => {};
  return { api: context, status };
}

test('studioCommentCounts tallies warm, cool and mixed comments per submission', () => {
  const { api } = frontend();
  api.APP_DATA.studioComments = [
    { studioId: 's1', tag: 'warm' }, { studioId: 's1', tag: 'warm' },
    { studioId: 's1', tag: 'cool' }, { studioId: 's2', tag: 'mixed' }
  ];
  assert.deepEqual({ ...api.studioCommentCounts('s1') }, { warm: 2, cool: 1, mixed: 0 });
  assert.deepEqual({ ...api.studioCommentCounts('s2') }, { warm: 0, cool: 0, mixed: 1 });
  assert.deepEqual({ ...api.studioCommentCounts('missing') }, { warm: 0, cool: 0, mixed: 0 });
});

test('toggleStudioCommentReaction uses the saved server count and blocks overlapping votes', async () => {
  let finish;
  let calls = 0;
  const { api } = frontend(() => {
    calls++;
    return new Promise(resolve => { finish = resolve; });
  });
  const request = api.toggleStudioCommentReaction('c1', 'agree');
  assert.equal(api.comment.agree, 1);
  await api.toggleStudioCommentReaction('c1', 'agree');
  assert.equal(calls, 1);
  finish({ ok: true, json: async () => ({ ok: true, count: 5 }) });
  await request;
  assert.equal(api.comment.agree, 5);
  assert.equal(api.loadMyStudioCommentVotes().c1[0], 'agree');
});

test('toggleStudioCommentReaction rolls back the count and shows the backend error on failure', async () => {
  const { api, status } = frontend(async () => ({
    ok: true, json: async () => ({ ok: false, error: 'Comment not found.' })
  }));
  await api.toggleStudioCommentReaction('c1', 'disagree');
  assert.equal(api.comment.disagree, 0);
  assert.equal((api.loadMyStudioCommentVotes().c1 || []).length, 0);
  assert.equal(status.textContent, 'Comment not found.');
});
