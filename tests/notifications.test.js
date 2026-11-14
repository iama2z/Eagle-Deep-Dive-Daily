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
