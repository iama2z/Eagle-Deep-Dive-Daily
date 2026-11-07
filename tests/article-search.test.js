const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const page = fs.readFileSync('index.html', 'utf8');
const start = page.indexOf('    function searchByKeywords(');
const end = page.indexOf('    // --- EVENT LISTENERS ---', start);

function element() {
  return {
    children: [], hidden: true, attributes: {}, listeners: {},
    classList: { add() {}, remove() {} },
    setAttribute(key, value) { this.attributes[key] = value; },
    appendChild(child) { this.children.push(child); },
    replaceChildren() { this.children = []; },
    addEventListener(name, callback) { this.listeners[name] = callback; }
  };
}

function context(index = null) {
  const hidden = [];
  const ctx = vm.createContext({
    articleFuse: index,
    document: {
      createElement: element,
      getElementById: id => ({ classList: { add: name => hidden.push([id, name]) } })
    },
    escapeHtml: value => String(value).replace(/[&<>"']/g, char => ({
      '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    })[char])
  });
  vm.runInContext(page.slice(start, end), ctx);
  return { ctx, hidden };
}

test('keywords are case-insensitive, deduplicated, and matched across any indexed field', () => {
  let query;
  const { ctx } = context();
  ctx.searchByKeywords('Video, students VIDEO!', {
    search(value) { query = JSON.parse(JSON.stringify(value)); return []; }
  });
  assert.deepEqual(query, {
    $and: ['video', 'students'].map(token => ({
      $or: ['title', 'category', 'desc', 'tags', 'body'].map(key => ({ [key]: token }))
    }))
  });
  assert.equal(ctx.searchByKeywords('... !', { search() { throw new Error('Unexpected search'); } }).length, 0);
});

test('empty queries and unavailable indexes give explicit feedback without generated answers', () => {
  const { ctx } = context();
  const output = element();
  ctx.searchArticles('bellringer', output);
  assert.match(output.children[0].textContent, /not ready/);
  ctx.articleFuse = { search() { throw new Error('Empty input should not search'); } };
  ctx.searchArticles('   ', output);
  assert.match(output.children[0].textContent, /Enter a topic/);
});

test('no matches are safe text and suggest refining the search', () => {
  const { ctx } = context({ search: () => [] });
  const output = element();
  ctx.searchArticles('<img onerror=alert(1)>', output);
  assert.equal(output.children.length, 1);
  assert.match(output.children[0].textContent, /No matches.*Try fewer keywords/);
  assert.equal(output.children[0].innerHTML, undefined);
  assert.equal(output.children[0].attributes.role, 'status');
});

test('results show escaped article excerpts, disclose the limit, and open the selected article', () => {
  let opened = false;
  const results = Array.from({ length: 8 }, (_, i) => ({
    item: {
      type: 'Article Playbook', title: `<Article ${i}>`, category: 'Teacher Support',
      desc: '<script>not markup</script>', action: () => { opened = true; }
    }
  }));
  const { ctx, hidden } = context({ search: () => results });
  const output = element();
  assert.equal(ctx.searchArticles('article', output).length, 8);
  assert.equal(output.children.length, 7);
  assert.match(output.children[0].textContent, /8 matches - showing the top 6/);
  const button = output.children[1];
  assert.equal(button.type, 'button');
  assert.match(button.innerHTML, /&lt;Article 0&gt;/);
  assert.match(button.innerHTML, /&lt;script&gt;/);
  button.listeners.click();
  assert.equal(opened, true);
  assert.deepEqual(hidden, [['searchResultsDropdown', 'hidden'], ['geminiModal', 'hidden']]);
});

test('article index excludes non-articles and uses the same weighted search options', () => {
  const indexes = [];
  const ctx = vm.createContext({
    APP_DATA: {
      articles: [{ id: 'article', title: 'Catch-up', category: 'Support', takeaway: 'Missing',
        subtitle: 'Students', body: ['Full text'], tags: ['catch-up'] }],
      quickTipsBySection: {},
      resources: [{ id: 'resource', title: 'Directory', tags: [] }],
      bulletin: [], techTools: []
    },
    smoreIssues: [], schoolUpdates: [],
    Fuse: function(items, options) { indexes.push({ items, options }); },
    fuse: null, articleFuse: null
  });
  const initStart = page.indexOf('    function initFuse()');
  const initEnd = page.indexOf('    // --- DYNAMIC CONTEXT-AWARE', initStart);
  vm.runInContext(page.slice(initStart, initEnd), ctx);
  ctx.initFuse();
  assert.equal(indexes[0].items.length, 2);
  assert.equal(indexes[1].items.length, 1);
  assert.equal(indexes[1].items[0].id, 'article');
  assert.equal(indexes[1].items[0].body, 'Full text');
  assert.equal(indexes[0].options, indexes[1].options);
  assert.equal(indexes[1].options.ignoreLocation, true);
  assert.equal(indexes[1].options.keys[0].weight, 3);
});

test('AI search controls, network calls, and legacy Sheet messaging are removed', () => {
  assert.doesNotMatch(page, /Powered by Gemini|Ask AI:|AI Proxy|workerUrlInput|searchModeToggleBtn|function askCoach/);
  assert.doesNotMatch(page.slice(start, end), /fetch\(|setTimeout\(/);
  assert.doesNotMatch(page, /setText\('coach\w+Text', 'coach_/);
  assert.match(page, /setText\('coachHeadingText', 'article_search_heading'\)/);
  assert.match(page, /data-prompt="missing students"/);
  assert.match(page, /data-prompt="student video"/);
  assert.match(page, /globalSearchSubmitBtn'\)\?\.addEventListener\('click', runGlobalSearch\)/);
});

test('search buttons, keyboard submission, shortcuts, and clearing the top query are wired', () => {
  const nodes = {};
  const chips = [{ ...element(), dataset: { prompt: 'student video' } }];
  const calls = [];
  const getNode = id => {
    if (!nodes[id]) {
      const node = element();
      node.value = '';
      node.classList = {
        add: () => { node.hidden = true; },
        remove: () => { node.hidden = false; }
      };
      nodes[id] = node;
    }
    return nodes[id];
  };
  const ctx = vm.createContext({
    document: {
      getElementById: getNode,
      querySelectorAll: () => chips,
      querySelector: () => ({ focus: () => calls.push('focus-result') })
    },
    fuse: {},
    renderSearchResults: (query, target) => { calls.push([query, target]); return []; },
    searchArticles: (query, target) => calls.push([query, target])
  });
  const listenerStart = page.indexOf('      // Search Input (');
  const listenerEnd = page.indexOf('      // Modals: Share Tip', listenerStart);
  vm.runInContext(page.slice(listenerStart, listenerEnd), ctx);
  const global = getNode('globalSearchInput');
  global.value = 'OrbitNote';
  global.listeners.input();
  assert.equal(getNode('searchResultsDropdown').hidden, false);
  assert.deepEqual(calls.pop(), ['OrbitNote', getNode('searchResultsList')]);
  getNode('globalSearchSubmitBtn').listeners.click();
  assert.deepEqual(calls.pop(), ['OrbitNote', getNode('searchResultsList')]);
  let prevented = 0;
  global.listeners.keydown({ key: 'Enter', preventDefault: () => prevented++ });
  assert.equal(calls.pop(), 'focus-result');
  assert.equal(prevented, 1);
  global.listeners.keydown({ key: 'Escape' });
  assert.equal(getNode('searchResultsDropdown').hidden, true);
  global.value = '';
  global.listeners.input();
  assert.equal(getNode('searchResultsDropdown').hidden, true);

  const sidebar = getNode('sidebarCoachInput');
  sidebar.value = 'bellringer';
  getNode('sidebarCoachSubmitBtn').listeners.click();
  assert.deepEqual(calls.pop(), ['bellringer', getNode('sidebarCoachOutput')]);
  sidebar.listeners.keydown({ key: 'Enter', shiftKey: true });
  assert.equal(prevented, 1);
  sidebar.listeners.keydown({ key: 'Enter', target: sidebar, preventDefault: () => prevented++ });
  assert.equal(prevented, 2);
  assert.deepEqual(calls.pop(), ['bellringer', getNode('sidebarCoachOutput')]);
  chips[0].listeners.click();
  assert.equal(sidebar.value, 'student video');
  assert.deepEqual(calls.pop(), ['student video', getNode('sidebarCoachOutput')]);
});
