const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const page = fs.readFileSync('index.html', 'utf8');

function contributionTypeHandler() {
  const start = page.indexOf("document.getElementById('tipType')?.addEventListener('change'");
  const end = page.indexOf("document.getElementById('shareTipForm')?.addEventListener", start);
  assert.notEqual(start, -1);
  assert.notEqual(end, -1);

  const elements = {};
  const classList = () => ({
    values: new Set(),
    add(name) { this.values.add(name); },
    toggle(name, enabled) {
      if (enabled) this.values.add(name);
      else this.values.delete(name);
    }
  });
  [
    'tipType', 'tipCoverImageField', 'tipCoverImage', 'tipLink', 'tipLinkLabel',
    'tipVideoHelp', 'shareTipModal', 'lessonTuneForm', 'tuneTitle'
  ].forEach(id => { elements[id] = { classList: classList() }; });

  const calls = { lounge: 0, views: [], scrolls: 0, focus: 0 };
  elements.tipType.addEventListener = (name, handler) => {
    assert.equal(name, 'change');
    calls.change = handler;
  };
  elements.lessonTuneForm.scrollIntoView = () => { calls.scrolls++; };
  elements.tuneTitle.focus = options => {
    assert.equal(options.preventScroll, true);
    calls.focus++;
  };
  const context = vm.createContext({
    document: { getElementById: id => elements[id] },
    tipModal: elements.shareTipModal,
    openLoungeModal: () => { calls.lounge++; },
    switchView: view => calls.views.push(view)
  });
  vm.runInContext(page.slice(start, end), context);
  return { elements, calls };
}

test('contribution type menu uses concise options', () => {
  const start = page.indexOf('<select id="tipType"');
  const end = page.indexOf('</select>', start);
  const options = [...page.slice(start, end).matchAll(/<option value="([^"]+)">([^<]+)<\/option>/g)]
    .map(([, value, label]) => [value, label]);
  assert.deepEqual(options, [
    ['article', 'Article'],
    ['lounge', 'Lounge'],
    ['video', 'Video'],
    ['feedback_submission', 'Feedback Submission']
  ]);
});

test('contribution types open the matching existing input form', () => {
  const { elements, calls } = contributionTypeHandler();
  const select = type => calls.change({ target: { value: type } });

  select('article');
  assert.equal(elements.tipCoverImageField.classList.values.has('hidden'), false);
  assert.equal(elements.tipLink.required, false);

  select('video');
  assert.equal(elements.tipCoverImageField.classList.values.has('hidden'), true);
  assert.equal(elements.tipLink.required, true);

  select('lounge');
  assert.equal(elements.shareTipModal.classList.values.has('hidden'), true);
  assert.equal(calls.lounge, 1);

  select('feedback_submission');
  assert.deepEqual(calls.views, ['view-studio']);
  assert.equal(calls.scrolls, 1);
  assert.equal(calls.focus, 1);
});
