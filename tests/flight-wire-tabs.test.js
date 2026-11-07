const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const page = fs.readFileSync('index.html', 'utf8');
const start = page.indexOf('    function schoolUpdateAudience(');
const end = page.indexOf("\n    document.querySelectorAll('[data-flight-tab]').forEach(tab => {", start);

test('MMS sources are separate from district and other schools', () => {
  const context = vm.createContext({});
  vm.runInContext(page.slice(start, end), context);
  for (const source of ['MMS News', 'MMS', 'Maize Middle School', ' maize middle school ']) {
    assert.equal(context.schoolUpdateAudience({ source }), 'mms');
  }
  for (const source of ['District News', 'Maize USD 266', 'Maize South Middle School', '', undefined]) {
    assert.equal(context.schoolUpdateAudience({ source }), 'district');
  }
});

test('tab selection hides other panels and updates accessibility state', () => {
  const panels = {};
  const tabs = ['mms', 'district', 'smore'].map(name => {
    panels['flightPanel-' + name] = { hidden: name !== 'mms' };
    return { dataset: { flightTab: name }, attributes: {},
      setAttribute(key, value) { this.attributes[key] = value; },
      classList: { toggle() {} } };
  });
  const context = vm.createContext({
    document: { querySelectorAll: () => tabs, getElementById: id => panels[id] }
  });
  vm.runInContext(page.slice(start, end), context);
  for (const name of ['district', 'smore', 'mms']) {
    context.selectFlightTab(name);
    for (const tab of tabs) {
      const selected = tab.dataset.flightTab === name;
      assert.equal(panels['flightPanel-' + tab.dataset.flightTab].hidden, !selected);
      assert.equal(tab.attributes['aria-selected'], String(selected));
      assert.equal(tab.tabIndex, selected ? 0 : -1);
    }
  }
});

test('only dated posts older than seven days are collapsed by default', () => {
  const context = vm.createContext({});
  vm.runInContext(page.slice(start, end), context);
  const now = Date.parse('2026-11-07T21:00:00Z');
  assert.equal(context.schoolUpdateIsOlderThanWeek({ published_at: '2026-10-31T20:59:59Z' }, now), true);
  assert.equal(context.schoolUpdateIsOlderThanWeek({ published_at: '2026-10-31T21:00:00Z' }, now), false);
  assert.equal(context.schoolUpdateIsOlderThanWeek({ published_at: '2026-11-01T12:00:00Z' }, now), false);
  assert.equal(context.schoolUpdateIsOlderThanWeek({ published_at: '' }, now), false);
  assert.equal(context.schoolUpdateIsOlderThanWeek({ published_at: 'not a date' }, now), false);
  assert.match(page, /Show \$\{older\.length\} older updates/);
  assert.match(page, /expandedSchoolUpdateGroups\.add\(schoolUpdateGroup\(item\)\)/);
});

test('search opens the matching school update tab before scrolling', () => {
  assert.match(page, /selectFlightTab\(schoolUpdateAudience\(item\)\);\s*revealSchoolUpdate\(item, index\);\s*document\.getElementById\('school-update-'/);
  for (const match of page.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)) {
    assert.doesNotThrow(() => new vm.Script(match[1]));
  }
});
