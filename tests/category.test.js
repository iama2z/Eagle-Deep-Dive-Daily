const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function load(file, startMarker) {
  const src = fs.readFileSync(file, 'utf8');
  const start = src.indexOf(startMarker);
  const end = src.indexOf('\n    }\n', start) > 0 && file.endsWith('.html')
    ? src.indexOf('\n    }\n', start) + 7 : src.indexOf('\n}\n', start) + 3;
  const context = vm.createContext({});
  vm.runInContext(src.slice(start, end) + '\nglobalThis.fn = ' + startMarker.match(/function (\w+)/)[1] + ';', context);
  return context.fn;
}

for (const [file, marker] of [['index.html', 'function canonicalCategory('], ['apps-script/Smore.gs', 'function canonicalCategory_(']]) {
  test(`${file} maps typed topics to the site's categories`, () => {
    const canonical = load(file, marker);
    assert.equal(canonical('student engagement'), 'Student Engagement');
    assert.equal(canonical('  CLASSROOM  MANAGEMENT '), 'Classroom Management');
    assert.equal(canonical('Teacher Support & Wellbeing'), 'Teacher Support');
    assert.equal(canonical('Tech and AI'), 'Tech & AI');
    assert.equal(canonical('Tech & AI Tips'), 'Tech & AI');
    assert.equal(canonical('Something else'), 'Something else');
    assert.equal(canonical(''), '');
  });
}
