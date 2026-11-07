const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const page = fs.readFileSync('index.html', 'utf8');
const start = page.indexOf('    const stopLoadingMessages =');
const end = page.indexOf('\n  </script>', start);
const code = page.slice(start, end);

function runLoadingMessages(random) {
  const message = { textContent: '' };
  let tick;
  let cleared = false;
  let delay;
  const math = Object.create(Math);
  math.random = random;
  const context = vm.createContext({
    Math: math,
    document: { getElementById: () => message },
    setInterval: (callback, ms) => { tick = callback; delay = ms; return 42; },
    clearInterval: timer => { assert.equal(timer, 42); cleared = true; }
  });
  vm.runInContext(`${code}\nglobalThis.stop = stopLoadingMessages;`, context);
  return { message, tick: () => tick(), stop: () => context.stop(), delay, isCleared: () => cleared };
}

test('messages start randomly and cycle without repetition', () => {
  const first = runLoadingMessages(() => 0);
  const second = runLoadingMessages(() => 0.99);
  assert.notEqual(first.message.textContent, second.message.textContent);
  assert.equal(first.delay, 2500);
  const messages = [first.message.textContent];
  for (let i = 0; i < 5; i++) {
    first.tick();
    messages.push(first.message.textContent);
  }
  assert.equal(new Set(messages).size, 6);
  first.tick();
  assert.equal(first.message.textContent, messages[0]);
});

test('the loading timer is stopped before status is replaced', () => {
  const loading = runLoadingMessages(() => 0);
  loading.stop();
  assert.equal(loading.isCleared(), true);
  assert.match(page, /stopLoadingMessages\(\);\s*status\.textContent = notice;/);
});
