const test = require('node:test');
const assert = require('node:assert/strict');
const { nextDuration, formatTime, minutesForKey } = require('../tools-timer.js');

test('digit keys map to minutes with 0 as 10', () => {
  assert.equal(minutesForKey('1'), 1);
  assert.equal(minutesForKey('5'), 5);
  assert.equal(minutesForKey('0'), 10);
  assert.equal(minutesForKey('a'), null);
});

test('same key within 1s doubles, otherwise resets', () => {
  const a = nextDuration(null, '1', 0);
  assert.equal(a.seconds, 60);
  const b = nextDuration(a, '1', 900);
  assert.equal(b.seconds, 120);
  assert.equal(nextDuration(b, '1', 1800).seconds, 240);
  assert.equal(nextDuration(b, '1', 3000).seconds, 60);
  assert.equal(nextDuration(b, '5', 1000).seconds, 300);
  assert.equal(nextDuration(a, 'x', 10), null);
});

test('formatTime', () => {
  assert.equal(formatTime(60), '1:00');
  assert.equal(formatTime(9.2), '0:10');
  assert.equal(formatTime(3725), '1:02:05');
});
