const test = require('node:test');
const assert = require('node:assert/strict');

const { makeTimestamp } = require('../src/utils/fs');

test('gera timestamps unicos mesmo no mesmo milissegundo', () => {
  const date = new Date(2026, 3, 9, 9, 15, 0, 123);
  const first = makeTimestamp(date);
  const second = makeTimestamp(date);
  const third = makeTimestamp(date);

  assert.equal(first, '20260409_091500_123');
  assert.equal(second, '20260409_091500_123_01');
  assert.equal(third, '20260409_091500_123_02');
});
