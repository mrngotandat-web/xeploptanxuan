// Chạy: node --test src/fitA4Page.test.mjs
import test from 'node:test';
import assert from 'node:assert/strict';
import { fitFontSize } from './fitA4Page.js';

test('giữ nguyên cỡ chữ khi không tràn', () => {
  assert.equal(fitFontSize(11.5, () => false), 11.5);
});

test('giảm dần tới cỡ đầu tiên vừa trang', () => {
  assert.equal(fitFontSize(11.5, (size) => size > 10.5), 10.5);
});

test('dừng ở cỡ tối thiểu nếu vẫn tràn', () => {
  assert.equal(fitFontSize(11.5, () => true), 8);
});
