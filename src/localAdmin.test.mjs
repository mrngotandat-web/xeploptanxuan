// Chạy: node --test src/localAdmin.test.mjs
import test from 'node:test';
import assert from 'node:assert/strict';
import { ADMIN_USERNAME, ADMIN_PASSWORD_HASH, sha256Hex, verifyLocalAdmin } from './localAdmin.js';

// SHA-256 của "abc" (vector chuẩn FIPS 180), không phải mật khẩu thật
const ABC_HASH = 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad';

test('sha256Hex khớp vector chuẩn', async () => {
  assert.equal(await sha256Hex('abc'), ABC_HASH);
});

test('chấp nhận đúng username và mật khẩu', async () => {
  assert.equal(await verifyLocalAdmin(ADMIN_USERNAME, 'abc', ABC_HASH), true);
});

test('từ chối sai mật khẩu', async () => {
  assert.equal(await verifyLocalAdmin(ADMIN_USERNAME, 'sai', ABC_HASH), false);
});

test('từ chối sai username', async () => {
  assert.equal(await verifyLocalAdmin('khac', 'abc', ABC_HASH), false);
});

test('hash thật là SHA-256 hợp lệ, không phải mật khẩu dạng chữ', () => {
  assert.match(ADMIN_PASSWORD_HASH, /^[0-9a-f]{64}$/);
});
