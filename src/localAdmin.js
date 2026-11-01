// Tài khoản admin tạm, dùng khi chưa có Supabase. Repo public nên chỉ lưu hash, không lưu mật khẩu.
// ponytail: chỉ kiểm tra ở trình duyệt, ai sửa được JS đều vượt qua; bỏ file này khi Supabase Auth chạy lại.
export const ADMIN_USERNAME = 'admin';
export const ADMIN_PASSWORD_HASH = '8cc491c07fb7c9d959bb11522fec933ebccb7e8cd9105cfff732f1891c8580dc';
export const LOCAL_ADMIN_SESSION = { user: { email: ADMIN_USERNAME } };

export async function sha256Hex(text) {
  const digest = await crypto.subtle.digest('SHA-256', new TextEncoder().encode(text));
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

export async function verifyLocalAdmin(username, password, expectedHash = ADMIN_PASSWORD_HASH) {
  return username === ADMIN_USERNAME && (await sha256Hex(password)) === expectedHash;
}
