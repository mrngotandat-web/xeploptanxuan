import React, { useState } from 'react';
import { supabase } from './supabaseClient';

export default function Login() {
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoading(true);
    const { error } = await supabase.auth.signInWithPassword({ email, password });
    if (error) alert("Lỗi đăng nhập: " + error.message);
    setLoading(false);
  };

  const handleSignUp = async (e) => {
    e.preventDefault();
    setLoading(true);
    const { error } = await supabase.auth.signUp({ email, password });
    if (error) alert("Lỗi đăng ký: " + error.message);
    else alert("Đăng ký thành công! Bạn có thể đăng nhập ngay.");
    setLoading(false);
  };

  return (
    <div style={{ display: 'flex', height: '100vh', justifyContent: 'center', alignItems: 'center', backgroundColor: '#f1f5f9' }}>
      <div style={{ background: '#fff', padding: '30px', borderRadius: '10px', boxShadow: '0 4px 6px rgba(0,0,0,0.1)', width: '350px' }}>
        <h2 style={{ textAlign: 'center', color: '#0369a1' }}>Đăng Nhập Hệ Thống</h2>
        <form>
          <input 
            type="email" placeholder="Email" value={email} onChange={e => setEmail(e.target.value)}
            style={{ width: '100%', padding: '10px', margin: '10px 0', borderRadius: '5px', border: '1px solid #ccc', boxSizing: 'border-box' }}
          />
          <input 
            type="password" placeholder="Mật khẩu" value={password} onChange={e => setPassword(e.target.value)}
            style={{ width: '100%', padding: '10px', marginBottom: '20px', borderRadius: '5px', border: '1px solid #ccc', boxSizing: 'border-box' }}
          />
          <button onClick={handleLogin} disabled={loading} style={{ width: '100%', padding: '10px', background: '#0284c7', color: '#fff', border: 'none', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer', marginBottom: '10px' }}>
            {loading ? 'Đang xử lý...' : 'ĐĂNG NHẬP'}
          </button>
          <button onClick={handleSignUp} disabled={loading} style={{ width: '100%', padding: '10px', background: '#fff', color: '#0284c7', border: '1px solid #0284c7', borderRadius: '5px', fontWeight: 'bold', cursor: 'pointer' }}>
            ĐĂNG KÝ TÀI KHOẢN MỚI
          </button>
        </form>
      </div>
    </div>
  );
}