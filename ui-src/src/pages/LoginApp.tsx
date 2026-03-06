import React, { useState } from 'react';
import { Lock, User } from 'lucide-react';
import { AdminApi } from '../api/main';

export function LoginApp() {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  const submit = async (event: React.FormEvent) => {
    event.preventDefault();
    if (!username.trim() || !password) {
      setError('请输入用户名和密码');
      return;
    }

    setLoading(true);
    setError('');
    try {
      await AdminApi.login(username.trim(), password);
      window.location.href = '/ui/admin.html';
    } catch (err) {
      setError(err instanceof Error ? err.message : '登录失败');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="login-page">
      <div className="login-card">
        <div className="login-logo">M</div>
        <h1>MSS AI PPT</h1>
        <p>管理后台登录</p>

        <form onSubmit={submit} className="login-form">
          <label>
            <span><User size={14} /> 用户名</span>
            <input value={username} onChange={(e) => setUsername(e.target.value)} autoComplete="username" />
          </label>
          <label>
            <span><Lock size={14} /> 密码</span>
            <input type="password" value={password} onChange={(e) => setPassword(e.target.value)} autoComplete="current-password" />
          </label>
          {error && <div className="login-error">{error}</div>}
          <button type="submit" className="btn-primary" disabled={loading}>
            {loading ? '登录中...' : '登录'}
          </button>
        </form>
      </div>
    </div>
  );
}
