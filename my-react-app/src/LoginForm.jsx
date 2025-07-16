import React, { useState, useEffect } from "react";

function LoginForm({ onLogin, onClose, loading, error, defaultEmail = "", adminLogin = false, zoomed = false }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [username, setUsername] = useState("");

  useEffect(() => {
    if (defaultEmail) setEmail(defaultEmail);
  }, [defaultEmail]);

  const handleSubmit = (e) => {
    e.preventDefault();
    if (onLogin) {
      if (adminLogin) {
        onLogin({ username, password });
      } else {
        onLogin({ email, password });
      }
    }
  };

  return (
    <div style={{ 
      background: 'linear-gradient(145deg, #1a1a2e 0%, #16213e 100%)', 
      borderRadius: '24px', 
      padding: '3.5rem', 
      boxShadow: '0 25px 50px rgba(0, 0, 0, 0.3), inset 0 1px 0 rgba(255, 255, 255, 0.1)', 
      width: '100%',
      maxWidth: '420px',
      border: '1px solid rgba(255, 255, 255, 0.1)',
      position: 'relative',
      overflow: 'hidden',
      zIndex: 1001,
      ...(zoomed ? { zoom: 0.67 } : {})
    }}>
      {/* Decorative gradient overlay */}
      <div style={{
        position: 'absolute',
        top: 0,
        left: 0,
        right: 0,
        height: '4px',
        background: 'linear-gradient(90deg, #ff6b6b 0%, #4ecdc4 50%, #45b7d1 100%)',
        borderRadius: '24px 24px 0 0'
      }} />
      {onClose && (
        <button onClick={onClose} style={{ position: 'absolute', top: 18, right: 18, background: 'none', border: 'none', color: '#94a3b8', fontSize: '1.5rem', cursor: 'pointer', zIndex: 2 }}>&times;</button>
      )}
      <h2 style={{ 
        fontSize: '2.4rem', 
        marginBottom: '3rem', 
        color: '#ffffff', 
        textAlign: 'center', 
        fontWeight: '800',
        marginTop: '0',
        letterSpacing: '-1px',
        lineHeight: '1.1'
      }}>
        Welcome Back
      </h2>
      <form style={{ display: 'flex', flexDirection: 'column', gap: '2.5rem' }} onSubmit={handleSubmit}>
        {adminLogin ? (
          <>
            <div>
              <label htmlFor="username" style={{ fontSize: '1.1rem', color: '#a8b2d1', fontWeight: '600', display: 'block', marginBottom: '1rem', textTransform: 'uppercase', letterSpacing: '1px' }}>Admin Username</label>
              <input type="text" id="username" placeholder="Enter admin username" value={username} onChange={e => setUsername(e.target.value)} style={{ width: '100%', padding: '1.25rem 1.5rem', border: '2px solid rgba(255, 255, 255, 0.1)', borderRadius: '12px', fontSize: '1.15rem', color: '#ffffff', backgroundColor: 'rgba(255, 255, 255, 0.05)', boxSizing: 'border-box', transition: 'all 0.3s ease', backdropFilter: 'blur(10px)' }} />
            </div>
            <div>
              <label htmlFor="password" style={{ fontSize: '1.1rem', color: '#a8b2d1', fontWeight: '600', display: 'block', marginBottom: '1rem', textTransform: 'uppercase', letterSpacing: '1px' }}>Password</label>
              <input type="password" id="password" placeholder="Enter admin password" value={password} onChange={e => setPassword(e.target.value)} style={{ width: '100%', padding: '1.25rem 1.5rem', border: '2px solid rgba(255, 255, 255, 0.1)', borderRadius: '12px', fontSize: '1.15rem', color: '#ffffff', backgroundColor: 'rgba(255, 255, 255, 0.05)', boxSizing: 'border-box', transition: 'all 0.3s ease', backdropFilter: 'blur(10px)' }} />
            </div>
          </>
        ) : (
          <>
            <div>
              <label htmlFor="email" style={{ fontSize: '1.1rem', color: '#a8b2d1', fontWeight: '600', display: 'block', marginBottom: '1rem', textTransform: 'uppercase', letterSpacing: '1px' }}>Email Address</label>
              <input type="email" id="email" placeholder="Enter your Email" value={email} onChange={e => setEmail(e.target.value)} style={{ width: '100%', padding: '1.25rem 1.5rem', border: '2px solid rgba(255, 255, 255, 0.1)', borderRadius: '12px', fontSize: '1.15rem', color: '#ffffff', backgroundColor: 'rgba(255, 255, 255, 0.05)', boxSizing: 'border-box', transition: 'all 0.3s ease', backdropFilter: 'blur(10px)' }} />
            </div>
            <div>
              <label htmlFor="password" style={{ fontSize: '1.1rem', color: '#a8b2d1', fontWeight: '600', display: 'block', marginBottom: '1rem', textTransform: 'uppercase', letterSpacing: '1px' }}>Password</label>
              <input type="password" id="password" placeholder="Enter your password" value={password} onChange={e => setPassword(e.target.value)} style={{ width: '100%', padding: '1.25rem 1.5rem', border: '2px solid rgba(255, 255, 255, 0.1)', borderRadius: '12px', fontSize: '1.15rem', color: '#ffffff', backgroundColor: 'rgba(255, 255, 255, 0.05)', boxSizing: 'border-box', transition: 'all 0.3s ease', backdropFilter: 'blur(10px)' }} />
            </div>
          </>
        )}
        {error && <div style={{ color: '#ff6b6b', marginTop: '-1rem', fontSize: '1rem', textAlign: 'center' }}>{error}</div>}
        <button type="submit" style={{ width: '100%', padding: '1.375rem 1.5rem', background: 'linear-gradient(135deg, #ff6b6b 0%, #4ecdc4 100%)', border: 'none', borderRadius: '12px', color: '#ffffff', fontSize: '1.25rem', fontWeight: '700', cursor: 'pointer', transition: 'all 0.3s ease', marginTop: '1.5rem', textTransform: 'uppercase', letterSpacing: '1px', boxShadow: '0 8px 25px rgba(255, 107, 107, 0.3)' }} disabled={loading}>
          {loading ? 'Loading...' : 'Sign In'}
        </button>
      </form>
    </div>
  );
}

export default LoginForm; 