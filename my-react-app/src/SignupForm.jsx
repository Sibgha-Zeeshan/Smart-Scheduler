import React, { useState } from "react";

function SignupForm({ onClose, onSignup }) {
  const [email, setEmail] = useState("");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [role, setRole] = useState("");
  const [emailError, setEmailError] = useState("");

  const handleSubmit = (e) => {
    e.preventDefault();
    if (!email.endsWith("@umt.edu.pk")) {
      setEmailError("Email must end with @umt.edu.pk");
      return;
    } else {
      setEmailError("");
    }
    if (!role || !username || !password) {
      setEmailError("All fields are required.");
      return;
    }
    if (onSignup) {
      onSignup({ email, username, password, role });
    }
    onClose();
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
      zoom: 0.67
    }}>
      {/* Decorative gradient overlay */}
      <div style={{
        position: 'absolute',
        top: 0,
        left: 0,
        right: 0,
        height: '4px',
        background: 'linear-gradient(90deg, #4ecdc4 0%, #22d3ee 50%, #ff6b6b 100%)',
        borderRadius: '24px 24px 0 0'
      }} />
      <button onClick={onClose} style={{ position: 'absolute', top: 18, right: 18, background: 'none', border: 'none', color: '#94a3b8', fontSize: '1.5rem', cursor: 'pointer', zIndex: 2 }}>&times;</button>
      <h2 style={{ 
        fontSize: '2.2rem', 
        marginBottom: '2.5rem', 
        color: '#ffffff', 
        textAlign: 'center', 
        fontWeight: '800',
        marginTop: '0',
        letterSpacing: '-1px',
        lineHeight: '1.1'
      }}>
        Create Your Account
      </h2>
      <form style={{ display: 'flex', flexDirection: 'column', gap: '2rem' }} onSubmit={handleSubmit}>
        <div>
          <label htmlFor="email" style={{ fontSize: '1.1rem', color: '#a8b2d1', fontWeight: '600', display: 'block', marginBottom: '1rem', textTransform: 'uppercase', letterSpacing: '1px' }}>Email Address</label>
          <input type="email" id="email" placeholder="Enter your Email" value={email} onChange={e => setEmail(e.target.value)} style={{ width: '100%', padding: '1.25rem 1.5rem', border: '2px solid rgba(255, 255, 255, 0.1)', borderRadius: '12px', fontSize: '1.15rem', color: '#ffffff', backgroundColor: 'rgba(255, 255, 255, 0.05)', boxSizing: 'border-box', transition: 'all 0.3s ease', backdropFilter: 'blur(10px)' }} />
          {emailError && <div style={{ color: '#ff6b6b', marginTop: '0.5rem', fontSize: '1rem' }}>{emailError}</div>}
        </div>
        <div>
          <label htmlFor="username" style={{ fontSize: '1.1rem', color: '#a8b2d1', fontWeight: '600', display: 'block', marginBottom: '1rem', textTransform: 'uppercase', letterSpacing: '1px' }}>Username</label>
          <input type="text" id="username" placeholder="Enter your username" value={username} onChange={e => setUsername(e.target.value)} style={{ width: '100%', padding: '1.25rem 1.5rem', border: '2px solid rgba(255, 255, 255, 0.1)', borderRadius: '12px', fontSize: '1.15rem', color: '#ffffff', backgroundColor: 'rgba(255, 255, 255, 0.05)', boxSizing: 'border-box', transition: 'all 0.3s ease', backdropFilter: 'blur(10px)' }} />
        </div>
        <div>
          <label htmlFor="password" style={{ fontSize: '1.1rem', color: '#a8b2d1', fontWeight: '600', display: 'block', marginBottom: '1rem', textTransform: 'uppercase', letterSpacing: '1px' }}>Password</label>
          <input type="password" id="password" placeholder="Create a password" value={password} onChange={e => setPassword(e.target.value)} style={{ width: '100%', padding: '1.25rem 1.5rem', border: '2px solid rgba(255, 255, 255, 0.1)', borderRadius: '12px', fontSize: '1.15rem', color: '#ffffff', backgroundColor: 'rgba(255, 255, 255, 0.05)', boxSizing: 'border-box', transition: 'all 0.3s ease', backdropFilter: 'blur(10px)' }} />
        </div>
        <div>
          <label htmlFor="role" style={{ fontSize: '1.1rem', color: '#a8b2d1', fontWeight: '600', display: 'block', marginBottom: '1rem', textTransform: 'uppercase', letterSpacing: '1px' }}>Role</label>
          <select id="role" value={role} onChange={e => setRole(e.target.value)} required style={{ width: '100%', padding: '1.25rem 1.5rem', border: '2px solid rgba(255, 255, 255, 0.1)', borderRadius: '12px', fontSize: '1.15rem', color: '#1e293b', backgroundColor: '#f1f5f9', boxSizing: 'border-box', transition: 'all 0.3s ease', backdropFilter: 'blur(10px)', appearance: 'none', WebkitAppearance: 'none', MozAppearance: 'none' }}>
            <option value="" style={{ color: '#64748b', backgroundColor: '#f1f5f9' }}>Select role</option>
            <option value="teacher" style={{ color: '#1e293b', backgroundColor: '#f1f5f9' }}>Teacher</option>
            <option value="student" style={{ color: '#1e293b', backgroundColor: '#f1f5f9' }}>Student</option>
          </select>
        </div>
        <button type="submit" style={{ width: '100%', padding: '1.375rem 1.5rem', background: 'linear-gradient(135deg, #4ecdc4 0%, #22d3ee 100%)', border: 'none', borderRadius: '12px', color: '#ffffff', fontSize: '1.25rem', fontWeight: '700', cursor: 'pointer', transition: 'all 0.3s ease', marginTop: '1.5rem', textTransform: 'uppercase', letterSpacing: '1px', boxShadow: '0 8px 25px rgba(34, 211, 238, 0.3)' }}>Sign Up</button>
      </form>
    </div>
  );
}

export default SignupForm; 