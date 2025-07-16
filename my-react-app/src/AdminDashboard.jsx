import React, { useState } from "react";
import HistoryPanel from "./HistoryPanel";

function AdminDashboard({ users, pendingSignups, rejectedUsers, onApprove, onReject, onBack }) {
  const [showHistory, setShowHistory] = useState(false);
  const stats = [
    { label: 'Total Users', value: users.length, icon: '👥', color: '#38bdf8' },
    { label: 'Pending Requests', value: pendingSignups.length, icon: '⏳', color: '#fbbf24' },
    { label: 'Rejected', value: rejectedUsers.length, icon: '❌', color: '#f43f5e' },
  ];
  return (
    <div style={{
      minHeight: '100vh',
      width: '100vw',
      background: 'linear-gradient(135deg, #18181b 0%, #23272f 50%, #38bdf8 100%)',
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'flex-start',
      padding: '0',
      position: 'relative',
      overflow: 'hidden',
      fontFamily: 'Segoe UI, Roboto, Arial, sans-serif',
    }}>
      {/* Soft blurred colored blob background */}
      <div style={{
        position: 'absolute',
        top: '-80px',
        left: '50%',
        transform: 'translateX(-50%)',
        width: '520px',
        height: '320px',
        background: 'radial-gradient(circle, rgba(56,189,248,0.18) 0%, transparent 70%)',
        filter: 'blur(48px)',
        zIndex: 0,
        opacity: 0.7,
        pointerEvents: 'none',
      }} />
      {/* Navbar */}
      <nav className="admin-navbar-glass" style={{
        width: 'calc(100% - 2rem)',
        margin: '0 1rem',
        background: 'rgba(30,41,59,0.82)',
        backdropFilter: 'blur(18px)',
        borderBottom: '1.5px solid #23272f',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        padding: '0.7rem 1.2rem',
        position: 'sticky',
        top: 0,
        zIndex: 10,
        boxShadow: '0 4px 32px 0 rgba(34,211,238,0.08)',
        minHeight: 40,
      }}>
        <div className="admin-navbar-logo responsive-admin-logo" style={{ display: 'flex', alignItems: 'center', gap: '0.7rem', fontWeight: 800, fontSize: '0.93rem', color: '#22d3ee', letterSpacing: '-1px', cursor: 'pointer', transition: 'transform 0.2s' }}>
          <svg className="admin-navbar-logoicon" width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="#22d3ee" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: 14, transition: 'transform 0.2s' }}><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/></svg>
          Smart Scheduler <span style={{ color: '#38bdf8', fontWeight: 700, marginLeft: 12 }}>Admin</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1.2rem' }}>
          <div className="admin-navbar-links" style={{ display: 'flex', alignItems: 'center', gap: '1.1rem', fontWeight: 600, fontSize: '1rem', marginRight: '1.2rem' }}>
            <a href="#dashboard" className="admin-navbar-link" style={{ color: '#cbd5e1', textDecoration: 'none', transition: 'color 0.2s, border-bottom 0.2s', paddingBottom: 2, borderBottom: '2px solid transparent' }} onClick={e => { e.preventDefault(); setShowHistory(false); }}>Dashboard</a>
            <a href="#history" className="admin-navbar-link" style={{ color: '#cbd5e1', textDecoration: 'none', transition: 'color 0.2s, border-bottom 0.2s', paddingBottom: 2, borderBottom: '2px solid transparent' }} onClick={e => { e.preventDefault(); setShowHistory(true); }}>History</a>
          </div>
          <button className="admin-navbar-logout" onClick={onBack} style={{
            background: 'linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%)',
            color: '#fff',
            border: 'none',
            borderRadius: '7px',
            padding: '0.5rem 1.2rem',
            fontWeight: 700,
            fontSize: '1rem',
            cursor: 'pointer',
            boxShadow: '0 2px 8px rgba(34,211,238,0.13)',
            transition: 'background 0.2s, transform 0.2s, box-shadow 0.2s',
            letterSpacing: '0.3px',
          }}>Logout</button>
        </div>
        <style>{`
          .admin-navbar-glass {
            border-bottom: 2.5px solid #22d3ee22;
          }
          .admin-navbar-logo:hover {
            transform: scale(1.06) rotate(-2deg);
          }
          .admin-navbar-logoicon {
            transition: transform 0.2s;
          }
          .admin-navbar-logo:hover .admin-navbar-logoicon {
            transform: scale(1.18) rotate(-8deg);
            filter: drop-shadow(0 2px 8px #22d3ee99);
          }
          .admin-navbar-link:hover, .admin-navbar-link:focus {
            color: #22d3ee !important;
            border-bottom: 2px solid #22d3ee !important;
          }
          .admin-navbar-logout:hover {
            background: linear-gradient(90deg, #38bdf8 0%, #22d3ee 100%);
            transform: translateY(-2px) scale(1.04);
            box-shadow: 0 8px 24px 0 #22d3ee33;
          }
          @media (max-width: 900px) {
            .admin-navbar-glass {
              padding: 1.2rem 1.2rem !important;
            }
          }
        `}</style>
      </nav>
      {showHistory ? (
        <HistoryPanel
          accepted={users}
          pending={pendingSignups}
          rejected={rejectedUsers}
          onClose={() => setShowHistory(false)}
        />
      ) : (
        <>
          {/* Decorative Blobs */}
          <div className="bg-blob bg-blob-1" style={{
            position: 'absolute',
            top: '8%',
            left: '8%',
            width: '260px',
            height: '260px',
            background: 'radial-gradient(circle, rgba(167, 139, 250, 0.13) 0%, transparent 70%)',
            borderRadius: '50%',
            animation: 'float 7s ease-in-out infinite, pulse 4s ease-in-out infinite',
            zIndex: 0,
          }} />
          <div className="bg-blob bg-blob-2" style={{
            position: 'absolute',
            bottom: '10%',
            right: '10%',
            width: '320px',
            height: '320px',
            background: 'radial-gradient(circle, rgba(124, 58, 237, 0.13) 0%, transparent 70%)',
            borderRadius: '50%',
            animation: 'float 9s ease-in-out infinite reverse, pulse 4s ease-in-out infinite 2s',
            zIndex: 0,
          }} />
          {/* Topbar */}
          <header style={{
            width: '100%',
            padding: '1.2rem 0 1.1rem 0',
            textAlign: 'center',
            background: 'rgba(35,39,47,0.92)',
            color: '#fff',
            boxShadow: '0 2px 24px rgba(56,189,248,0.10)',
            marginBottom: '1.2rem',
            position: 'relative',
            zIndex: 1,
            borderRadius: '0 0 1.2rem 1.2rem',
            animation: 'fadeInUpAdmin 1.1s cubic-bezier(.39,.575,.56,1.000) 0.1s both',
          }}>
            {/* Back Icon */}
            {onBack && (
              <button
                onClick={onBack}
                className="admin-back-btn responsive-back-btn"
                style={{
                  position: 'absolute',
                  top: 24,
                  left: 32,
                  background: 'rgba(56,189,248,0.10)',
                  border: 'none',
                  borderRadius: '50%',
                  width: 44,
                  height: 44,
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  cursor: 'pointer',
                  boxShadow: '0 2px 8px rgba(56,189,248,0.10)',
                  transition: 'background 0.2s, transform 0.2s',
                  zIndex: 2,
                  outline: 'none',
                  animation: 'fadeInUp 1s ease-out 0.2s both',
                  padding: 0,
                }}
                aria-label="Back"
              >
                <svg className="responsive-back-icon" width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: 'block', transition: 'stroke 0.2s' }}>
                  <path d="M15 18l-6-6 6-6" />
                </svg>
              </button>
            )}
            <h1 style={{
              fontSize: '1.25rem',
              fontWeight: 800,
              margin: 0,
              letterSpacing: '-1px',
              textShadow: '0 2px 10px #38bdf855, 0 0 16px #38bdf899',
              animation: 'fadeInUpAdminTitle 1.2s cubic-bezier(.39,.575,.56,1.000) 0.2s both, glow 2s ease-in-out infinite alternate',
              background: 'linear-gradient(90deg, #38bdf8 0%, #22d3ee 100%)',
              WebkitBackgroundClip: 'text',
              WebkitTextFillColor: 'transparent',
              backgroundClip: 'text',
              textFillColor: 'transparent',
            }}>
              Admin Dashboard
            </h1>
            <p style={{ fontSize: '0.92rem', color: '#cbd5e1', marginTop: '0.4rem', fontWeight: 500, letterSpacing: '0.3px', animation: 'fadeInUp 1s ease-out 0.3s both' }}>
              Welcome, <span style={{ color: '#fff', fontWeight: 700 }}>UMT Admin</span>! Manage user signups and approve or reject new accounts below.
            </p>
            {/* Stats Bar */}
            <div style={{
              display: 'flex',
              justifyContent: 'center',
              alignItems: 'center',
              gap: '1.1rem',
              margin: '1.1rem auto 0 auto',
              padding: '1.1rem 3.5rem', // increased padding
              background: 'rgba(24,24,27,0.97)',
              borderRadius: '1rem',
              boxShadow: '0 2px 8px rgba(56,189,248,0.10)',
              maxWidth: 900, // increased maxWidth
              width: '100%', // ensure it stretches
              animation: 'fadeInUpAdminStats 1.1s cubic-bezier(.39,.575,.56,1.000) 0.5s both',
            }}>
              {stats.map((stat, i) => (
                <div key={stat.label} style={{
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'center',
                  minWidth: 70,
                  padding: '0.3rem 0.7rem',
                  borderRadius: '7px',
                  background: '#23272f',
                  boxShadow: '0 2px 4px rgba(30,41,59,0.07)',
                  border: `1.2px solid ${stat.color}`,
                  animation: `bounceIn 0.8s ease-out ${0.2 + i * 0.1}s both`,
                }}>
                  <span style={{ fontSize: '1.1rem', marginBottom: '0.1rem', color: stat.color }}>{stat.icon}</span>
                  <span style={{ fontWeight: 700, fontSize: '0.89rem', color: stat.color }}>{stat.value}</span>
                  <span style={{ color: '#cbd5e1', fontSize: '0.81rem', marginTop: '0.05rem', fontWeight: 500 }}>{stat.label}</span>
                </div>
              ))}
            </div>
          </header>
          <main style={{
            width: '100%',
            maxWidth: 1300, // increased from 1000
            margin: '0 auto',
            padding: '4rem 3rem', // increased padding
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            zIndex: 1,
          }}>
            <section style={{ width: '100%' }}>
              <h2 style={{ color: '#38bdf8', fontSize: '1.4rem', fontWeight: 700, marginBottom: '2.5rem', textAlign: 'center', letterSpacing: '-0.5px', animation: 'glow 2s ease-in-out infinite alternate 1s' }}>
                Pending Signup Requests
              </h2>
              {pendingSignups.length === 0 ? (
                <div style={{ color: '#38bdf8', fontSize: '0.92rem', textAlign: 'center', marginTop: '2.5rem', background: '#23272f', borderRadius: '10px', padding: '1.5rem 0', animation: 'fadeInUp 1s ease-out 0.5s both', boxShadow: '0 2px 8px rgba(30,41,59,0.07)' }}>
                  No pending signups. All caught up!
                </div>
              ) : (
                <div style={{ width: '100%', overflowX: 'auto', marginTop: '1rem', animation: 'fadeInUp 1s' }}>
                  <table style={{ width: '100%', minWidth: 500, borderCollapse: 'collapse', background: 'rgba(35,39,47,0.92)', borderRadius: '16px', color: '#f1f5f9', fontSize: '1.05rem', boxShadow: '0 2px 8px rgba(56,189,248,0.10)' }}>
                    <thead>
                      <tr style={{ background: 'rgba(56,189,248,0.10)' }}>
                        <th style={{ padding: '1rem 1.5rem', color: '#38bdf8', fontWeight: 800, textAlign: 'left', borderBottom: '2px solid #23272f', fontSize: '1.08rem', letterSpacing: '-0.5px' }}>Name</th>
                        <th style={{ padding: '1rem 1.5rem', color: '#38bdf8', fontWeight: 800, textAlign: 'left', borderBottom: '2px solid #23272f', fontSize: '1.08rem', letterSpacing: '-0.5px' }}>Email</th>
                        <th style={{ padding: '1rem 1.5rem', color: '#38bdf8', fontWeight: 800, textAlign: 'left', borderBottom: '2px solid #23272f', fontSize: '1.08rem', letterSpacing: '-0.5px' }}>Role</th>
                        <th style={{ padding: '1rem 1.5rem', color: '#38bdf8', fontWeight: 800, textAlign: 'center', borderBottom: '2px solid #23272f', fontSize: '1.08rem', letterSpacing: '-0.5px' }}>Actions</th>
                      </tr>
                    </thead>
                    <tbody>
                      {pendingSignups.map((user, idx) => (
                        <tr key={user.email} style={{ borderBottom: '1px solid #23272f', animation: `fadeInUp 0.7s cubic-bezier(.23,1.01,.32,1) both ${0.1 + idx * 0.1}s` }}>
                          <td style={{ padding: '1rem 1.5rem', fontWeight: 700, color: '#38bdf8' }}>{user.username}</td>
                          <td style={{ padding: '1rem 1.5rem', color: '#cbd5e1' }}>{user.email}</td>
                          <td style={{ padding: '1rem 1.5rem', color: '#fbbf24', fontWeight: 600, textTransform: 'capitalize' }}>{user.role}</td>
                          <td style={{ padding: '1rem 1.5rem', textAlign: 'center' }}>
                            <button onClick={() => onApprove(user)} style={{
                              background: 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)',
                              color: '#fff',
                              border: 'none',
                              borderRadius: '6px',
                              padding: '0.35rem 0.8rem',
                              fontWeight: 700,
                              fontSize: '0.98rem',
                              cursor: 'pointer',
                              boxShadow: '0 1px 4px rgba(56,189,248,0.10)',
                              marginRight: '0.5rem',
                              transition: 'background 0.2s, transform 0.2s',
                              animation: 'bounceIn 0.8s ease-out 0.1s both',
                            }}>Approve</button>
                            <button onClick={() => onReject(user)} style={{
                              background: 'linear-gradient(90deg, #f43f5e 0%, #23272f 100%)',
                              color: '#fff',
                              border: 'none',
                              borderRadius: '6px',
                              padding: '0.35rem 0.8rem',
                              fontWeight: 700,
                              fontSize: '0.98rem',
                              cursor: 'pointer',
                              boxShadow: '0 1px 4px rgba(244,63,94,0.10)',
                              transition: 'background 0.2s, transform 0.2s',
                              animation: 'bounceIn 0.8s ease-out 0.2s both',
                            }}>Reject</button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </section>
          </main>
        </>
      )}
      {/* Animations and theme styles */}
      <style>{`
        @keyframes slideDown {
          from { opacity: 0; transform: translateY(-20px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeInUp {
          from { opacity: 0; transform: translateY(30px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes slideInUp {
          from { opacity: 0; transform: translateY(50px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes bounceIn {
          0% { opacity: 0; transform: scale(0.3); }
          50% { opacity: 1; transform: scale(1.05); }
          70% { transform: scale(0.9); }
          100% { opacity: 1; transform: scale(1); }
        }
        @keyframes glow {
          0% { text-shadow: 0 0 5px #38bdf855; }
          100% { text-shadow: 0 0 20px #38bdf8cc, 0 0 30px #38bdf899; }
        }
        @keyframes float {
          0%, 100% { transform: translateY(0px); }
          50% { transform: translateY(-20px); }
        }
        @keyframes pulse {
          0%, 100% { transform: scale(1); opacity: 0.8; }
          50% { transform: scale(1.1); opacity: 1; }
        }
        .bg-blob {
          filter: blur(60px);
          opacity: 0.13;
          pointer-events: none;
        }
        .admin-back-btn:hover {
          background: #38bdf8;
          transform: scale(1.08) translateX(-2px);
        }
        .admin-back-btn:hover svg {
          stroke: #fff;
        }
        button:hover {
          transform: translateY(-2px) scale(1.05);
          box-shadow: 0 8px 25px rgba(56,189,248,0.13);
        }
        @keyframes fadeInUpAdmin {
          from { opacity: 0; transform: translateY(40px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeInUpAdminTitle {
          from { opacity: 0; transform: translateY(60px) scale(0.95); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        @keyframes fadeInUpAdminStats {
          from { opacity: 0; transform: translateY(30px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @media (max-width: 1100px) {
          .admin-dashboard-glass {
            max-width: 98vw !important;
            padding: 1.2rem 0.5rem !important;
          }
          main {
            padding: 2rem 0.5rem !important;
          }
        }
        @media (max-width: 900px) {
          .admin-dashboard-glass {
            max-width: 100vw !important;
            padding: 0.7rem 0.2rem !important;
            border-radius: 1.2rem !important;
          }
          nav, header, main, section {
            padding-left: 0.2rem !important;
            padding-right: 0.2rem !important;
          }
          .admin-navbar-logo {
            font-size: 1rem !important;
          }
          h1 {
            font-size: 1.1rem !important;
          }
          .stats-bar {
            flex-direction: column !important;
            gap: 0.7rem !important;
            padding: 0.7rem 0.2rem !important;
          }
          .responsive-admin-logo {
            font-size: 0.95rem !important;
            gap: 0.4rem !important;
          }
          .responsive-admin-logo .admin-navbar-logoicon {
            width: 22px !important;
            height: 22px !important;
            margin-right: 7px !important;
          }
          .responsive-back-btn {
            width: 32px !important;
            height: 32px !important;
            left: 10px !important;
            top: 12px !important;
          }
          .responsive-back-icon {
            width: 16px !important;
            height: 16px !important;
          }
        }
        @media (max-width: 700px) {
          .admin-dashboard-glass {
            padding: 0.5rem 0.05rem !important;
            border-radius: 0.7rem !important;
          }
          nav, header, main, section {
            padding-left: 0.05rem !important;
            padding-right: 0.05rem !important;
          }
          .admin-navbar-logo {
            font-size: 0.85rem !important;
          }
          h1 {
            font-size: 0.9rem !important;
          }
          .stats-bar {
            flex-direction: column !important;
            gap: 0.5rem !important;
            padding: 0.5rem 0.05rem !important;
          }
          .responsive-admin-logo {
            font-size: 0.8rem !important;
            gap: 0.2rem !important;
          }
          .responsive-admin-logo .admin-navbar-logoicon {
            width: 16px !important;
            height: 16px !important;
            margin-right: 4px !important;
          }
          .responsive-back-btn {
            width: 24px !important;
            height: 24px !important;
            left: 4px !important;
            top: 6px !important;
          }
          .responsive-back-icon {
            width: 12px !important;
            height: 12px !important;
          }
        }
        @media (max-width: 500px) {
          .admin-dashboard-glass {
            padding: 0.2rem 0.01rem !important;
            border-radius: 0.4rem !important;
          }
          h1 {
            font-size: 0.8rem !important;
          }
          .responsive-admin-logo {
            font-size: 0.65rem !important;
            gap: 0.1rem !important;
          }
          .responsive-admin-logo .admin-navbar-logoicon {
            width: 12px !important;
            height: 12px !important;
            margin-right: 2px !important;
          }
        }
      `}</style>
    </div>
  );
}

export default AdminDashboard; 