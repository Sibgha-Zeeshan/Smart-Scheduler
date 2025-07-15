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
      background: '#18181b',
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'flex-start',
      padding: '0',
      position: 'relative',
      overflow: 'hidden',
    }}>
      {/* Navbar */}
      <nav className="admin-navbar-glass" style={{
        width: 'calc(100% - 4rem)',
        margin: '0 2rem', // Add left and right margin
        background: 'rgba(30,41,59,0.82)',
        backdropFilter: 'blur(18px)',
        borderBottom: '1.5px solid #23272f',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        padding: '2.2rem 3rem', // 3rem left and right padding
        position: 'sticky',
        top: 0,
        zIndex: 10,
        boxShadow: '0 4px 32px 0 rgba(34,211,238,0.08)',
        minHeight: 68,
      }}>
        <div className="admin-navbar-logo" style={{ display: 'flex', alignItems: 'center', gap: '1.2rem', fontWeight: 800, fontSize: '1.85rem', color: '#22d3ee', letterSpacing: '-1px', cursor: 'pointer', transition: 'transform 0.2s' }}>
          <svg width="30" height="30" viewBox="0 0 24 24" fill="none" stroke="#22d3ee" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: 14, transition: 'transform 0.2s' }} className="admin-navbar-logoicon"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M16 2v4"/><path d="M8 2v4"/></svg>
          Smart Scheduler <span style={{ color: '#38bdf8', fontWeight: 700, marginLeft: 12 }}>Admin</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '3.2rem' }}>
          <div className="admin-navbar-links" style={{ display: 'flex', alignItems: 'center', gap: '2.8rem', fontWeight: 600, fontSize: '1.35rem', marginRight: '2.5rem' }}>
            <a href="#dashboard" className="admin-navbar-link" style={{ color: '#cbd5e1', textDecoration: 'none', transition: 'color 0.2s, border-bottom 0.2s', paddingBottom: 2, borderBottom: '2px solid transparent' }} onClick={e => { e.preventDefault(); setShowHistory(false); }}>Dashboard</a>
            <a href="#history" className="admin-navbar-link" style={{ color: '#cbd5e1', textDecoration: 'none', transition: 'color 0.2s, border-bottom 0.2s', paddingBottom: 2, borderBottom: '2px solid transparent' }} onClick={e => { e.preventDefault(); setShowHistory(true); }}>History</a>
          </div>
          <button className="admin-navbar-logout" onClick={onBack} style={{
            background: 'linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%)',
            color: '#fff',
            border: 'none',
            borderRadius: '9px',
            padding: '1.3rem 3.2rem',
            fontWeight: 700,
            fontSize: '1.35rem',
            cursor: 'pointer',
            boxShadow: '0 2px 12px rgba(34,211,238,0.13)',
            transition: 'background 0.2s, transform 0.2s, box-shadow 0.2s',
            letterSpacing: '0.5px',
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
            padding: '2.5rem 0 1.5rem 0',
            textAlign: 'center',
            background: '#23272f',
            color: '#fff',
            boxShadow: '0 2px 12px rgba(30,41,59,0.10)',
            marginBottom: '2.5rem',
            position: 'relative',
            zIndex: 1,
            animation: 'slideDown 0.8s ease-out',
          }}>
            {/* Back Icon */}
            {onBack && (
              <button
                onClick={onBack}
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
                }}
                className="admin-back-btn"
                aria-label="Back"
              >
                <svg width="26" height="26" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: 'block', transition: 'stroke 0.2s' }}>
                  <path d="M15 18l-6-6 6-6" />
                </svg>
              </button>
            )}
            <h1 style={{ fontSize: '2.3rem', fontWeight: 800, margin: 0, letterSpacing: '-1px', textShadow: '0 2px 12px #23272f55', animation: 'glow 2s ease-in-out infinite alternate' }}>Admin Dashboard</h1>
            <p style={{ fontSize: '1.1rem', color: '#cbd5e1', marginTop: '0.7rem', fontWeight: 500, letterSpacing: '0.5px', animation: 'fadeInUp 1s ease-out 0.3s both' }}>
              Welcome, <span style={{ color: '#fff', fontWeight: 700 }}>UMT Admin</span>! Manage user signups and approve or reject new accounts below.
            </p>
            {/* Stats Bar */}
            <div style={{
              display: 'flex',
              justifyContent: 'center',
              alignItems: 'center',
              gap: '3.5rem',
              margin: '2.5rem auto 0 auto',
              padding: '1.5rem 2.5rem',
              background: '#18181b',
              borderRadius: '16px',
              boxShadow: '0 2px 12px rgba(30,41,59,0.10)',
              maxWidth: 800,
              animation: 'fadeInUp 1s ease-out 0.5s both',
            }}>
              {stats.map((stat, i) => (
                <div key={stat.label} style={{
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'center',
                  minWidth: 120,
                  padding: '0.7rem 1.5rem',
                  borderRadius: '12px',
                  background: '#23272f',
                  boxShadow: '0 2px 8px rgba(30,41,59,0.07)',
                  border: `2px solid ${stat.color}`,
                  animation: `bounceIn 0.8s ease-out ${0.2 + i * 0.1}s both`,
                }}>
                  <span style={{ fontSize: '1.7rem', marginBottom: '0.2rem', color: stat.color }}>{stat.icon}</span>
                  <span style={{ fontWeight: 700, fontSize: '1.15rem', color: stat.color }}>{stat.value}</span>
                  <span style={{ color: '#cbd5e1', fontSize: '0.97rem', marginTop: '0.1rem', fontWeight: 500 }}>{stat.label}</span>
                </div>
              ))}
            </div>
          </header>
          <main style={{
            width: '100%',
            maxWidth: 1000,
            margin: '0 auto',
            padding: '2.5rem',
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
                <div style={{ color: '#38bdf8', fontSize: '1.1rem', textAlign: 'center', marginTop: '2.5rem', background: '#23272f', borderRadius: '10px', padding: '1.5rem 0', animation: 'fadeInUp 1s ease-out 0.5s both', boxShadow: '0 2px 8px rgba(30,41,59,0.07)' }}>
                  No pending signups. All caught up!
                </div>
              ) : (
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(auto-fit, minmax(340px, 1fr))',
                  gap: '2.5rem',
                  width: '100%',
                  marginTop: '1rem',
                }}>
                  {pendingSignups.map((user, idx) => (
                    <div key={user.email} style={{
                      background: '#23272f',
                      borderRadius: '14px',
                      boxShadow: '0 4px 16px rgba(30,41,59,0.13)',
                      padding: '2rem 1.5rem',
                      color: '#f1f5f9',
                      display: 'flex',
                      flexDirection: 'column',
                      alignItems: 'flex-start',
                      justifyContent: 'center',
                      minHeight: '140px',
                      border: '1.5px solid #23272f',
                      position: 'relative',
                      transition: 'transform 0.2s, box-shadow 0.2s',
                      animation: `slideInUp 0.7s cubic-bezier(.23,1.01,.32,1) both ${0.1 + idx * 0.1}s`,
                    }}>
                      <div style={{ fontSize: '1.13rem', fontWeight: 700, color: '#38bdf8', marginBottom: '0.3rem', animation: 'fadeInUp 1s ease-out 0.2s both' }}>{user.username}</div>
                      <div style={{ fontSize: '1.01rem', color: '#cbd5e1', marginBottom: '0.3rem', animation: 'fadeInUp 1s ease-out 0.3s both' }}>{user.email}</div>
                      <div style={{ fontSize: '0.97rem', color: '#fbbf24', fontWeight: 600, marginBottom: '1rem', textTransform: 'capitalize', animation: 'fadeInUp 1s ease-out 0.4s both' }}>{user.role}</div>
                      <div style={{ display: 'flex', gap: '0.7rem', marginTop: 'auto' }}>
                        <button onClick={() => onApprove(user)} style={{
                          background: 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)',
                          color: '#fff',
                          border: 'none',
                          borderRadius: '7px',
                          padding: '0.6rem 1.3rem',
                          fontWeight: 700,
                          fontSize: '1.01rem',
                          cursor: 'pointer',
                          boxShadow: '0 2px 8px rgba(56,189,248,0.13)',
                          transition: 'background 0.2s, transform 0.2s',
                          animation: 'bounceIn 0.8s ease-out 0.1s both',
                        }}>Approve</button>
                        <button onClick={() => onReject(user)} style={{
                          background: 'linear-gradient(90deg, #f43f5e 0%, #23272f 100%)',
                          color: '#fff',
                          border: 'none',
                          borderRadius: '7px',
                          padding: '0.6rem 1.3rem',
                          fontWeight: 700,
                          fontSize: '1.01rem',
                          cursor: 'pointer',
                          boxShadow: '0 2px 8px rgba(244,63,94,0.13)',
                          transition: 'background 0.2s, transform 0.2s',
                          animation: 'bounceIn 0.8s ease-out 0.2s both',
                        }}>Reject</button>
                      </div>
                    </div>
                  ))}
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
      `}</style>
    </div>
  );
}

export default AdminDashboard; 