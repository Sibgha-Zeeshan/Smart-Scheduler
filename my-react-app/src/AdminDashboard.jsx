import React, { useState } from "react";
import HistoryPanel from "./HistoryPanel";
import TimetableUpload from "./TimetableUpload";

function AdminDashboard({ users, pendingSignups, rejectedUsers, onApprove, onReject, onBack, adminUser }) {
  const [showHistory, setShowHistory] = useState(false);
  const [showTimetable, setShowTimetable] = useState(false);
  const [navOpen, setNavOpen] = useState(false);
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
            <a href="#dashboard" className="admin-navbar-link" style={{ color: '#cbd5e1', textDecoration: 'none', transition: 'color 0.2s, border-bottom 0.2s', paddingBottom: 2, borderBottom: '2px solid transparent' }} onClick={e => { e.preventDefault(); setShowHistory(false); setShowTimetable(false); }}>Dashboard</a>
            <a href="#history" className="admin-navbar-link" style={{ color: '#cbd5e1', textDecoration: 'none', transition: 'color 0.2s, border-bottom 0.2s', paddingBottom: 2, borderBottom: '2px solid transparent' }} onClick={e => { e.preventDefault(); setShowHistory(true); setShowTimetable(false); }}>History</a>
            <a href="#timetable" className="admin-navbar-link" style={{ color: '#cbd5e1', textDecoration: 'none', transition: 'color 0.2s, border-bottom 0.2s', paddingBottom: 2, borderBottom: '2px solid transparent' }} onClick={e => { e.preventDefault(); setShowHistory(false); setShowTimetable(true); }}>Timetable</a>
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
          {/* Hamburger Menu for Mobile */}
          <button
            className="admin-hamburger"
            onClick={() => setNavOpen(!navOpen)}
            style={{
              display: 'none',
              background: 'none',
              border: 'none',
              cursor: 'pointer',
              padding: '0.5rem',
              marginLeft: 'auto',
              zIndex: 120,
            }}
            aria-label="Toggle navigation menu"
          >
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#22d3ee" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
              <line x1="3" y1="6" x2="21" y2="6"></line>
              <line x1="3" y1="12" x2="21" y2="12"></line>
              <line x1="3" y1="18" x2="21" y2="18"></line>
            </svg>
          </button>
          {/* Mobile Dropdown Menu */}
          <div className={`admin-mobile-menu ${navOpen ? 'open' : ''}`} style={{
            display: 'none',
            position: 'absolute',
            top: '100%',
            right: 0,
            background: 'rgba(26,26,46,0.98)',
            boxShadow: '0 8px 32px 0 rgba(34,211,238,0.13), 0 1.5px 12px 0 #0ea5e955',
            borderRadius: '0 0 18px 18px',
            zIndex: 110,
            padding: '1.2rem 1.5rem',
            gap: '1.1rem',
            flexDirection: 'column',
            alignItems: 'stretch',
            minWidth: '180px',
          }}>
            <a href="#dashboard" className="admin-mobile-link" style={{ color: '#cbd5e1', textDecoration: 'none', padding: '0.8rem 1rem', borderRadius: '8px', transition: 'background 0.2s', fontWeight: 600 }} onClick={e => { e.preventDefault(); setShowHistory(false); setShowTimetable(false); setNavOpen(false); }}>Dashboard</a>
            <a href="#history" className="admin-mobile-link" style={{ color: '#cbd5e1', textDecoration: 'none', padding: '0.8rem 1rem', borderRadius: '8px', transition: 'background 0.2s', fontWeight: 600 }} onClick={e => { e.preventDefault(); setShowHistory(true); setShowTimetable(false); setNavOpen(false); }}>History</a>
            <a href="#timetable" className="admin-mobile-link" style={{ color: '#cbd5e1', textDecoration: 'none', padding: '0.8rem 1rem', borderRadius: '8px', transition: 'background 0.2s', fontWeight: 600 }} onClick={e => { e.preventDefault(); setShowHistory(false); setShowTimetable(true); setNavOpen(false); }}>Timetable</a>
            <button className="admin-mobile-logout" onClick={() => { onBack(); setNavOpen(false); }} style={{
              background: 'linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%)',
              color: '#fff',
              border: 'none',
              borderRadius: '8px',
              padding: '0.8rem 1rem',
              fontWeight: 700,
              fontSize: '1rem',
              cursor: 'pointer',
              boxShadow: '0 2px 8px rgba(34,211,238,0.13)',
              transition: 'background 0.2s, transform 0.2s',
              letterSpacing: '0.3px',
            }}>Logout</button>
          </div>
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
      ) : showTimetable ? (
        <TimetableUpload
          onBack={() => setShowTimetable(false)}
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
            {adminUser ? (
              <div style={{
                fontSize: '1rem',
                color: '#38bdf8',
                fontWeight: 600,
                marginTop: '0.3rem',
                marginBottom: '0.2rem',
                textAlign: 'center',
                letterSpacing: '0.2px',
                animation: 'fadeInUp 1s ease-out 0.3s both',
                position: 'relative',
                display: 'inline-block',
              }}>
                Welcome, {adminUser.username}!
                <span style={{
                  display: 'block',
                  width: '40px',
                  height: '2px',
                  margin: '0.2rem auto 0 auto',
                  borderRadius: '1px',
                  background: 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)',
                  opacity: 0.6,
                  animation: 'pulseBar 2.2s infinite',
                }} />
              </div>
            ) : (
              <p style={{ fontSize: '0.92rem', color: '#cbd5e1', marginTop: '0.4rem', fontWeight: 500, letterSpacing: '0.3px', animation: 'fadeInUp 1s ease-out 0.3s both' }}>
                Welcome, <span style={{ color: '#fff', fontWeight: 700 }}>UMT Admin</span>! Manage user signups and approve or reject new accounts below.
              </p>
            )}
            {/* Stats Bar */}
            <div style={{
              display: 'flex',
              justifyContent: 'center',
              alignItems: 'center',
              gap: '1.1rem',
              margin: '1.1rem auto 0 auto',
              padding: '1.1rem 3.5rem',
              background: 'rgba(24,24,27,0.97)',
              borderRadius: '1rem',
              boxShadow: '0 2px 8px rgba(56,189,248,0.10)',
              maxWidth: 900,
              width: '100%',
              animation: 'fadeInUpAdminStats 1.1s cubic-bezier(.39,.575,.56,1.000) 0.5s both',
              flexWrap: 'wrap',
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
            maxWidth: 1300,
            margin: '0 auto',
            padding: '4rem 3rem',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            zIndex: 1,
            overflowX: 'auto',
            boxSizing: 'border-box',
          }}>
            <section style={{ width: '100%' }}>
              <h2 style={{ color: '#38bdf8', fontSize: '1.4rem', fontWeight: 700, marginBottom: '2.5rem', textAlign: 'center', letterSpacing: '-0.5px', animation: 'glow 2s ease-in-out infinite alternate 1s', wordWrap: 'break-word' }}>
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
        @keyframes dropdownFadeIn {
          from { opacity: 0; transform: translateY(-10px); }
          to { opacity: 1; transform: translateY(0); }
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
          /* Mobile responsive stats bar */
          header > div[style*='display: flex'] {
            flex-direction: column !important;
            gap: 0.8rem !important;
            padding: 0.8rem 0.2rem !important;
            max-width: 100vw !important;
            align-items: center !important;
          }
          header > div[style*='display: flex'] > div {
            min-width: 120px !important;
            padding: 0.8rem 1rem !important;
            font-size: 1rem !important;
            width: 100% !important;
            max-width: 280px !important;
          }
          header > div[style*='display: flex'] > div > span:first-child {
            font-size: 1.2rem !important;
          }
          header > div[style*='display: flex'] > div > span:nth-child(2) {
            font-size: 1.1rem !important;
          }
          header > div[style*='display: flex'] > div > span:last-child {
            font-size: 0.9rem !important;
          }
          /* Mobile responsive navbar */
          nav {
            padding: 0.5rem 0.2rem !important;
            min-height: 35px !important;
          }
          .admin-navbar-links {
            gap: 0.5rem !important;
            font-size: 0.85rem !important;
            margin-right: 0.5rem !important;
          }
          .admin-navbar-logout {
            padding: 0.3rem 0.7rem !important;
            font-size: 0.85rem !important;
          }
          /* Hamburger menu for mobile */
          .admin-hamburger {
            display: block !important;
          }
          .admin-navbar-links {
            display: none !important;
          }
          .admin-navbar-logout {
            display: none !important;
          }
          .admin-mobile-menu {
            display: none !important;
          }
          .admin-mobile-menu.open {
            display: flex !important;
            animation: dropdownFadeIn 0.25s cubic-bezier(.39,.575,.56,1.000);
          }
          .admin-mobile-link:hover {
            background: rgba(56,189,248,0.10) !important;
          }
          .admin-mobile-logout:hover {
            background: linear-gradient(90deg, #38bdf8 0%, #22d3ee 100%) !important;
            transform: translateY(-2px) scale(1.04) !important;
          }
          /* Mobile responsive header */
          header {
            padding: 0.7rem 0 0.5rem 0 !important;
            margin-bottom: 0.7rem !important;
          }
          header h1 {
            font-size: 1rem !important;
            margin-bottom: 0.2rem !important;
          }
          header p {
            font-size: 0.8rem !important;
            margin-top: 0.2rem !important;
          }
          /* Mobile responsive main content */
          main {
            padding: 1rem 0.2rem !important;
          }
          main h2 {
            font-size: 1.1rem !important;
            margin-bottom: 1.5rem !important;
          }
          /* Mobile responsive table */
          table {
            font-size: 0.85rem !important;
            min-width: 300px !important;
          }
          table th, table td {
            padding: 0.5rem 0.3rem !important;
            font-size: 0.8rem !important;
          }
          table th {
            font-size: 0.85rem !important;
          }
          /* Mobile responsive buttons */
          button {
            padding: 0.25rem 0.5rem !important;
            font-size: 0.75rem !important;
            margin-right: 0.3rem !important;
          }
        }
        @media (max-width: 700px) {
          /* Enhanced mobile responsive stats bar */
          header > div[style*='display: flex'] {
            flex-direction: column !important;
            gap: 0.6rem !important;
            padding: 0.6rem 0.1rem !important;
            align-items: center !important;
          }
          header > div[style*='display: flex'] > div {
            min-width: 100px !important;
            padding: 0.6rem 0.8rem !important;
            font-size: 0.9rem !important;
            width: 100% !important;
            max-width: 250px !important;
          }
          /* Enhanced hamburger menu for mobile */
          .admin-hamburger {
            display: block !important;
            padding: 0.3rem !important;
          }
          .admin-hamburger svg {
            width: 20px !important;
            height: 20px !important;
          }
          header > div[style*='display: flex'] > div > span:first-child {
            font-size: 1.1rem !important;
          }
          header > div[style*='display: flex'] > div > span:nth-child(2) {
            font-size: 1rem !important;
          }
          header > div[style*='display: flex'] > div > span:last-child {
            font-size: 0.8rem !important;
          }
          /* Enhanced mobile responsive navbar */
          nav {
            padding: 0.3rem 0.1rem !important;
            min-height: 30px !important;
          }
          .admin-navbar-links {
            gap: 0.3rem !important;
            font-size: 0.75rem !important;
            margin-right: 0.3rem !important;
          }
          .admin-navbar-logout {
            padding: 0.2rem 0.5rem !important;
            font-size: 0.75rem !important;
          }
          /* Enhanced mobile responsive header */
          header {
            padding: 0.5rem 0 0.3rem 0 !important;
            margin-bottom: 0.5rem !important;
          }
          header h1 {
            font-size: 0.9rem !important;
            margin-bottom: 0.1rem !important;
          }
          header p {
            font-size: 0.7rem !important;
            margin-top: 0.1rem !important;
          }
          /* Enhanced mobile responsive main content */
          main {
            padding: 0.5rem 0.1rem !important;
          }
          main h2 {
            font-size: 1rem !important;
            margin-bottom: 1rem !important;
          }
          /* Enhanced mobile responsive table */
          table {
            font-size: 0.75rem !important;
            min-width: 250px !important;
          }
          table th, table td {
            padding: 0.3rem 0.2rem !important;
            font-size: 0.7rem !important;
          }
          table th {
            font-size: 0.75rem !important;
          }
          /* Enhanced mobile responsive buttons */
          button {
            padding: 0.2rem 0.4rem !important;
            font-size: 0.65rem !important;
            margin-right: 0.2rem !important;
          }
        }
        @media (max-width: 500px) {
          /* Ultra mobile responsive stats bar */
          header > div[style*='display: flex'] {
            flex-direction: column !important;
            gap: 0.5rem !important;
            padding: 0.5rem 0.05rem !important;
            align-items: center !important;
          }
          header > div[style*='display: flex'] > div {
            min-width: 80px !important;
            padding: 0.5rem 0.6rem !important;
            font-size: 0.8rem !important;
            width: 100% !important;
            max-width: 220px !important;
          }
          /* Ultra mobile hamburger menu */
          .admin-hamburger {
            display: block !important;
            padding: 0.2rem !important;
          }
          .admin-hamburger svg {
            width: 18px !important;
            height: 18px !important;
          }
          header > div[style*='display: flex'] > div > span:first-child {
            font-size: 1rem !important;
          }
          header > div[style*='display: flex'] > div > span:nth-child(2) {
            font-size: 0.9rem !important;
          }
          header > div[style*='display: flex'] > div > span:last-child {
            font-size: 0.7rem !important;
          }
          /* Ultra mobile responsive navbar */
          nav {
            padding: 0.2rem 0.05rem !important;
            min-height: 25px !important;
          }
          .admin-navbar-links {
            gap: 0.2rem !important;
            font-size: 0.65rem !important;
            margin-right: 0.2rem !important;
          }
          .admin-navbar-logout {
            padding: 0.15rem 0.4rem !important;
            font-size: 0.65rem !important;
          }
          /* Ultra mobile responsive header */
          header {
            padding: 0.3rem 0 0.2rem 0 !important;
            margin-bottom: 0.3rem !important;
          }
          header h1 {
            font-size: 0.8rem !important;
            margin-bottom: 0.05rem !important;
          }
          header p {
            font-size: 0.6rem !important;
            margin-top: 0.05rem !important;
          }
          /* Ultra mobile responsive main content */
          main {
            padding: 0.3rem 0.05rem !important;
          }
          main h2 {
            font-size: 0.9rem !important;
            margin-bottom: 0.7rem !important;
          }
          /* Ultra mobile responsive table */
          table {
            font-size: 0.65rem !important;
            min-width: 200px !important;
          }
          table th, table td {
            padding: 0.2rem 0.1rem !important;
            font-size: 0.6rem !important;
          }
          table th {
            font-size: 0.65rem !important;
          }
          /* Ultra mobile responsive buttons */
          button {
            padding: 0.15rem 0.3rem !important;
            font-size: 0.55rem !important;
            margin-right: 0.15rem !important;
          }
        }
      `}</style>
    </div>
  );
}

export default AdminDashboard; 