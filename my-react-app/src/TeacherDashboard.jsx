import React, { useState, useEffect } from "react";
import { ArrowLeft, Users, Calendar, BookOpen, Info } from "lucide-react";

const teacherProfile = {
  name: "Prof. Ahmed Khan",
  email: "teacher@umt.edu.pk",
  department: "Computer Science",
  avatar: "https://randomuser.me/api/portraits/men/32.jpg",
};

// Helper function to download CSV
function downloadCSV(data, filename) {
  const csvRows = [
    ['Day', 'Classes'],
    ...data.map(row => [row.day, row.classes.join(', ')]),
  ];
  const csvContent = csvRows.map(e => e.join(',')).join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv' });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.setAttribute('hidden', '');
  a.setAttribute('href', url);
  a.setAttribute('download', filename);
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// Placeholder for API timetable data
const sampleTimetable = [
  { day: 'Monday', classes: ['CS101', 'CS201'] },
  { day: 'Tuesday', classes: ['CS102', 'CS202'] },
  { day: 'Wednesday', classes: ['CS103', 'CS203'] },
  { day: 'Thursday', classes: ['CS104', 'CS204'] },
  { day: 'Friday', classes: ['CS105', 'CS205'] },
];

function TeacherDashboard({ onBack, teacherUser }) {
  // Start with no timetable; will be set by API in the future
  const [timetable, setTimetable] = useState([]);
  // For demo/testing, you can uncomment the next line:
  // const [timetable, setTimetable] = useState(sampleTimetable);

  // Only use teacherUser, never fallback to default
  const profile = teacherUser ? {
    name: teacherUser.username,
    email: teacherUser.email,
    department: teacherUser.department || '',
  } : null;

  // In the future, fetch from API here
  // useEffect(() => {
  //   fetch('/api/teacher-timetable')
  //     .then(res => res.json())
  //     .then(data => setTimetable(data));
  // }, []);

  return (
    <div style={{
      minHeight: '100vh',
      width: '100vw',
      background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'flex-start',
      position: 'relative',
      overflow: 'hidden',
    }}>
      {/* Decorative Blobs */}
      <div style={{
        position: 'absolute',
        top: '8%',
        left: '8%',
        width: '220px',
        height: '220px',
        background: 'radial-gradient(circle, rgba(34,211,238,0.13) 0%, transparent 70%)',
        borderRadius: '50%',
        filter: 'blur(60px)',
        opacity: 0.13,
        pointerEvents: 'none',
        zIndex: 0,
      }} />
      <div style={{
        position: 'absolute',
        bottom: '10%',
        right: '10%',
        width: '280px',
        height: '280px',
        background: 'radial-gradient(circle, rgba(59,130,246,0.13) 0%, transparent 70%)',
        borderRadius: '50%',
        filter: 'blur(60px)',
        opacity: 0.13,
        pointerEvents: 'none',
        zIndex: 0,
      }} />
      {/* Topbar */}
      <header style={{
        width: '100%',
        padding: '2.2rem 0 1.2rem 0',
        textAlign: 'center',
        background: '#23272f',
        color: '#fff',
        boxShadow: '0 2px 12px rgba(30,41,59,0.10)',
        marginBottom: '2.5rem',
        position: 'relative',
        zIndex: 1,
        animation: 'slideDown 0.8s ease-out',
      }}>
        {onBack && (
          <button
            onClick={onBack}
            style={{
              position: 'absolute',
              top: 24,
              left: 32,
              background: 'rgba(34,211,238,0.10)',
              border: 'none',
              borderRadius: '50%',
              width: 44,
              height: 44,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              cursor: 'pointer',
              boxShadow: '0 2px 8px rgba(34,211,238,0.10)',
              transition: 'background 0.18s, transform 0.18s',
              zIndex: 2,
              outline: 'none',
              animation: 'fadeInUp 1s ease-out 0.2s both',
            }}
            aria-label="Back"
          >
            <ArrowLeft size={26} color="#22d3ee" />
        </button>
        )}
        <h1 style={{ fontSize: '2.1rem', fontWeight: 800, margin: 0, letterSpacing: '-1px', textShadow: '0 2px 12px #23272f55', animation: 'glow 2s ease-in-out infinite alternate' }}>Teacher Dashboard</h1>
        <p style={{ fontSize: '1.08rem', color: '#cbd5e1', marginTop: '0.7rem', fontWeight: 500, letterSpacing: '0.5px', animation: 'fadeInUp 1s ease-out 0.3s both' }}>
          Welcome, <span style={{ color: '#22d3ee', fontWeight: 700 }}>Professor</span>! Manage your classes, schedule, and resources below.
        </p>
      </header>
      <main style={{
        width: '100%',
        maxWidth: 900,
        margin: '0 auto',
        padding: '2.5rem 1.5rem',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        zIndex: 1,
        gap: '2.5rem',
      }}>
        {/* Profile Card */}
        {/* Removed profile card as per user request */}
      {/* Welcome Message */}
        <section style={{ width: '100%', maxWidth: 700, marginBottom: '1.5rem' }}>
          <div style={{
            background: 'rgba(34,211,238,0.07)',
            borderRadius: '14px',
            padding: '1.5rem 2rem',
            color: '#22d3ee',
            fontWeight: 500,
            fontSize: '1.13rem',
            textAlign: 'center',
            boxShadow: '0 2px 8px rgba(30,41,59,0.07)',
            animation: 'fadeInUp 1s',
          }}>
          Here you can manage your classes, view your schedule, and connect with your students. Smart Scheduler makes your teaching experience seamless and organized.
      </div>
        </section>
      {/* Teacher Features Grid */}
        <section style={{ width: '100%', maxWidth: 900 }}>
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(260px, 1fr))',
            gap: '2rem',
            width: '100%',
            margin: '0 auto',
            animation: 'fadeInUp 1s',
          }}>
            <div style={{ background: 'rgba(34,211,238,0.08)', borderRadius: 14, padding: '2rem', textAlign: 'center', border: '1px solid rgba(34,211,238,0.15)', boxShadow: '0 2px 12px rgba(34,211,238,0.05)', color: '#f1f5f9', transition: 'transform 0.2s, box-shadow 0.2s' }}>
          <Users size={36} color="#22d3ee" style={{ marginBottom: 10 }} />
              <div style={{ fontWeight: 600, fontSize: '1.1rem', marginBottom: 6 }}>Class Management</div>
              <div style={{ color: '#94a3b8', fontSize: '0.98rem' }}>View, add, or edit your classes and student lists.</div>
        </div>
            <div style={{ background: 'rgba(52,211,153,0.08)', borderRadius: 14, padding: '2rem', textAlign: 'center', border: '1px solid rgba(52,211,153,0.15)', boxShadow: '0 2px 12px rgba(52,211,153,0.05)', color: '#f1f5f9', transition: 'transform 0.2s, box-shadow 0.2s' }}>
          <Calendar size={36} color="#34d399" style={{ marginBottom: 10 }} />
              <div style={{ fontWeight: 600, fontSize: '1.1rem', marginBottom: 6 }}>Schedule Overview</div>
              <div style={{ color: '#94a3b8', fontSize: '0.98rem' }}>See your upcoming classes and important dates.</div>
        </div>
            <div style={{ background: 'rgba(59,130,246,0.08)', borderRadius: 14, padding: '2rem', textAlign: 'center', border: '1px solid rgba(59,130,246,0.15)', boxShadow: '0 2px 12px rgba(59,130,246,0.05)', color: '#f1f5f9', transition: 'transform 0.2s, box-shadow 0.2s' }}>
          <BookOpen size={36} color="#3b82f6" style={{ marginBottom: 10 }} />
              <div style={{ fontWeight: 600, fontSize: '1.1rem', marginBottom: 6 }}>Resources</div>
              <div style={{ color: '#94a3b8', fontSize: '0.98rem' }}>Upload and share course materials with your students.</div>
        </div>
      </div>
        </section>
        {/* Timetable Section */}
        <section style={{ width: '100%', maxWidth: 800, margin: '2.5rem auto 0 auto' }}>
          {/* Info Banner */}
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '1rem',
            background: 'rgba(34,211,238,0.10)',
            borderRadius: 14,
            padding: '1.1rem 2rem',
            boxShadow: '0 2px 12px rgba(34,211,238,0.07)',
            maxWidth: 600,
            width: '100%',
            margin: '0 auto 2.2rem auto',
          }}>
            <Info size={32} color="#22d3ee" style={{ flexShrink: 0 }} />
            <div style={{ fontWeight: 600, fontSize: '1.18rem', color: '#22d3ee', letterSpacing: '0px' }}>
              Your timetable will appear below when available. Download it for easy access!
            </div>
          </div>
          <div style={{
            background: "rgba(35,39,47,0.92)",
            borderRadius: "22px",
            boxShadow: "0 12px 48px rgba(34,211,238,0.18)",
            padding: "3rem 2.5rem 2.5rem 2.5rem",
            marginBottom: "2.5rem",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            animation: "fadeInUp 1s",
            border: '2.5px solid rgba(34,211,238,0.18)',
            backdropFilter: 'blur(14px)',
            minHeight: 320,
          }}>
            <h2 style={{ color: "#22d3ee", fontWeight: 800, fontSize: "2rem", marginBottom: "1.6rem", letterSpacing: "-1px", textShadow: '0 2px 12px #22d3ee55' }}>Your Timetable</h2>
            {timetable.length === 0 ? (
              <div style={{ color: '#94a3b8', fontSize: '1.25rem', padding: '3rem 0', textAlign: 'center', fontWeight: 500 }}>
                No timetable available yet.<br />Please check back later.
              </div>
            ) : (
              <>
                <div style={{ width: "100%", overflowX: "auto", marginBottom: "2.2rem" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", background: "rgba(34,211,238,0.03)", borderRadius: "12px", color: "#cbd5e1", fontSize: "1.18rem", boxShadow: '0 2px 8px rgba(34,211,238,0.07)' }}>
                    <thead>
                      <tr style={{ background: "rgba(34,211,238,0.10)" }}>
                        <th style={{ padding: "1.1rem 1.5rem", color: "#22d3ee", fontWeight: 800, textAlign: "left", borderBottom: "2px solid #23272f", fontSize: '1.18rem', letterSpacing: '-0.5px' }}>Day</th>
                        <th style={{ padding: "1.1rem 1.5rem", color: "#22d3ee", fontWeight: 800, textAlign: "left", borderBottom: "2px solid #23272f", fontSize: '1.18rem', letterSpacing: '-0.5px' }}>Classes</th>
                      </tr>
                    </thead>
                    <tbody>
                      {timetable.map((row, i) => (
                        <tr key={row.day} style={{ borderBottom: "1px solid #23272f" }}>
                          <td style={{ padding: "1.1rem 1.5rem", fontWeight: 700 }}>{row.day}</td>
                          <td style={{ padding: "1.1rem 1.5rem" }}>{row.classes.join(", ")}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <button
                  onClick={() => downloadCSV(timetable, "teacher-timetable.csv")}
                  style={{
                    background: "linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%)",
                    color: "#fff",
                    border: "none",
                    borderRadius: "14px",
                    padding: "1.2rem 2.8rem",
                    fontWeight: 800,
                    fontSize: "1.18rem",
                    cursor: "pointer",
                    opacity: 1,
                    boxShadow: "0 6px 24px rgba(34,211,238,0.18)",
                    marginTop: "0.5rem",
                    letterSpacing: "0.7px",
                    transition: "all 0.2s",
                    outline: 'none',
                  }}
                  onMouseOver={e => e.currentTarget.style.background = 'linear-gradient(90deg, #06b6d4 0%, #22d3ee 100%)'}
                  onMouseOut={e => e.currentTarget.style.background = 'linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%)'}
                  onFocus={e => e.currentTarget.style.boxShadow = '0 0 0 4px #22d3eeaa'}
                  onBlur={e => e.currentTarget.style.boxShadow = '0 6px 24px rgba(34,211,238,0.18)'}
                >
                  Download Timetable
                </button>
              </>
            )}
          </div>
        </section>
      </main>
      <style>{`
        @keyframes slideDown {
          from { opacity: 0; transform: translateY(-20px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeInUp {
          from { opacity: 0; transform: translateY(30px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes glow {
          0% { text-shadow: 0 0 5px #22d3ee55; }
          100% { text-shadow: 0 0 20px #22d3eecc, 0 0 30px #22d3ee99; }
        }
        button:hover {
          background: #22d3ee;
          transform: scale(1.08) translateX(-2px);
        }
        button:hover svg {
          stroke: #fff;
        }
        div[style*='rgba(34,211,238,0.08)']:hover,
        div[style*='rgba(52,211,153,0.08)']:hover,
        div[style*='rgba(59,130,246,0.08)']:hover {
          box-shadow: 0 8px 32px rgba(34,211,238,0.13);
          transform: translateY(-2px) scale(1.04);
        }
      `}</style>
    </div>
  );
}

export default TeacherDashboard; 