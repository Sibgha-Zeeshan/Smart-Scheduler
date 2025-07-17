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
      background: 'linear-gradient(120deg, #0f172a 0%, #1e293b 100%)',
      fontFamily: 'Inter, Roboto, system-ui, sans-serif',
      display: 'flex',
      flexDirection: 'column',
      alignItems: 'center',
      justifyContent: 'flex-start',
      position: 'relative',
      overflow: 'hidden',
    }}>
      {/* Decorative Blobs */}
      {/* Animated SVG Blob */}
      <svg style={{position:'absolute',top:'-80px',left:'-120px',zIndex:0,opacity:0.22,filter:'blur(2px)',animation:'blobMove 18s ease-in-out infinite'}} width="420" height="420" viewBox="0 0 420 420" fill="none" xmlns="http://www.w3.org/2000/svg">
        <defs>
          <linearGradient id="blobGrad" x1="0" y1="0" x2="1" y2="1">
            <stop offset="0%" stopColor="#38bdf8"/>
            <stop offset="100%" stopColor="#0ea5e9"/>
          </linearGradient>
        </defs>
        <path fill="url(#blobGrad)" d="M320,60Q380,120,340,200Q300,280,200,320Q100,360,60,260Q20,160,100,100Q180,40,320,60Z"/>
      </svg>
      <svg style={{position:'absolute',bottom:'-100px',right:'-120px',zIndex:0,opacity:0.18,filter:'blur(2px)',animation:'blobMove2 22s ease-in-out infinite'}} width="420" height="420" viewBox="0 0 420 420" fill="none" xmlns="http://www.w3.org/2000/svg">
        <defs>
          <linearGradient id="blobGrad2" x1="0" y1="0" x2="1" y2="1">
            <stop offset="0%" stopColor="#6366f1"/>
            <stop offset="100%" stopColor="#38bdf8"/>
          </linearGradient>
        </defs>
        <path fill="url(#blobGrad2)" d="M320,60Q380,120,340,200Q300,280,200,320Q100,360,60,260Q20,160,100,100Q180,40,320,60Z"/>
      </svg>
      {/* Topbar */}
      <header style={{
        width: '100%',
        padding: '1.2rem 0 0.7rem 0', // reduced
        textAlign: 'center',
        background: '#23272f',
        color: '#fff',
        boxShadow: '0 2px 12px rgba(30,41,59,0.10)',
        marginBottom: '1.2rem', // reduced
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
            <ArrowLeft size={100} color="#22d3ee" style={{ display: 'block' }} />
          </button>
        )}
        <h1
          className="teacher-dashboard-heading"
          style={{
            fontSize: '1.2rem',
            fontWeight: 900,
            margin: 0,
            letterSpacing: '-1px',
            background: 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)',
            WebkitBackgroundClip: 'text',
            WebkitTextFillColor: 'transparent',
            backgroundClip: 'text',
            textAlign: 'center',
            position: 'relative',
            animation: 'slideInLeft 1.2s',
            lineHeight: 1.05,
            textShadow: '0 2px 16px #38bdf855, 0 0 24px #0ea5e999',
            filter: 'drop-shadow(0 2px 8px #38bdf855)',
          }}
        >
          Teacher Dashboard
          <span
            style={{
              display: 'block',
              width: 44,
              height: 5,
              margin: '0.3rem auto 0 auto',
              borderRadius: 3,
              background: 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)',
              boxShadow: '0 0 16px #38bdf855',
              opacity: 0.85,
              animation: 'pulseBar 2.5s infinite',
            }}
          />
        </h1>
        <div
          className="teacher-dashboard-subheading"
          style={{
            display: 'block',
            fontSize: '0.95rem',
            color: '#38bdf8',
            fontWeight: 500,
            margin: '0.3rem auto 0.2rem auto',
            textAlign: 'center',
            lineHeight: 1.2,
          }}
        >
          Download conflict free classes, schedule, and resources below.
        </div>
      </header>
      <main style={{
        width: '100%',
        maxWidth: 900,
        margin: '0 auto',
        padding: '1.2rem 0.7rem', // reduced
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        zIndex: 1,
        gap: '1.2rem', // reduced
        boxSizing: 'border-box',
      }}>
        {/* Profile Card */}
        {/* Removed profile card as per user request */}
      {/* Welcome Message */}
        <section style={{ width: '100%', maxWidth: 700, marginBottom: '1.5rem' }}>
          {/* Removed the div containing the text */}
        </section>
        {/* Timetable Section */}
        <section style={{ width: '100%', maxWidth: 800, margin: '1.2rem auto 0 auto' }}>
          {/* Info Banner */}
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '0.3rem',
            background: 'rgba(34,211,238,0.10)',
            borderRadius: 12,
            padding: '0.1rem 0.3rem',
            boxShadow: '0 2px 8px rgba(34,211,238,0.07)',
            width: '100%',
            margin: '0 auto 1.1rem auto',
            fontSize: '0.68rem',
            fontWeight: 500,
            textAlign: 'center',
            justifyContent: 'center',
            animation: 'fadeInUp 1s',
          }}>
            <Info size={32} color="#22d3ee" style={{ flexShrink: 0 }} />
            <div style={{ fontWeight: 600, fontSize: '0.68rem', color: '#22d3ee', letterSpacing: '0px' }}>
              Your timetable will appear below when available. Download it for easy access!
            </div>
          </div>
          <div style={{
            background: "rgba(35,39,47,0.82)",
            borderRadius: "28px",
            boxShadow: "0 8px 48px 0 rgba(34,211,238,0.22), 0 1.5px 12px 0 #0ea5e955",
            padding: "1.2rem 0.7rem 1.1rem 0.7rem",
            marginBottom: "1.2rem",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            animation: "fadeInUp 1s",
            border: '2.5px solid rgba(34,211,238,0.22)',
            backdropFilter: 'blur(22px)',
            minHeight: 180,
            filter: 'drop-shadow(0 0 24px #38bdf855)',
          }}>
            {/* Personalized Welcome Message */}
            <div style={{
              fontSize: '1.1rem',
              fontWeight: 700,
              color: '#38bdf8',
              marginBottom: '0.5rem',
              letterSpacing: '-0.5px',
              textAlign: 'center',
              background: 'rgba(34,211,238,0.08)',
              borderRadius: '999px',
              padding: '0.3em 1.2em',
              boxShadow: '0 2px 12px #38bdf822',
              border: '1.5px solid #38bdf822',
              backdropFilter: 'blur(2px)',
              display: 'inline-block',
            }}>
              {`Welcome, ${profile && profile.name ? profile.name : 'Professor'}!`}
            </div>
            <h2 style={{ color: "#22d3ee", fontWeight: 800, fontSize: "1.3rem", marginBottom: "0.7rem", letterSpacing: "-1px", textShadow: '0 2px 12px #22d3ee55' }}>Your Timetable</h2>
            {timetable.length === 0 ? (
              <div style={{ color: '#94a3b8', fontSize: '1.1rem', padding: '1.2rem 0', textAlign: 'center', fontWeight: 500 }}>
                No timetable available yet.<br />Please check back later.
              </div>
            ) : (
              <>
                <div style={{ width: "100%", overflowX: "auto", marginBottom: "1.1rem" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", background: "rgba(34,211,238,0.06)", borderRadius: "14px", color: "#cbd5e1", fontSize: "1.08rem", boxShadow: '0 2px 16px rgba(34,211,238,0.13)', overflow: 'hidden' }}>
                    <thead>
                      <tr style={{ background: "rgba(34,211,238,0.13)" }}>
                        <th style={{ padding: "0.7rem 0.9rem", color: "#22d3ee", fontWeight: 800, textAlign: "left", borderBottom: "2px solid #23272f", fontSize: '1rem', letterSpacing: '-0.5px', textShadow: '0 1px 6px #22d3ee33' }}>Day</th>
                        <th style={{ padding: "0.7rem 0.9rem", color: "#22d3ee", fontWeight: 800, textAlign: "left", borderBottom: "2px solid #23272f", fontSize: '1rem', letterSpacing: '-0.5px', textShadow: '0 1px 6px #22d3ee33' }}>Classes</th>
                      </tr>
                    </thead>
                    <tbody>
                      {timetable.map((row, i) => (
                        <tr key={row.day} style={{ borderBottom: "1px solid #23272f", transition: 'background 0.18s' }}
                           onMouseOver={e => e.currentTarget.style.background = 'rgba(34,211,238,0.08)'}
                           onMouseOut={e => e.currentTarget.style.background = ''}>
                          <td style={{ padding: "0.7rem 0.9rem", fontWeight: 700 }}>{row.day}</td>
                          <td style={{ padding: "0.7rem 0.9rem" }}>{row.classes.join(", ")}</td>
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
                    padding: "0.7rem 1.5rem", // reduced
                    fontWeight: 800,
                    fontSize: "1rem",
                    cursor: "pointer",
                    opacity: 1,
                    boxShadow: "0 6px 24px rgba(34,211,238,0.18)",
                    marginTop: "0.5rem",
                    letterSpacing: "0.7px",
                    transition: "all 0.2s",
                    outline: 'none',
                    animation: 'pulseGlow 2.2s infinite',
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
        @keyframes pulseBar {
          0%, 100% { opacity: 0.85; box-shadow: 0 0 16px #38bdf855; }
          50% { opacity: 1; box-shadow: 0 0 32px #38bdf8cc; }
        }
        @keyframes pulseGlow {
          0%, 100% { box-shadow: 0 6px 24px rgba(34,211,238,0.18), 0 0 0 0 #22d3ee55; }
          50% { box-shadow: 0 6px 32px rgba(34,211,238,0.28), 0 0 0 8px #22d3ee33; }
        }
        @keyframes blobMove {
          0%, 100% { transform: scale(1) translateY(0px) translateX(0px); }
          50% { transform: scale(1.08) translateY(30px) translateX(40px); }
        }
        @keyframes blobMove2 {
          0%, 100% { transform: scale(1) translateY(0px) translateX(0px); }
          50% { transform: scale(1.12) translateY(-30px) translateX(-40px); }
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
        
        /* Mobile Responsive Styles */
        @media (max-width: 768px) {
          /* Header adjustments */
          header {
            padding: 0.8rem 0 0.4rem 0 !important;
            margin-bottom: 2rem !important;
          }
          
          /* Back button adjustments */
          header button {
            top: 14px !important;
            left: 14px !important;
            width: 32px !important;
            height: 32px !important;
          }
          
          /* Heading adjustments */
          .teacher-dashboard-heading {
            font-size: 1rem !important;
            margin: 0 3rem !important;
          }
          
          .teacher-dashboard-subheading {
            font-size: 0.8rem !important;
            margin: 0.5rem 2rem 0.3rem 2rem !important;
          }
          
          /* Main content adjustments */
          main {
            padding: 1.5rem 1rem 1.5rem 3rem !important;
            gap: 2rem !important;
            margin-right: 1rem !important;
            border-right: 2px solid rgba(34,211,238,0.2) !important;
          }
          
          /* Section adjustments */
          section {
            margin-bottom: 2rem !important;
          }
          
          /* Info banner adjustments */
          div[style*='background: rgba(34,211,238,0.10)'] {
            padding: 0.5rem 0.8rem !important;
            font-size: 0.7rem !important;
            margin-bottom: 1.5rem !important;
          }
          
          /* Main card adjustments */
          div[style*='background: rgba(35,39,47,0.82)'] {
            padding: 1.5rem 1rem !important;
            border-radius: 16px !important;
            min-height: 140px !important;
            margin: 0 1rem !important;
          }
          
          /* Welcome message adjustments */
          div[style*='font-size: 1.1rem'][style*='color: #38bdf8'] {
            font-size: 0.9rem !important;
            padding: 0.3em 1em !important;
            margin-bottom: 1rem !important;
          }
          
          /* Table heading adjustments */
          h2[style*='color: #22d3ee'] {
            font-size: 1rem !important;
            margin-bottom: 1rem !important;
          }
          
          /* Table adjustments */
          table {
            font-size: 0.85rem !important;
            border-radius: 8px !important;
            margin: 0 0.5rem !important;
          }
          
          th, td {
            padding: 0.6rem 0.7rem !important;
            font-size: 0.85rem !important;
          }
          
          /* Download button adjustments */
          button[style*='background: linear-gradient(90deg, #22d3ee'] {
            padding: 0.6rem 1.2rem !important;
            font-size: 0.85rem !important;
            border-radius: 8px !important;
            margin-top: 1rem !important;
          }
          
          /* Decorative blobs adjustments */
          svg[style*='position: absolute'] {
            width: 300px !important;
            height: 300px !important;
          }
        }
        
        @media (max-width: 480px) {
          /* Header adjustments */
          header {
            padding: 0.7rem 0 0.3rem 0 !important;
            margin-bottom: 1.5rem !important;
          }
          
          /* Back button adjustments */
          header button {
            top: 12px !important;
            left: 12px !important;
            width: 30px !important;
            height: 30px !important;
          }
          
          /* Heading adjustments */
          .teacher-dashboard-heading {
            font-size: 0.9rem !important;
            margin: 0 2.5rem !important;
          }
          
          .teacher-dashboard-subheading {
            font-size: 0.75rem !important;
            margin: 0.4rem 1.8rem 0.2rem 1.8rem !important;
          }
          
          /* Main content adjustments */
          main {
            padding: 1.2rem 0.8rem 1.2rem 2.5rem !important;
            gap: 1.5rem !important;
            margin-right: 0.8rem !important;
            border-right: 2px solid rgba(34,211,238,0.2) !important;
          }
          
          /* Info banner adjustments */
          div[style*='background: rgba(34,211,238,0.10)'] {
            padding: 0.4rem 0.7rem !important;
            font-size: 0.65rem !important;
            margin-bottom: 1.2rem !important;
          }
          
          /* Main card adjustments */
          div[style*='background: rgba(35,39,47,0.82)'] {
            padding: 1.2rem 0.8rem !important;
            border-radius: 12px !important;
            min-height: 120px !important;
            margin: 0 0.8rem !important;
          }
          
          /* Welcome message adjustments */
          div[style*='font-size: 1.1rem'][style*='color: #38bdf8'] {
            font-size: 0.8rem !important;
            padding: 0.25em 0.8em !important;
            margin-bottom: 0.8rem !important;
          }
          
          /* Table heading adjustments */
          h2[style*='color: #22d3ee'] {
            font-size: 0.9rem !important;
            margin-bottom: 0.8rem !important;
          }
          
          /* Table adjustments */
          table {
            font-size: 0.75rem !important;
            border-radius: 6px !important;
            margin: 0 0.4rem !important;
          }
          
          th, td {
            padding: 0.5rem 0.6rem !important;
            font-size: 0.75rem !important;
          }
          
          /* Download button adjustments */
          button[style*='background: linear-gradient(90deg, #22d3ee'] {
            padding: 0.5rem 1rem !important;
            font-size: 0.8rem !important;
            border-radius: 6px !important;
            margin-top: 0.8rem !important;
          }
          
          /* Decorative blobs adjustments */
          svg[style*='position: absolute'] {
            width: 250px !important;
            height: 250px !important;
          }
        }
        
        @media (max-width: 360px) {
          /* Header adjustments */
          header {
            padding: 0.6rem 0 0.3rem 0 !important;
            margin-bottom: 1.2rem !important;
          }
          
          /* Back button adjustments */
          header button {
            top: 10px !important;
            left: 10px !important;
            width: 28px !important;
            height: 28px !important;
          }
          
          /* Heading adjustments */
          .teacher-dashboard-heading {
            font-size: 0.8rem !important;
            margin: 0 2rem !important;
          }
          
          .teacher-dashboard-subheading {
            font-size: 0.7rem !important;
            margin: 0.3rem 1.5rem 0.2rem 1.5rem !important;
          }
          
          /* Main content adjustments */
          main {
            padding: 1rem 0.6rem 1rem 2rem !important;
            gap: 1.2rem !important;
            margin-right: 0.6rem !important;
            border-right: 2px solid rgba(34,211,238,0.2) !important;
          }
          
          /* Info banner adjustments */
          div[style*='background: rgba(34,211,238,0.10)'] {
            padding: 0.3rem 0.6rem !important;
            font-size: 0.6rem !important;
            margin-bottom: 1rem !important;
          }
          
          /* Main card adjustments */
          div[style*='background: rgba(35,39,47,0.82)'] {
            padding: 1rem 0.6rem !important;
            border-radius: 10px !important;
            min-height: 100px !important;
            margin: 0 0.6rem !important;
          }
          
          /* Welcome message adjustments */
          div[style*='font-size: 1.1rem'][style*='color: #38bdf8'] {
            font-size: 0.75rem !important;
            padding: 0.2em 0.6em !important;
            margin-bottom: 0.6rem !important;
          }
          
          /* Table heading adjustments */
          h2[style*='color: #22d3ee'] {
            font-size: 0.8rem !important;
            margin-bottom: 0.6rem !important;
          }
          
          /* Table adjustments */
          table {
            font-size: 0.7rem !important;
            border-radius: 4px !important;
            margin: 0 0.3rem !important;
          }
          
          th, td {
            padding: 0.4rem 0.5rem !important;
            font-size: 0.7rem !important;
          }
          
          /* Download button adjustments */
          button[style*='background: linear-gradient(90deg, #22d3ee'] {
            padding: 0.4rem 0.8rem !important;
            font-size: 0.75rem !important;
            border-radius: 4px !important;
            margin-top: 0.6rem !important;
          }
          
          /* Decorative blobs adjustments */
          svg[style*='position: absolute'] {
            width: 200px !important;
            height: 200px !important;
          }
        }
        
        /* Ensure table is scrollable on mobile */
        section[style*='max-width: 800px'] > div[style*='background'] {
          width: 100%;
          min-width: 0;
          overflow-x: auto;
        }
        
        table {
          width: 100%;
          min-width: 300px;
        }
        
        /* Ensure right margin and border are visible on mobile */
        @media (max-width: 768px) {
          main {
            box-sizing: border-box !important;
            width: calc(100% - 2rem) !important;
            margin-left: 1rem !important;
            margin-right: 1rem !important;
          }
        }
        
        @media (max-width: 480px) {
          main {
            width: calc(100% - 1.6rem) !important;
            margin-left: 0.8rem !important;
            margin-right: 0.8rem !important;
          }
        }
        
        @media (max-width: 360px) {
          main {
            width: calc(100% - 1.2rem) !important;
            margin-left: 0.6rem !important;
            margin-right: 0.6rem !important;
          }
        }
        
        /* Touch-friendly button interactions */
        @media (max-width: 768px) {
          button {
            min-height: 44px !important;
            min-width: 44px !important;
          }
          
          button:active {
            transform: scale(0.95) !important;
          }
        }
      `}</style>
    </div>
  );
}

export default TeacherDashboard; 