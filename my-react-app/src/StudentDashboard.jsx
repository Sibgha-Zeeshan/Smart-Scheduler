import React, { useState, useEffect } from "react";
import { Calendar, Info } from "lucide-react";

const sampleTimetable = [
  { day: "Monday", classes: ["Mathematics", "Physics"] },
  { day: "Tuesday", classes: ["English", "Biology"] },
  { day: "Wednesday", classes: ["Chemistry", "Mathematics"] },
  { day: "Thursday", classes: ["Physics", "English"] },
  { day: "Friday", classes: ["Biology", "Chemistry"] },
];

// Helper function to download CSV
function downloadCSV(data, filename) {
  const csvRows = [
    ["Day", "Classes"],
    ...data.map((row) => [row.day, row.classes.join(", ")]),
  ];
  const csvContent = csvRows.map((e) => e.join(",")).join("\n");
  const blob = new Blob([csvContent], { type: "text/csv" });
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.setAttribute("hidden", "");
  a.setAttribute("href", url);
  a.setAttribute("download", filename);
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

function StudentDashboard({ onBack, username }) {
  // Start with no timetable; will be set by API in the future
  const [timetable, setTimetable] = useState([]);

  // In the future, fetch from API here
  // useEffect(() => {
  //   fetch('/api/student-timetable')
  //     .then(res => res.json())
  //     .then(data => setTimetable(data));
  // }, []);

  return (
    <div style={{
      minHeight: "100vh",
      width: "100vw",
      background: "linear-gradient(135deg, #0f172a 0%, #1e293b 100%)",
      display: "flex",
      flexDirection: "column",
      alignItems: "center",
      justifyContent: "flex-start",
      padding: 0,
      position: "relative",
      overflow: "hidden",
      zoom: 1,
    }}>
      {/* Decorative Animated Blobs for Visual Appeal */}
      <div className="student-dashboard-blob1" style={{
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
        animation: 'blobMove 7s ease-in-out infinite alternate',
      }} />
      <div className="student-dashboard-blob2" style={{
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
        animation: 'blobMove 9s ease-in-out infinite alternate-reverse',
      }} />
      <div className="student-dashboard-blob3" style={{
        position: 'absolute',
        top: '60%',
        left: '50%',
        width: '180px',
        height: '180px',
        background: 'radial-gradient(circle, rgba(52,211,153,0.10) 0%, transparent 70%)',
        borderRadius: '50%',
        filter: 'blur(60px)',
        opacity: 0.10,
        pointerEvents: 'none',
        zIndex: 0,
        animation: 'blobMove 11s ease-in-out infinite alternate',
      }} />
      {/* Topbar with Back Icon Button */}
      <header style={{
        width: "100%",
        padding: "3.5rem 0 2rem 0",
        textAlign: "center",
        background: "rgba(35,39,47,0.97)",
        color: "#fff",
        boxShadow: "0 2px 24px rgba(30,41,59,0.13)",
        marginBottom: "2.5rem",
        position: "relative",
        zIndex: 1,
        animation: "slideDown 0.8s ease-out",
        backdropFilter: 'blur(16px)',
      }}>
        {onBack && (
          <button
            onClick={onBack}
            style={{
              position: "absolute",
              top: 36,
              left: 48,
              background: "rgba(56,189,248,0.13)",
              border: "none",
              borderRadius: "50%",
              width: 56,
              height: 56,
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
              cursor: "pointer",
              boxShadow: "0 2px 12px rgba(56,189,248,0.10)",
              transition: "background 0.18s, transform 0.18s",
              zIndex: 2,
              outline: "none",
              animation: "fadeInUp 1s ease-out 0.2s both",
            }}
            className="student-back-btn"
            aria-label="Back"
          >
            <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ display: "block", transition: "stroke 0.2s" }}>
              <path d="M15 18l-6-6 6-6" />
            </svg>
          </button>
        )}
        {username && (
          <div style={{
            fontSize: '1.13rem',
            color: '#38bdf8',
            fontWeight: 700,
            marginBottom: '0.7rem',
            textAlign: 'center',
            letterSpacing: '0.2px',
            animation: 'bounceIn 1.1s',
            position: 'relative',
            display: 'inline-block',
          }}>
            Welcome, {username}!
            <span style={{
              display: 'block',
              width: 44,
              height: 3,
              margin: '0.3rem auto 0 auto',
              borderRadius: 2,
              background: 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)',
              opacity: 0.6,
              animation: 'pulseBar 2.2s infinite',
            }} />
          </div>
        )}
        <h1
          className="student-dashboard-heading"
          style={{
            fontSize: '1.7rem',
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
            lineHeight: 1.1,
            textShadow: '0 2px 12px #38bdf855, 0 0 18px #38bdf899',
          }}
        >
          Student Dashboard
          <span
            style={{
              display: 'block',
              width: 60,
              height: 5,
              margin: '0.5rem auto 0 auto',
              borderRadius: 4,
              background: 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)',
              opacity: 0.7,
              animation: 'pulseBar 2.5s infinite',
            }}
          />
        </h1>
        <div
          className="student-dashboard-subheading"
          style={{
            display: 'inline-block',
            fontSize: '0.98rem',
            color: '#38bdf8',
            background: 'rgba(56,189,248,0.10)',
            borderRadius: 9999,
            padding: '0.35em 1.2em',
            fontWeight: 500,
            margin: '0.7rem auto 0.5rem auto',
            textAlign: 'center',
            boxShadow: '0 2px 8px rgba(56,189,248,0.07)',
            animation: 'fadeInUp 1.4s',
            lineHeight: 1.3,
            border: '1.5px solid #38bdf822',
            backdropFilter: 'blur(2px)',
          }}
        >
          Access your personalized, conflict-free timetable below. Download it for easy reference and stay ahead in your academic journey!
        </div>
      </header>
      {/* Main Glassmorphism Content Wrapper */}
      <div className="student-dashboard-main-card" style={{
        width: '100%',
        maxWidth: 550,
        margin: '3.5rem auto',
        background: 'rgba(30,41,59,0.72)',
        borderRadius: 28,
        boxShadow: '0 8px 48px 0 rgba(34,211,238,0.10)',
        border: '2.5px solid rgba(34,211,238,0.10)',
        backdropFilter: 'blur(12px)',
        padding: '1.2rem 5rem 1.2rem 2rem',
        position: 'relative',
        zIndex: 1,
        animation: 'fadeInUp 1.2s',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
      }}>
        <main style={{
          width: "100%",
          margin: "0 auto",
          padding: "3rem 2rem 2rem 2rem",
          display: "flex",
          flexDirection: "column",
          alignItems: "center", // Center children horizontally
          zIndex: 1,
          gap: '2.5rem', // Add gap between sections
        }}>
          {/* Welcome Message & Timetable Card */}
          <section style={{ width: "100%", margin: 0, display: 'flex', flexDirection: 'column', gap: '2rem', alignItems: 'center' }}>
            {/* Info Banner */}
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '0.3rem',
              background: 'rgba(56,189,248,0.10)',
              borderRadius: 12,
              padding: '0.2rem 0.5rem',
              boxShadow: '0 2px 8px rgba(56,189,248,0.07)',
              width: '100%',
              margin: '0 auto 1.1rem auto',
              fontSize: '0.75rem',
              fontWeight: 500,
              textAlign: 'center',
              justifyContent: 'center',
              animation: 'fadeInUp 1s',
            }}>
              <Info size={32} color="#38bdf8" style={{ flexShrink: 0 }} />
              <div style={{ fontWeight: 600, fontSize: '0.78rem', color: '#38bdf8', letterSpacing: '0px' }}>
                Your timetable will appear below when available. Download it for easy access!
              </div>
            </div>
            <div style={{
              background: "rgba(35,39,47,0.82)",
              borderRadius: "22px",
              boxShadow: "0 8px 32px rgba(56,189,248,0.13)",
              padding: "1.2rem 1.1rem 1.1rem 1.1rem",
              margin: 0,
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
              animation: "fadeInUp 1s",
              border: '2.5px solid',
              borderImage: 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%) 1',
              backdropFilter: 'blur(18px)',
              minHeight: 180,
              width: '100%',
            }}>
              <h2 style={{ color: "#38bdf8", fontWeight: 700, fontSize: "1.1rem", marginBottom: "0.7rem", letterSpacing: "-0.5px", textShadow: '0 1px 4px #38bdf855', width: '100%', textAlign: 'center' }}>Your Conflict-Free Timetable</h2>
              {timetable.length === 0 ? (
                <div style={{ color: '#94a3b8', fontSize: '0.92rem', padding: '1.2rem 0', textAlign: 'center', fontWeight: 400, width: '100%' }}>
                  No timetable available yet.<br />Please check back later.
                </div>
              ) : (
                <>
                  <div className="student-dashboard-table-wrapper" style={{ width: "100%", overflowX: "auto", marginBottom: "2.2rem", display: 'flex', justifyContent: 'center' }}>
                    <table style={{ minWidth: 340, width: "80%", borderCollapse: "collapse", background: "rgba(56,189,248,0.03)", borderRadius: "12px", color: "#cbd5e1", fontSize: "1.18rem", boxShadow: '0 2px 8px rgba(56,189,248,0.07)', textAlign: 'center', margin: '0 auto' }}>
                      <thead>
                        <tr style={{ background: "rgba(56,189,248,0.10)" }}>
                          <th style={{ padding: "1.1rem 1.5rem", color: "#38bdf8", fontWeight: 800, textAlign: "center", borderBottom: "2px solid #23272f", fontSize: '1.18rem', letterSpacing: '-0.5px' }}>Day</th>
                          <th style={{ padding: "1.1rem 1.5rem", color: "#38bdf8", fontWeight: 800, textAlign: "center", borderBottom: "2px solid #23272f", fontSize: '1.18rem', letterSpacing: '-0.5px' }}>Classes</th>
                        </tr>
                      </thead>
                      <tbody>
                        {timetable.map((row, i) => (
                          <tr key={row.day} style={{ borderBottom: "1px solid #23272f" }}>
                            <td style={{ padding: "1.1rem 1.5rem", fontWeight: 700, textAlign: 'center' }}>{row.day}</td>
                            <td style={{ padding: "1.1rem 1.5rem", textAlign: 'center' }}>{row.classes.join(", ")}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                  <button
                    className="student-dashboard-download-btn"
                    onClick={() => downloadCSV(timetable, "student-timetable.csv")}
                    style={{
                      background: "linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)",
                      color: "#fff",
                      border: "none",
                      borderRadius: "14px",
                      padding: "1.2rem 2.8rem",
                      fontWeight: 800,
                      fontSize: "1.18rem",
                      cursor: "pointer",
                      opacity: 1,
                      boxShadow: "0 6px 24px rgba(56,189,248,0.18)",
                      marginTop: "0.5rem",
                      letterSpacing: "0.7px",
                      transition: "all 0.2s",
                      outline: 'none',
                      alignSelf: 'center',
                      minWidth: 220,
                    }}
                    onMouseOver={e => e.currentTarget.style.background = 'linear-gradient(90deg, #0ea5e9 0%, #38bdf8 100%)'}
                    onMouseOut={e => e.currentTarget.style.background = 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)'}
                    onFocus={e => e.currentTarget.style.boxShadow = '0 0 0 4px #38bdf8aa'}
                    onBlur={e => e.currentTarget.style.boxShadow = '0 6px 24px rgba(56,189,248,0.18)'}
                  >
                    Download Timetable
                  </button>
                </>
              )}
            </div>
          </section>
        </main>
      </div> {/* End Glassmorphism Content Wrapper */}
      <footer style={{
        width: '100%',
        textAlign: 'center',
        color: '#94a3b8',
        fontSize: '0.85rem',
        fontWeight: 400,
        padding: '1.1rem 0 0.7rem 0',
        letterSpacing: '0.1px',
        marginTop: 'auto',
        opacity: 0.7,
        zIndex: 2,
      }}>
        "Success is the sum of small efforts, repeated day in and day out."<br />
        <span style={{ color: '#38bdf8', fontWeight: 700 }}>Smart Scheduler</span>
      </footer>
      <style>{`
        @keyframes fadeInUp {
          from { opacity: 0; transform: translateY(30px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes bounceIn {
          0% { opacity: 0; transform: scale(0.7) translateY(-40px); }
          60% { opacity: 1; transform: scale(1.05) translateY(10px); }
          80% { transform: scale(0.98) translateY(-2px); }
          100% { opacity: 1; transform: scale(1) translateY(0); }
        }
        @keyframes slideInLeft {
          0% { opacity: 0; transform: translateX(-60px); }
          100% { opacity: 1; transform: translateX(0); }
        }
        @keyframes pulseBar {
          0% { opacity: 0.5; }
          50% { opacity: 1; }
          100% { opacity: 0.5; }
        }
        @keyframes blobMove {
          0% { transform: scale(1) translateY(0); }
          100% { transform: scale(1.08) translateY(30px); }
        }
        @keyframes glow {
          0% { text-shadow: 0 0 5px #38bdf855; }
          100% { text-shadow: 0 0 20px #38bdf8cc, 0 0 30px #38bdf899; }
        }
        .student-back-btn:hover {
          background: #38bdf8;
          transform: scale(1.08) translateX(-2px);
        }
        .student-back-btn:hover svg {
          stroke: #fff;
        }
        @media (max-width: 900px) {
          .student-dashboard-main-card {
            max-width: 98vw !important;
            padding: 1.2rem 0.7rem 1.2rem 0.7rem !important;
          }
          .student-dashboard-heading {
            font-size: 1.2rem !important;
          }
          .student-dashboard-subheading {
            font-size: 0.85rem !important;
            padding: 0.25em 0.7em !important;
          }
          .student-dashboard-table-wrapper {
            padding: 0 0.2rem !important;
          }
        }
        @media (max-width: 600px) {
          .student-dashboard-main-card {
            max-width: 100vw !important;
            padding: 0.5rem 0.1rem 0.5rem 0.1rem !important;
            border-radius: 10px !important;
          }
          .student-dashboard-heading {
            font-size: 0.98rem !important;
          }
          .student-dashboard-subheading {
            font-size: 0.7rem !important;
            padding: 0.18em 0.4em !important;
          }
          .student-back-btn {
            left: 8px !important;
            top: 8px !important;
            width: 32px !important;
            height: 32px !important;
          }
          .student-dashboard-blob1 {
            width: 80px !important;
            height: 80px !important;
            top: 2% !important;
            left: 2% !important;
          }
          .student-dashboard-blob2 {
            width: 100px !important;
            height: 100px !important;
            bottom: 2% !important;
            right: 2% !important;
          }
          .student-dashboard-blob3 {
            width: 50px !important;
            height: 50px !important;
            top: 75% !important;
            left: 60% !important;
          }
          .student-dashboard-table-wrapper {
            padding: 0 0.05rem !important;
          }
          .student-dashboard-download-btn {
            width: 100% !important;
            font-size: 1rem !important;
            padding: 1rem 0 !important;
            min-width: 0 !important;
          }
        }
      `}</style>
    </div>
  );
}

export default StudentDashboard; 