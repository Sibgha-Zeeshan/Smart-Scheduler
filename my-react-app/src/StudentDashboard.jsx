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

function StudentDashboard({ onBack }) {
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
    }}>
      {/* Decorative Blobs */}
      <div style={{
        position: 'absolute',
        top: '8%',
        left: '8%',
        width: '320px',
        height: '320px',
        background: 'radial-gradient(circle, rgba(56,189,248,0.13) 0%, transparent 70%)',
        borderRadius: '50%',
        filter: 'blur(80px)',
        opacity: 0.13,
        pointerEvents: 'none',
        zIndex: 0,
      }} />
      <div style={{
        position: 'absolute',
        bottom: '10%',
        right: '10%',
        width: '380px',
        height: '380px',
        background: 'radial-gradient(circle, rgba(59,130,246,0.13) 0%, transparent 70%)',
        borderRadius: '50%',
        filter: 'blur(80px)',
        opacity: 0.13,
        pointerEvents: 'none',
        zIndex: 0,
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
        <h1 style={{
          fontSize: "2.8rem",
          fontWeight: 900,
          margin: 0,
          letterSpacing: "-2px",
          textShadow: "0 2px 18px #38bdf855, 0 0 30px #38bdf8cc",
          background: 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)',
          WebkitBackgroundClip: 'text',
          WebkitTextFillColor: 'transparent',
          backgroundClip: 'text',
          animation: "glow 2s ease-in-out infinite alternate"
        }}>Student Dashboard</h1>
        <div style={{ fontSize: '1.35rem', color: '#cbd5e1', marginTop: '1.2rem', fontWeight: 500, letterSpacing: '0.5px', maxWidth: 700, marginLeft: 'auto', marginRight: 'auto', lineHeight: 1.5 }}>
          Access your personalized, conflict-free timetable below. Download it for easy reference and stay ahead in your academic journey!
        </div>
      </header>
      <main style={{
        width: "100%",
        maxWidth: 1100,
        margin: "0 auto",
        padding: "3rem 2rem 2rem 2rem",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        zIndex: 1,
      }}>
        {/* Welcome Message & Timetable Card */}
        <section style={{ width: "100%", maxWidth: 800, margin: "0 auto" }}>
          {/* Info Banner */}
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: '1rem',
            background: 'rgba(56,189,248,0.10)',
            borderRadius: 14,
            padding: '1.1rem 2rem',
            boxShadow: '0 2px 12px rgba(56,189,248,0.07)',
            maxWidth: 600,
            width: '100%',
            margin: '0 auto 2.2rem auto',
          }}>
            <Info size={32} color="#38bdf8" style={{ flexShrink: 0 }} />
            <div style={{ fontWeight: 600, fontSize: '1.18rem', color: '#38bdf8', letterSpacing: '0px' }}>
              Your timetable will appear below when available. Download it for easy access!
            </div>
          </div>
          <div style={{
            background: "rgba(35,39,47,0.92)",
            borderRadius: "22px",
            boxShadow: "0 12px 48px rgba(56,189,248,0.18)",
            padding: "3rem 2.5rem 2.5rem 2.5rem",
            marginBottom: "2.5rem",
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            animation: "fadeInUp 1s",
            border: '2.5px solid rgba(56,189,248,0.18)',
            backdropFilter: 'blur(14px)',
            minHeight: 320,
          }}>
            <h2 style={{ color: "#38bdf8", fontWeight: 800, fontSize: "2rem", marginBottom: "1.6rem", letterSpacing: "-1px", textShadow: '0 2px 12px #38bdf855' }}>Your Conflict-Free Timetable</h2>
            {timetable.length === 0 ? (
              <div style={{ color: '#94a3b8', fontSize: '1.25rem', padding: '3rem 0', textAlign: 'center', fontWeight: 500 }}>
                No timetable available yet.<br />Please check back later.
              </div>
            ) : (
              <>
                <div style={{ width: "100%", overflowX: "auto", marginBottom: "2.2rem" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", background: "rgba(56,189,248,0.03)", borderRadius: "12px", color: "#cbd5e1", fontSize: "1.18rem", boxShadow: '0 2px 8px rgba(56,189,248,0.07)' }}>
                    <thead>
                      <tr style={{ background: "rgba(56,189,248,0.10)" }}>
                        <th style={{ padding: "1.1rem 1.5rem", color: "#38bdf8", fontWeight: 800, textAlign: "left", borderBottom: "2px solid #23272f", fontSize: '1.18rem', letterSpacing: '-0.5px' }}>Day</th>
                        <th style={{ padding: "1.1rem 1.5rem", color: "#38bdf8", fontWeight: 800, textAlign: "left", borderBottom: "2px solid #23272f", fontSize: '1.18rem', letterSpacing: '-0.5px' }}>Classes</th>
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
      <footer style={{
        width: '100%',
        textAlign: 'center',
        color: '#94a3b8',
        fontSize: '1.13rem',
        fontWeight: 500,
        padding: '2.5rem 0 1.2rem 0',
        letterSpacing: '0.2px',
        marginTop: 'auto',
        opacity: 0.85,
      }}>
        "Success is the sum of small efforts, repeated day in and day out."<br />
        <span style={{ color: '#38bdf8', fontWeight: 700 }}>Smart Scheduler</span>
      </footer>
      <style>{`
        @keyframes fadeInUp {
          from { opacity: 0; transform: translateY(30px); }
          to { opacity: 1; transform: translateY(0); }
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
      `}</style>
    </div>
  );
}

export default StudentDashboard; 