import React, { useState } from "react";
import { Calendar, CheckCircle, Sparkles, Users, ShieldCheck, Clock } from "lucide-react";
import LoginForm from "./LoginForm";
import SignupForm from "./SignupForm";
import AdminDashboard from "./AdminDashboard";
import StudentDashboard from "./StudentDashboard";
import TeacherDashboard from "./TeacherDashboard";

const initialUsers = [];

function App() {
  const [showSignup, setShowSignup] = useState(false);
  const [showLogin, setShowLogin] = useState(false);
  const [loading, setLoading] = useState(false);
  const [loginError, setLoginError] = useState("");
  const [userRole, setUserRole] = useState("");
  const [users, setUsers] = useState(initialUsers);
  const [pendingSignups, setPendingSignups] = useState([]);
  const [rejectedUsers, setRejectedUsers] = useState([]);
  const [showSignupSubmitted, setShowSignupSubmitted] = useState(false);
  const [pendingLoginEmail, setPendingLoginEmail] = useState("");
  const [adminLogin, setAdminLogin] = useState(false);
  const [adminEmail, setAdminEmail] = useState("");
  const [signupError, setSignupError] = useState("");

  const handleLogin = ({ email, password, username }) => {
    setLoading(true);
    setLoginError("");
    setTimeout(() => {
      // Admin login by username
      if (username === 'umtadmin' && password === 'admin123') {
        setLoading(false);
        setUserRole('admin');
        setShowLogin(false);
        return;
      }
      // Check for rejected user
      if (rejectedUsers.includes(email)) {
        setLoading(false);
        setLoginError('Your signup request was not accepted by the admin.');
        return;
      }
      // Normal user login by email
      const user = users.find(
        (u) => u.email === email && u.password === password
      );
      if (!user) {
        // Check if user is in pendingSignups
        const pending = pendingSignups.find(
          (u) => u.email === email && u.password === password
        );
        setLoading(false);
        if (pending) {
          setPendingLoginEmail(email);
          setLoginError("");
        } else {
          setLoginError("Invalid email or password, or user not found.");
        }
        return;
      }
      setLoading(false);
      setUserRole(user.role);
      setShowLogin(false);
    }, 1500);
  };

  const handleSignup = (signupData) => {
    // Check for duplicate email or username in users or pendingSignups
    const emailExists = users.some(u => u.email === signupData.email) || pendingSignups.some(u => u.email === signupData.email);
    const usernameExists = users.some(u => u.username === signupData.username) || pendingSignups.some(u => u.username === signupData.username);
    if (emailExists) {
      setSignupError("This email is already registered or pending approval.");
      return;
    }
    if (usernameExists) {
      setSignupError("This username is already taken or pending approval.");
      return;
    }
    setPendingSignups((prev) => [...prev, signupData]);
    setShowSignup(false);
    setShowSignupSubmitted(true);
  };

  const handleApprove = (user) => {
    setUsers((prev) => [...prev, user]);
    setPendingSignups((prev) => prev.filter((u) => u.email !== user.email));
  };

  const handleReject = (user) => {
    setPendingSignups((prev) => prev.filter((u) => u.email !== user.email));
    setRejectedUsers((prev) => [...prev, user.email]);
  };

  // Show dashboard if logged in
  if (userRole === "admin")
    return <AdminDashboard users={users} pendingSignups={pendingSignups} rejectedUsers={rejectedUsers} onApprove={handleApprove} onReject={handleReject} onBack={() => setUserRole("")} />;
  if (userRole === "teacher") return <TeacherDashboard onBack={() => setUserRole("")} />;
  if (userRole === "student") return <StudentDashboard onBack={() => setUserRole("")} />;

  return (
    <>
      {/* Top Navigation Bar */}
      <nav className="topbar glass-navbar" style={{ background: 'linear-gradient(135deg, rgba(26, 26, 46, 0.95) 0%, rgba(22, 33, 62, 0.95) 100%)', backdropFilter: 'blur(20px)', borderBottom: '1px solid rgba(255, 255, 255, 0.1)', animation: 'slideDown 0.8s ease-out' }}>
        <div className="topbar-inner" style={{ justifyContent: 'space-between', paddingLeft: '4rem', paddingRight: '4rem' }}>
          <div className="topbar-logo" style={{ justifyContent: 'flex-start', marginLeft: '-4rem', paddingLeft: 0, gap: '0.2rem', alignItems: 'center' }}>
            <span className="logo-text gradient-text" style={{ marginLeft: 0, background: 'linear-gradient(135deg, #22d3ee 0%, #06b6d4 100%)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent', backgroundClip: 'text', fontWeight: '700', animation: 'glow 2s ease-in-out infinite alternate' }}>Smart Scheduler</span>
          </div>
          <div className="topbar-actions" style={{ justifyContent: 'flex-end' }}>
            <button className="topbar-btn login-btn" style={{ background: 'rgba(255, 255, 255, 0.1)', border: '1px solid rgba(255, 255, 255, 0.2)', color: '#f1f5f9', transition: 'all 0.3s ease', animation: 'bounceIn 0.8s ease-out 0.2s both' }} onClick={() => { setShowLogin(true); setAdminLogin(false); }}>Login</button>
            <button className="topbar-btn signup-btn" style={{ background: 'linear-gradient(135deg, #22d3ee 0%, #06b6d4 100%)', border: 'none', color: 'white', transition: 'all 0.3s ease', animation: 'bounceIn 0.8s ease-out 0.4s both' }} onClick={() => setShowSignup(true)}>Sign Up</button>
            <button className="topbar-btn" style={{ background: 'rgba(34, 211, 238, 0.15)', border: '1px solid #22d3ee', color: '#22d3ee', marginLeft: '1rem', transition: 'all 0.3s ease', animation: 'bounceIn 0.8s ease-out 0.8s both' }} onClick={() => { setShowLogin(true); setAdminLogin(true); setAdminEmail('admin@umt.edu.pk'); }}>Login as Admin</button>
          </div>
        </div>
      </nav>

      {/* Login Modal */}
      {showLogin && (
        <>
          <div style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100vw',
            height: '100vh',
            background: 'rgba(16, 23, 42, 0.85)',
            zIndex: 1000,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            animation: 'fadeInUp 0.5s ease'
          }}>
            <div style={{ position: 'relative' }}>
              <button onClick={() => setShowLogin(false)} style={{ position: 'absolute', top: 18, right: 18, background: 'none', border: 'none', color: '#94a3b8', fontSize: '1.5rem', cursor: 'pointer', zIndex: 2 }}>&times;</button>
              <LoginForm onLogin={handleLogin} loading={loading} error={loginError} onClose={() => setShowLogin(false)} defaultEmail={adminLogin ? adminEmail : ""} adminLogin={adminLogin} zoomed={true} />
            </div>
          </div>
        </>
      )}

      {/* Signup Modal */}
      {showSignup && (
        <>
          <div style={{
            position: 'fixed',
            top: 0,
            left: 0,
            width: '100vw',
            height: '100vh',
            background: 'rgba(16, 23, 42, 0.85)',
            zIndex: 1000,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            animation: 'fadeInUp 0.5s ease'
          }}>
            <SignupForm onClose={() => setShowSignup(false)} onSignup={handleSignup} users={users} pendingSignups={pendingSignups} />
          </div>
          {signupError && (
            <div style={{
              position: 'fixed',
              top: 0,
              left: 0,
              width: '100vw',
              height: '100vh',
              background: 'rgba(16, 23, 42, 0.7)',
              zIndex: 2000,
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
            }}>
              <div style={{
                background: 'linear-gradient(145deg, #1a1a2e 0%, #16213e 100%)',
                borderRadius: '12px',
                padding: '1.1rem 1.5rem',
                boxShadow: '0 8px 24px rgba(0,0,0,0.18)',
                maxWidth: '320px',
                textAlign: 'center',
                color: '#f1f5f9',
                border: '1px solid rgba(255,255,255,0.08)'
              }}>
                <h3 style={{ color: '#ff6b6b', marginBottom: '0.7rem', fontSize: '1.05rem' }}>Signup Error</h3>
                <div style={{ fontSize: '0.97rem', color: '#f1f5f9', marginBottom: '1.1rem' }}>{signupError}</div>
                <button onClick={() => setSignupError("")} style={{ background: 'linear-gradient(135deg, #22d3ee 0%, #06b6d4 100%)', border: 'none', color: 'white', borderRadius: '6px', padding: '0.4rem 1rem', fontWeight: '600', fontSize: '0.93rem', cursor: 'pointer' }}>OK</button>
              </div>
            </div>
          )}
        </>
      )}

      {/* Signup Submitted Modal */}
      {showSignupSubmitted && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          background: 'rgba(16, 23, 42, 0.85)',
          zIndex: 1100,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          animation: 'fadeInUp 0.5s ease'
        }}>
          <div style={{
            background: 'linear-gradient(145deg, #1a1a2e 0%, #16213e 100%)',
            borderRadius: '14px',
            padding: '1.5rem 2.2rem', // wider horizontal padding
            boxShadow: '0 12px 24px rgba(0, 0, 0, 0.18)',
            maxWidth: '340px', // wider
            textAlign: 'center',
            color: '#f1f5f9',
            border: '1px solid rgba(255,255,255,0.08)'
          }}>
            <h2 style={{ color: '#22d3ee', marginBottom: '0.8rem', fontSize: '1.1rem' }}>Signup Request Sent!</h2>
            <p style={{ fontSize: '0.93rem', color: '#94a3b8', marginBottom: '1.1rem' }}>
              Your signup request has been submitted.<br />
              Please wait for admin approval before logging in.
            </p>
            <button onClick={() => setShowSignupSubmitted(false)} style={{ background: 'linear-gradient(135deg, #22d3ee 0%, #06b6d4 100%)', border: 'none', color: 'white', borderRadius: '6px', padding: '0.5rem 1.1rem', fontWeight: '600', fontSize: '0.93rem', cursor: 'pointer' }}>OK</button>
          </div>
        </div>
      )}

      {/* Pending Signup In Process Message */}
      {pendingLoginEmail && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          width: '100vw',
          height: '100vh',
          background: 'rgba(16, 23, 42, 0.85)',
          zIndex: 1200,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          animation: 'fadeInUp 0.5s ease'
        }}>
          <div style={{
            background: 'linear-gradient(145deg, #1a1a2e 0%, #16213e 100%)',
            borderRadius: '24px',
            padding: '3rem 2.5rem',
            boxShadow: '0 25px 50px rgba(0, 0, 0, 0.3)',
            maxWidth: '400px',
            textAlign: 'center',
            color: '#f1f5f9',
            border: '1px solid rgba(255,255,255,0.1)'
          }}>
            <h2 style={{ color: '#06b6d4', marginBottom: '1.5rem' }}>Request In Process</h2>
            <p style={{ fontSize: '1.15rem', color: '#94a3b8', marginBottom: '2rem' }}>
              Your signup request is in process.<br />
              Please wait for admin approval before logging in.
            </p>
            <button onClick={() => setPendingLoginEmail("")} style={{ background: 'linear-gradient(135deg, #22d3ee 0%, #06b6d4 100%)', border: 'none', color: 'white', borderRadius: '8px', padding: '0.8rem 2rem', fontWeight: '600', fontSize: '1.1rem', cursor: 'pointer' }}>OK</button>
          </div>
        </div>
      )}

      <div className="main-bg" style={{ background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #334155 100%)', minHeight: '100vh', zoom: 0.67 }}>
        {/* Decorative Blobs */}
        <div className="bg-blob bg-blob-1" style={{ 
          position: 'absolute', 
          top: '10%', 
          left: '10%', 
          width: '300px', 
          height: '300px', 
          background: 'radial-gradient(circle, rgba(34, 211, 238, 0.1) 0%, transparent 70%)',
          borderRadius: '50%',
          animation: 'float 6s ease-in-out infinite alternate, pulse 4s ease-in-out infinite, blobGlow 8s ease-in-out infinite alternate'
        }} />
        <div className="bg-blob bg-blob-2" style={{ 
          position: 'absolute', 
          bottom: '20%', 
          right: '15%', 
          width: '400px', 
          height: '400px', 
          background: 'radial-gradient(circle, rgba(255, 107, 107, 0.1) 0%, transparent 70%)',
          borderRadius: '50%',
          animation: 'float 8s ease-in-out infinite reverse alternate, pulse 4s ease-in-out infinite 2s, blobGlow 10s ease-in-out infinite alternate'
        }} />
        <div className="bg-blob bg-blob-3" style={{ 
          position: 'absolute', 
          top: '50%', 
          left: '5%', 
          width: '200px', 
          height: '200px', 
          background: 'radial-gradient(circle, rgba(78, 205, 196, 0.1) 0%, transparent 70%)',
          borderRadius: '50%',
          animation: 'float 10s ease-in-out infinite alternate, rotate 20s linear infinite, blobGlow 12s ease-in-out infinite alternate'
        }} />

        {/* Hero Section - text and login form side by side */}
        <section className="hero-flex-row" style={{ width: '100%', maxWidth: 1600, margin: '5rem auto 4rem auto', display: 'flex', flexDirection: 'row', alignItems: 'flex-start', justifyContent: 'center', gap: '8rem', background: 'none', boxShadow: 'none', borderRadius: 0, padding: 0, position: 'relative', zIndex: 1 }}>
          <div className="hero-text-col" style={{ flex: 1, minWidth: 260, maxWidth: 520, textAlign: 'left', display: 'flex', flexDirection: 'column', alignItems: 'flex-start', justifyContent: 'center', padding: '4rem', animation: 'fadeInUpHero 1.1s cubic-bezier(.39,.575,.56,1.000) 0.1s both' }}>
            <h1 className="hero-title gradient-text" style={{ marginBottom: '1.5rem', textAlign: 'left', fontSize: '3.5rem', background: 'linear-gradient(135deg, #22d3ee 0%, #06b6d4 100%)', WebkitBackgroundClip: 'text', WebkitTextFillColor: 'transparent', backgroundClip: 'text', fontWeight: '700', letterSpacing: '-1px', lineHeight: '1.2', animation: 'fadeInUpHeroTitle 1.2s cubic-bezier(.39,.575,.56,1.000) 0.2s both, glow 3s ease-in-out infinite alternate' }}>Smart Scheduler</h1>
            <p style={{ fontSize: '1.5rem', color: '#f1f5f9', lineHeight: '1.6', fontWeight: 400, marginBottom: '2.5rem', maxWidth: 600, animation: 'fadeInUpHeroDesc 1.2s cubic-bezier(.39,.575,.56,1.000) 0.5s both' }}>
              Welcome to Smart Scheduler, your intelligent assistant for building a conflict-free timetable. Effortlessly organize your classes, meetings, and personal events in one place. Our platform is designed to simplify your scheduling experience, helping you stay focused on what matters most. Join us and take the stress out of planning!
            </p>
          </div>
          <div className="hero-login-form" style={{ flex: 1, minWidth: 400, maxWidth: 500, minHeight: 600, height: '600px', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '4rem', animation: 'scaleInLogin 1.1s cubic-bezier(.39,.575,.56,1.000) 0.7s both' }}>
            <LoginForm 
              onLogin={handleLogin}
              loading={loading}
              error={loginError}
              defaultEmail={adminLogin ? adminEmail : ""}
              adminLogin={adminLogin}
            />
          </div>
        </section>

        {/* How It Works Section - modern, minimal, visually appealing */}
        <section style={{ width: '100%', maxWidth: 900, margin: '0 auto 3.5rem auto', padding: '0 1.5rem', textAlign: 'center' }}>
          <h2 style={{ fontSize: '2rem', color: '#22d3ee', fontWeight: 700, marginBottom: '2.2rem', letterSpacing: '-1px' }}>How It Works</h2>
          <div style={{
            display: 'flex',
            flexDirection: 'row',
            justifyContent: 'center',
            alignItems: 'flex-start',
            gap: '0',
            flexWrap: 'wrap',
            maxWidth: 800,
            margin: '0 auto',
            position: 'relative',
          }}>
            <div style={{ flex: 1, minWidth: 180, maxWidth: 240, display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '0 1.5rem' }}>
              <Users size={32} color="#22d3ee" style={{ marginBottom: '0.7rem' }} />
              <div style={{ fontSize: '1.13rem', color: '#1e293b', fontWeight: 700, marginBottom: '0.3rem' }}>Sign up or log in</div>
              <div style={{ fontSize: '1rem', color: '#64748b', fontWeight: 400 }}>Create your account to get started</div>
            </div>
            <div className="howitworks-divider" style={{ width: 40, height: 2, background: 'linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%)', alignSelf: 'center', margin: '0 0.5rem', borderRadius: 2, display: 'none' }} />
            <div style={{ flex: 1, minWidth: 180, maxWidth: 240, display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '0 1.5rem' }}>
              <Calendar size={32} color="#22d3ee" style={{ marginBottom: '0.7rem' }} />
              <div style={{ fontSize: '1.13rem', color: '#1e293b', fontWeight: 700, marginBottom: '0.3rem' }}>Add your events</div>
              <div style={{ fontSize: '1rem', color: '#64748b', fontWeight: 400 }}>Input your classes, meetings, or tasks</div>
            </div>
            <div className="howitworks-divider" style={{ width: 40, height: 2, background: 'linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%)', alignSelf: 'center', margin: '0 0.5rem', borderRadius: 2, display: 'none' }} />
            <div style={{ flex: 1, minWidth: 180, maxWidth: 240, display: 'flex', flexDirection: 'column', alignItems: 'center', padding: '0 1.5rem' }}>
              <CheckCircle size={32} color="#22d3ee" style={{ marginBottom: '0.7rem' }} />
              <div style={{ fontSize: '1.13rem', color: '#1e293b', fontWeight: 700, marginBottom: '0.3rem' }}>Get your schedule</div>
              <div style={{ fontSize: '1rem', color: '#64748b', fontWeight: 400 }}>Enjoy a smart, conflict-free timetable</div>
            </div>
          </div>
          <style>{`
            @media (min-width: 700px) {
              .howitworks-divider { display: block !important; }
            }
            @media (max-width: 699px) {
              .howitworks-divider { display: none !important; }
              section[style*='How It Works'] > div { flex-direction: column !important; gap: 2.5rem !important; }
            }
          `}</style>
        </section>

        {/* About Us Component */}
        <section style={{ width: '100vw', minWidth: '100vw', maxWidth: '100vw', position: 'static', left: 0, right: 0, margin: 0, padding: 0, background: 'linear-gradient(120deg, #16213e 0%, #0f172a 100%)', overflow: 'hidden', zIndex: 2 }}>
          {/* Soft mesh SVG pattern overlay */}
          <svg style={{ position: 'absolute', top: 0, left: 0, width: '100%', minWidth: '100vw', maxWidth: '100vw', height: '100%', zIndex: 1, opacity: 0.18, pointerEvents: 'none', transform: 'scale(1.5)', transformOrigin: 'top left' }} viewBox="0 0 1440 320" fill="none" xmlns="http://www.w3.org/2000/svg" preserveAspectRatio="none">
            <path fill="#22d3ee" fillOpacity="0.18" d="M0,160L60,170.7C120,181,240,203,360,197.3C480,192,600,160,720,133.3C840,107,960,85,1080,101.3C1200,117,1320,171,1380,197.3L1440,224L1440,320L1380,320C1320,320,1200,320,1080,320C960,320,840,320,720,320C600,320,480,320,360,320C240,320,120,320,60,320L0,320Z" />
          </svg>
          <div className="aboutus-fadein" style={{ maxWidth: 900, margin: '3.5rem auto', padding: '1.2rem 2.5rem', display: 'flex', flexDirection: 'column', alignItems: 'center', position: 'relative', zIndex: 2, textAlign: 'center', background: 'rgba(16,23,42,0.92)', borderRadius: '2rem', boxShadow: '0 8px 32px 0 rgba(31, 38, 135, 0.18)' }}>
            <h2 className="aboutus-heading-anim" style={{ fontSize: '2.3rem', fontWeight: 900, marginBottom: '1.2rem', letterSpacing: '-1.5px', lineHeight: 1.15, color: '#22d3ee' }}>About Us</h2>
            <div className="aboutus-mission-anim" style={{ color: '#38bdf8', fontWeight: 600, fontSize: '1.25rem', marginBottom: '1.7rem', fontStyle: 'italic', lineHeight: 1.6, maxWidth: 700 }}>
              Empowering you to make the most of your time.
            </div>
            <p className="aboutus-paragraph-anim" style={{ fontSize: '1.22rem', color: '#f1f5f9', lineHeight: 2, maxWidth: 900, marginBottom: '3.2rem', fontWeight: 400 }}>
              Smart Scheduler is built by a diverse team of educators, technologists, and dreamers who believe in the power of organization and simplicity. Our vision is to create a world where students and teachers can focus on what truly matters—learning, teaching, and personal growth—while we handle the scheduling. We value transparency, innovation, and a user-first approach in everything we do.
            </p>
            <div className="aboutus-divider-anim" style={{ width: '0%', height: 3, background: 'linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%)', borderRadius: 2, margin: '0 auto 2.5rem auto', transition: 'width 1.2s cubic-bezier(.39,.575,.56,1.000)', marginBottom: '2.5rem' }} />
            <div style={{ display: 'flex', flexDirection: 'row', gap: '2.5rem', flexWrap: 'wrap', alignItems: 'stretch', justifyContent: 'center', marginTop: '0', width: '100%' }}>
              <div className="aboutus-feature aboutus-feature-1" style={{ flex: 1, minWidth: 220, maxWidth: 320, display: 'flex', flexDirection: 'column', alignItems: 'center', marginBottom: '2rem', opacity: 0, transform: 'translateY(40px)' }}>
                <span className="aboutus-icon aboutus-icon-pulse" style={{ marginBottom: '1rem' }}><Calendar size={36} color="#22d3ee" /></span>
                <div style={{ fontWeight: 700, color: '#f1f5f9', fontSize: '1.18rem', marginBottom: '0.5rem', textAlign: 'center' }}>Easy Scheduling</div>
                <div style={{ color: '#bae6fd', fontSize: '1.08rem', lineHeight: 1.7, textAlign: 'center' }}>
                  Effortlessly organize your classes, meetings, and events with our intuitive scheduling tools.
            </div>
          </div>
              <div className="aboutus-feature aboutus-feature-2" style={{ flex: 1, minWidth: 220, maxWidth: 320, display: 'flex', flexDirection: 'column', alignItems: 'center', marginBottom: '2rem', opacity: 0, transform: 'translateY(40px)' }}>
                <span className="aboutus-icon aboutus-icon-pulse" style={{ marginBottom: '1rem' }}><Users size={36} color="#22d3ee" /></span>
                <div style={{ fontWeight: 700, color: '#f1f5f9', fontSize: '1.18rem', marginBottom: '0.5rem', textAlign: 'center' }}>User Friendly</div>
                <div style={{ color: '#bae6fd', fontSize: '1.08rem', lineHeight: 1.7, textAlign: 'center' }}>
                  Designed for everyone—students, teachers, and admins. Simple, clean, and easy to use.
                </div>
            </div>
              <div className="aboutus-feature aboutus-feature-3" style={{ flex: 1, minWidth: 220, maxWidth: 320, display: 'flex', flexDirection: 'column', alignItems: 'center', marginBottom: '2rem', opacity: 0, transform: 'translateY(40px)' }}>
                <span className="aboutus-icon aboutus-icon-pulse" style={{ marginBottom: '1rem' }}><CheckCircle size={36} color="#22d3ee" /></span>
                <div style={{ fontWeight: 700, color: '#f1f5f9', fontSize: '1.18rem', marginBottom: '0.5rem', textAlign: 'center' }}>Reliable Support</div>
                <div style={{ color: '#bae6fd', fontSize: '1.08rem', lineHeight: 1.7, textAlign: 'center' }}>
                  Our team is here to help you every step of the way, ensuring a smooth scheduling experience.
            </div>
            </div>
            </div>
          </div>
          <style>{`
            .aboutus-fadein {
              animation: aboutus-fadein 1s cubic-bezier(.39,.575,.56,1.000) 0.2s both;
            }
            @keyframes aboutus-fadein {
              from { opacity: 0; transform: translateY(40px); }
              to { opacity: 1; transform: translateY(0); }
            }
            .aboutus-heading-anim {
              animation: aboutus-heading-scalein 0.8s cubic-bezier(.39,.575,.56,1.000) 0.3s both;
            }
            @keyframes aboutus-heading-scalein {
              from { opacity: 0; transform: scale(0.85); }
              to { opacity: 1; transform: scale(1); }
            }
            .aboutus-mission-anim {
              animation: aboutus-mission-fadeleft 0.8s cubic-bezier(.39,.575,.56,1.000) 0.5s both;
            }
            @keyframes aboutus-mission-fadeleft {
              from { opacity: 0; transform: translateX(-40px); }
              to { opacity: 1; transform: translateX(0); }
            }
            .aboutus-paragraph-anim {
              animation: aboutus-paragraph-faderight 0.8s cubic-bezier(.39,.575,.56,1.000) 0.7s both;
            }
            @keyframes aboutus-paragraph-faderight {
              from { opacity: 0; transform: translateX(40px); }
              to { opacity: 1; transform: translateX(0); }
            }
            .aboutus-divider-anim {
              animation: aboutus-divider-grow 1.2s cubic-bezier(.39,.575,.56,1.000) 1.1s both;
            }
            @keyframes aboutus-divider-grow {
              from { width: 0%; }
              to { width: 80%; }
            }
            .aboutus-feature {
              animation: aboutus-feature-fadein 0.8s cubic-bezier(.39,.575,.56,1.000) both;
              transition: transform 0.25s cubic-bezier(.39,.575,.56,1.000), box-shadow 0.25s;
            }
            .aboutus-feature-1 { animation-delay: 1.2s; }
            .aboutus-feature-2 { animation-delay: 1.4s; }
            .aboutus-feature-3 { animation-delay: 1.6s; }
            .aboutus-feature:hover {
              transform: translateY(-8px) scale(1.04);
              box-shadow: 0 12px 32px 0 rgba(34,211,238,0.10);
            }
            @keyframes aboutus-feature-fadein {
              from { opacity: 0; transform: translateY(40px); }
              to { opacity: 1; transform: translateY(0); }
            }
            .aboutus-icon {
              transition: transform 0.25s cubic-bezier(.39,.575,.56,1.000), filter 0.25s;
            }
            .aboutus-icon:hover {
              transform: scale(1.18) rotate(-6deg);
              filter: drop-shadow(0 2px 8px #22d3ee99);
            }
            .aboutus-icon-pulse {
              animation: aboutus-icon-pulse 1.2s cubic-bezier(.39,.575,.56,1.000) 1.1s both;
            }
            @keyframes aboutus-icon-pulse {
              0% { transform: scale(0.7); opacity: 0; }
              60% { transform: scale(1.15); opacity: 1; }
              100% { transform: scale(1); opacity: 1; }
            }
            @media (max-width: 900px) {
              section[style*='About Us'] > div[style*='display: flex'] { flex-direction: column !important; gap: 0 !important; }
              section[style*='About Us'] > div[style*='display: flex'] > div { margin-bottom: 2rem !important; max-width: 100% !important; }
            }
          `}</style>
        </section>

        {/* Divider */}
        <div style={{ width: '100%', maxWidth: 900, margin: '0 auto', borderTop: '1.5px solid rgba(34,211,238,0.13)', marginBottom: '2.5rem' }} />

        {/* Modern Footer */}
        <footer style={{ textAlign: 'center', padding: '3rem 0 1.5rem 0', color: '#94a3b8', borderTop: '4px solid', borderImage: 'linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%) 1', marginTop: '2rem', fontSize: '1.1rem', background: 'none', position: 'relative' }}>
          <div style={{ marginBottom: '1.2rem', fontWeight: 700, color: '#22d3ee', fontSize: '1.3rem', letterSpacing: '-0.5px', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '0.7rem' }}>
            <Sparkles size={28} color="#22d3ee" style={{ verticalAlign: 'middle' }} />
            Smart Scheduler &mdash; Plan Smarter, Live Better
          </div>
          <div>
          &copy; {new Date().getFullYear()} Smart Scheduler. All rights reserved.
          </div>
        </footer>
      </div>

      <style jsx>{`
        @keyframes slideInLeft {
          from {
            opacity: 0;
            transform: translateX(-50px);
          }
          to {
            opacity: 1;
            transform: translateX(0);
          }
        }
        
        @keyframes slideInRight {
          from {
            opacity: 0;
            transform: translateX(50px);
          }
          to {
            opacity: 1;
            transform: translateX(0);
          }
        }
        
        @keyframes slideDown {
          from {
            opacity: 0;
            transform: translateY(-20px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }
        
        @keyframes fadeInUp {
          from {
            opacity: 0;
            transform: translateY(30px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }
        
        @keyframes slideInUp {
          from {
            opacity: 0;
            transform: translateY(50px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }
        
        @keyframes bounceIn {
          0% {
            opacity: 0;
            transform: scale(0.3);
          }
          50% {
            opacity: 1;
            transform: scale(1.05);
          }
          70% {
            transform: scale(0.9);
          }
          100% {
            opacity: 1;
            transform: scale(1);
          }
        }
        
        @keyframes scaleIn {
          from {
            opacity: 0;
            transform: scale(0.8);
          }
          to {
            opacity: 1;
            transform: scale(1);
          }
        }
        
        @keyframes zoomIn {
          from {
            opacity: 0;
            transform: scale(0.5);
          }
          to {
            opacity: 1;
            transform: scale(1);
          }
        }
        
        @keyframes float {
          0%, 100% {
            transform: translateY(0px);
          }
          50% {
            transform: translateY(-20px);
          }
        }
        
        @keyframes pulse {
          0%, 100% {
            transform: scale(1);
            opacity: 0.8;
          }
          50% {
            transform: scale(1.1);
            opacity: 1;
          }
        }
        
        @keyframes glow {
          0% {
            text-shadow: 0 0 5px rgba(34, 211, 238, 0.5);
          }
          100% {
            text-shadow: 0 0 20px rgba(34, 211, 238, 0.8), 0 0 30px rgba(34, 211, 238, 0.6);
          }
        }
        
        @keyframes bounce {
          0%, 20%, 50%, 80%, 100% {
            transform: translateY(0);
          }
          40% {
            transform: translateY(-10px);
          }
          60% {
            transform: translateY(-5px);
          }
        }
        
        @keyframes spin {
          from {
            transform: rotate(0deg);
          }
          to {
            transform: rotate(360deg);
          }
        }
        
        @keyframes rotate {
          from {
            transform: rotate(0deg);
          }
          to {
            transform: rotate(360deg);
          }
        }
        
        @keyframes wiggle {
          0%, 7% {
            transform: rotateZ(0);
          }
          15% {
            transform: rotateZ(-15deg);
          }
          20% {
            transform: rotateZ(10deg);
          }
          25% {
            transform: rotateZ(-10deg);
          }
          30% {
            transform: rotateZ(6deg);
          }
          35% {
            transform: rotateZ(-4deg);
          }
          40%, 100% {
            transform: rotateZ(0);
          }
        }
        
        @keyframes countUp {
          from {
            opacity: 0;
            transform: translateY(20px);
          }
          to {
            opacity: 1;
            transform: translateY(0);
          }
        }
        
        .feature-card:hover,
        .benefit-card:hover {
          transform: translateY(-5px) scale(1.02);
          box-shadow: 0 20px 40px rgba(0, 0, 0, 0.2);
        }
        
        .cta-btn:hover {
          transform: translateY(-2px) scale(1.05);
          box-shadow: 0 8px 25px rgba(34, 211, 238, 0.4);
        }
        
        .topbar-btn:hover {
          transform: translateY(-1px) scale(1.05);
          box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
        }
        
        .stat-block:hover {
          transform: scale(1.05);
          transition: all 0.3s ease;
        }
        /* Responsive styles for main landing page and navbar */
        @media (max-width: 900px) {
          .hero-flex-row {
            flex-direction: column !important;
            gap: 1.2rem !important;
            margin-top: 1.2rem !important;
            margin-bottom: 1.2rem !important;
            padding: 0 0.2rem !important;
            align-items: center !important;
          }
          .hero-text-col, .hero-login-form {
            max-width: 100% !important;
            min-width: 0 !important;
            padding: 1rem 0.2rem !important;
            height: auto !important;
            align-items: center !important;
            text-align: center !important;
            justify-content: center !important;
          }
          .hero-title.gradient-text {
            text-align: center !important;
            font-size: 2rem !important;
            margin-left: auto !important;
            margin-right: auto !important;
          }
          .topbar-inner {
            flex-direction: column !important;
            align-items: center !important;
            gap: 0.5rem !important;
            padding-left: 0.2rem !important;
            padding-right: 0.2rem !important;
            height: auto !important;
          }
          .topbar-actions {
            flex-wrap: wrap !important;
            justify-content: center !important;
            gap: 0.5rem !important;
          }
        }
        @media (max-width: 700px) {
          .hero-title.gradient-text {
            font-size: 1.3rem !important;
          }
          .hero-flex-row {
            margin-top: 0.7rem !important;
            margin-bottom: 0.7rem !important;
            gap: 0.7rem !important;
          }
          .hero-text-col, .hero-login-form {
            padding: 0.7rem 0.1rem !important;
          }
          .aboutus-fadein {
            padding: 0.7rem 0.2rem !important;
            margin: 1.2rem auto !important;
          }
          .aboutus-heading-anim {
            font-size: 1.1rem !important;
          }
          .topbar-inner {
            flex-direction: column !important;
            gap: 0.3rem !important;
            height: auto !important;
          }
          .topbar, .glass-navbar {
            height: auto !important;
          }
        }
        @media (max-width: 500px) {
          .hero-title.gradient-text {
            font-size: 0.95rem !important;
          }
          .hero-flex-row {
            gap: 0.2rem !important;
          }
          .aboutus-fadein {
            padding: 0.2rem 0.05rem !important;
          }
          .topbar-inner {
            padding-left: 0.1rem !important;
            padding-right: 0.1rem !important;
            height: auto !important;
          }
          .topbar, .glass-navbar {
            height: auto !important;
          }
        }
        /* Compact font and spacing for all */
        .main-bg, .topbar, .glass-navbar, .topbar-inner, .hero-flex-row, .hero-text-col, .hero-login-form {
          font-size: 0.98rem;
        }
        .topbar-btn, .signup-btn {
          border-radius: 8px !important;
        }
        @media (max-width: 900px) {
          section[style*='About Us'] svg {
            transform: scale(2) !important;
          }
        }
        @media (max-width: 600px) {
          section[style*='About Us'] svg {
            transform: scale(2.5) !important;
          }
        }
        .topbar, .glass-navbar {
          height: 48px !important;
          min-height: 48px !important;
          padding: 0 !important;
        }
        .topbar-inner {
          height: 48px !important;
          min-height: 48px !important;
          padding-left: 7rem !important;
          padding-right: 2.5rem !important;
        }
        .logo-text {
          font-size: 1.15rem !important;
        }
        .topbar-btn, .signup-btn {
          font-size: 0.98rem !important;
          padding: 0.35rem 0.9rem !important;
        }
        @media (max-width: 900px) {
          .topbar, .glass-navbar, .topbar-inner {
            height: 40px !important;
            min-height: 40px !important;
          }
          .topbar-inner {
            padding-left: 4rem !important;
            padding-right: 1.5rem !important;
          }
          .logo-text {
            font-size: 1rem !important;
          }
          .topbar-btn, .signup-btn {
            font-size: 0.93rem !important;
            padding: 0.28rem 0.7rem !important;
          }
        }
        @media (max-width: 600px) {
          .topbar, .glass-navbar, .topbar-inner {
            height: 36px !important;
            min-height: 36px !important;
          }
          .topbar-inner {
            padding-left: 2rem !important;
            padding-right: 0.7rem !important;
          }
          .logo-text {
            font-size: 0.88rem !important;
          }
          .topbar-btn, .signup-btn {
            font-size: 0.88rem !important;
            padding: 0.18rem 0.5rem !important;
          }
        }
        .main-bg {
          animation: gradientBG 18s ease-in-out infinite alternate;
          background-size: 200% 200% !important;
        }
        @keyframes gradientBG {
          0% {
            background-position: 0% 50%;
          }
          100% {
            background-position: 100% 50%;
          }
        }
        @keyframes fadeInUpHero {
          from { opacity: 0; transform: translateY(40px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeInUpHeroTitle {
          from { opacity: 0; transform: translateY(60px) scale(0.95); }
          to { opacity: 1; transform: translateY(0) scale(1); }
        }
        @keyframes fadeInUpHeroDesc {
          from { opacity: 0; transform: translateY(30px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes scaleInLogin {
          from { opacity: 0; transform: scale(0.8); }
          to { opacity: 1; transform: scale(1); }
        }
        @keyframes blobGlow {
          0% { filter: drop-shadow(0 0 0px #22d3ee00); }
          100% { filter: drop-shadow(0 0 32px #22d3ee55); }
        }
      `}</style>
    </>
  );
}

export default App;
