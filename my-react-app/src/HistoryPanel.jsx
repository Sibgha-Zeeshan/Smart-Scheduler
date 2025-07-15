import React, { useState } from "react";
import { CheckCircle, Edit2, Trash2 } from "lucide-react";

function HistoryPanel({ accepted }) {
  // Local state for edit/delete placeholders
  const [editIdx, setEditIdx] = useState(null);
  const [form, setForm] = useState({ username: '', email: '', role: '' });
  const [acceptedList, setAcceptedList] = useState(accepted);

  // Placeholder handlers
  const handleEdit = (idx) => {
    setEditIdx(idx);
    setForm({
      username: acceptedList[idx].username || '',
      email: acceptedList[idx].email || '',
      role: acceptedList[idx].role || '',
    });
  };
  const handleSave = (idx) => {
    setAcceptedList(list => list.map((item, i) => i === idx ? { ...item, ...form } : item));
    setEditIdx(null);
  };
  const handleDelete = (idx) => {
    setAcceptedList(list => list.filter((_, i) => i !== idx));
  };

  return (
    <div style={{
      width: '100%', minHeight: '80vh',
      background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 60%, #22d3ee 100%)',
      color: '#f1f5f9',
      padding: '3rem 0',
      display: 'flex', flexDirection: 'column', alignItems: 'center',
      position: 'relative',
      overflow: 'hidden',
    }}>
      {/* Glassmorphism + gradient border card */}
      <div style={{
        position: 'absolute',
        top: 0, left: 0, width: '100%', height: '100%',
        background: 'rgba(30,41,59,0.65)',
        backdropFilter: 'blur(18px)',
        zIndex: 0,
      }} />
      <h2 style={{ fontSize: '2.3rem', color: '#22d3ee', fontWeight: 900, marginBottom: '2.5rem', letterSpacing: '-1.5px', zIndex: 2, display: 'flex', alignItems: 'center', gap: 12, textShadow: '0 2px 12px #22d3ee33' }}>
        <CheckCircle size={32} color="#38bdf8" style={{ verticalAlign: 'middle' }} />
        Accepted Users
      </h2>
      <div className="history-glass-card" style={{
        width: '95%',
        maxWidth: 1100,
        borderRadius: 24,
        background: 'rgba(30,41,59,0.85)',
        boxShadow: '0 8px 48px #22d3ee33, 0 2px 12px #22d3ee22',
        padding: '2.5rem 1.5rem',
        zIndex: 2,
        overflowX: 'auto',
        border: '2.5px solid',
        borderImage: 'linear-gradient(90deg, #38bdf8 0%, #22d3ee 100%) 1',
        animation: 'fadeInUp 1s cubic-bezier(.39,.575,.56,1.000) both',
      }}>
        <table className="history-table-modern" style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, background: 'none', borderRadius: 18, overflow: 'hidden' }}>
          <thead>
            <tr style={{ color: '#38bdf8', fontWeight: 900, fontSize: '1.18rem', background: 'rgba(34,211,238,0.13)' }}>
              <th style={{ padding: '1.1rem 0.7rem', textAlign: 'left', borderBottom: '2.5px solid #38bdf855', letterSpacing: '-0.5px' }}>Username</th>
              <th style={{ padding: '1.1rem 0.7rem', textAlign: 'left', borderBottom: '2.5px solid #38bdf855', letterSpacing: '-0.5px' }}>Email</th>
              <th style={{ padding: '1.1rem 0.7rem', textAlign: 'left', borderBottom: '2.5px solid #38bdf855', letterSpacing: '-0.5px' }}>Role</th>
              <th style={{ padding: '1.1rem 0.7rem', textAlign: 'center', borderBottom: '2.5px solid #38bdf855', letterSpacing: '-0.5px' }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {acceptedList.length === 0 ? (
              <tr><td colSpan={4} style={{ color: '#64748b', fontSize: '1.1rem', textAlign: 'center', padding: '2rem 0' }}>No accepted users found.</td></tr>
            ) : acceptedList.map((item, idx) => (
              <tr key={idx} style={{
                borderBottom: '1.5px solid #334155',
                background: idx % 2 === 0 ? 'rgba(34,211,238,0.04)' : 'rgba(30,41,59,0.12)',
                transition: 'background 0.2s',
                borderRadius: 12,
                boxShadow: editIdx === idx ? '0 2px 12px #22d3ee33' : 'none',
              }}>
                {editIdx === idx ? (
                  <>
                    <td><input value={form.username} onChange={e => setForm(f => ({ ...f, username: e.target.value }))} className="history-input-modern" style={inputStyle} /></td>
                    <td><input value={form.email} onChange={e => setForm(f => ({ ...f, email: e.target.value }))} className="history-input-modern" style={inputStyle} /></td>
                    <td><input value={form.role} onChange={e => setForm(f => ({ ...f, role: e.target.value }))} className="history-input-modern" style={inputStyle} /></td>
                    <td style={{ textAlign: 'center' }}>
                      <button className="history-btn-modern" style={btnStyle} onClick={() => handleSave(idx)}>Save</button>
                      <button className="history-btn-modern" style={btnStyle} onClick={() => setEditIdx(null)}>Cancel</button>
                    </td>
                  </>
                ) : (
                  <>
                    <td style={{ fontWeight: 700, fontSize: '1.07rem', letterSpacing: '-0.5px' }}>{item.username || <span style={{ color: '#64748b' }}>N/A</span>}</td>
                    <td style={{ color: '#a5b4fc', fontWeight: 500 }}>{item.email}</td>
                    <td style={{ color: '#38bdf8', fontWeight: 600 }}>{item.role}
                      <span style={{ marginLeft: 10, verticalAlign: 'middle', display: 'inline-block' }}>
                        <span className="status-badge" style={{
                          background: 'linear-gradient(90deg, #38bdf8 0%, #22d3ee 100%)',
                          color: '#fff',
                          borderRadius: 8,
                          padding: '0.18em 0.85em',
                          fontSize: '0.93em',
                          fontWeight: 700,
                          marginLeft: 2,
                          letterSpacing: '0.5px',
                          boxShadow: '0 2px 8px #22d3ee33',
                        }}>Accepted</span>
                      </span>
                    </td>
                    <td style={{ textAlign: 'center' }}>
                      <button className="history-icon-btn-modern" style={iconBtnStyle} onClick={() => handleEdit(idx)} title="Edit"><Edit2 size={20} color="#38bdf8" /></button>
                      <button className="history-icon-btn-modern" style={iconBtnStyle} onClick={() => handleDelete(idx)} title="Delete"><Trash2 size={20} color="#f43f5e" /></button>
                    </td>
                  </>
                )}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <style>{`
        .history-glass-card {
          animation: fadeInUp 1s cubic-bezier(.39,.575,.56,1.000) both;
        }
        .history-table-modern th, .history-table-modern td {
          border-radius: 12px;
        }
        .history-table-modern tr {
          transition: background 0.18s;
        }
        .history-table-modern tr:hover {
          background: rgba(56,189,248,0.09) !important;
        }
        .history-input-modern {
          background: #18181b;
          color: #f1f5f9;
          border: 1.5px solid #334155;
          border-radius: 7px;
          padding: 0.5rem 0.9rem;
          font-size: 1rem;
          outline: none;
          transition: border 0.2s, box-shadow 0.2s;
        }
        .history-input-modern:focus {
          border: 1.5px solid #22d3ee;
          box-shadow: 0 2px 8px #22d3ee33;
        }
        .history-btn-modern {
          background: linear-gradient(90deg, #22d3ee 0%, #06b6d4 100%);
          color: #fff;
          border: none;
          border-radius: 7px;
          padding: 0.45rem 1.1rem;
          font-weight: 700;
          font-size: 1rem;
          cursor: pointer;
          margin-left: 0.2rem;
          transition: background 0.2s, box-shadow 0.2s, transform 0.2s;
        }
        .history-btn-modern:hover {
          background: linear-gradient(90deg, #38bdf8 0%, #22d3ee 100%);
          box-shadow: 0 4px 16px #22d3ee33;
          transform: translateY(-2px) scale(1.04);
        }
        .history-icon-btn-modern {
          background: none;
          border: none;
          cursor: pointer;
          margin: 0 0.2rem;
          padding: 0.2rem;
          border-radius: 5px;
          transition: background 0.18s;
        }
        .history-icon-btn-modern:hover {
          background: #1e293b;
        }
        @keyframes fadeInUp {
          from { opacity: 0; transform: translateY(40px); }
          to { opacity: 1; transform: translateY(0); }
        }
      `}</style>
    </div>
  );
}

const inputStyle = {
  background: '#18181b',
  color: '#f1f5f9',
  border: '1.5px solid #334155',
  borderRadius: 7,
  padding: '0.5rem 0.9rem',
  fontSize: '1rem',
  outline: 'none',
  width: '100%',
};

const btnStyle = {
  marginLeft: 4,
  marginRight: 4,
  fontWeight: 700,
  fontSize: '1rem',
  borderRadius: 7,
  padding: '0.45rem 1.1rem',
  border: 'none',
  cursor: 'pointer',
};

const iconBtnStyle = {
  background: 'none',
  border: 'none',
  cursor: 'pointer',
  margin: '0 0.2rem',
  padding: '0.2rem',
  borderRadius: 5,
};

export default HistoryPanel; 