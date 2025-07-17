import React, { useState } from "react";
import { Upload, FileSpreadsheet, AlertCircle, CheckCircle, Clock, Edit3 } from "lucide-react";

function TimetableUpload({ onBack }) {
  const [selectedFile, setSelectedFile] = useState(null);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadStatus, setUploadStatus] = useState(null); // 'success', 'error', null
  const [uploadMessage, setUploadMessage] = useState("");

  const handleFileSelect = (event) => {
    const file = event.target.files[0];
    if (file) {
      // Check if file is Excel
      const validTypes = [
        'application/vnd.ms-excel',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel.sheet.macroEnabled.12',
        'application/vnd.ms-excel.template.macroEnabled.12'
      ];
      
      if (validTypes.includes(file.type) || file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        setSelectedFile(file);
        setUploadStatus(null);
        setUploadMessage("");
      } else {
        setUploadStatus('error');
        setUploadMessage("Please select a valid Excel file (.xlsx or .xls)");
        setSelectedFile(null);
      }
    }
  };

  const handleUpload = async () => {
    if (!selectedFile) {
      setUploadStatus('error');
      setUploadMessage("Please select a file first");
      return;
    }

    setIsUploading(true);
    setUploadStatus(null);
    setUploadMessage("");

    try {
      const formData = new FormData();
      formData.append('timetable_file', selectedFile);

      // API call to upload Excel file
      const response = await fetch('/api/admin/upload-timetable', {
        method: 'POST',
        body: formData,
        headers: {
          // Don't set Content-Type header, let browser set it with boundary
        },
      });

      if (response.ok) {
        const result = await response.json();
        setUploadStatus('success');
        setUploadMessage(result.message || "Conflict-free timetable generated successfully! Your optimized schedule is ready.");
        setSelectedFile(null);
        // Reset file input
        const fileInput = document.getElementById('file-input');
        if (fileInput) fileInput.value = '';
      } else {
        const errorData = await response.json();
        setUploadStatus('error');
        setUploadMessage(errorData.message || "Failed to generate conflict-free timetable. Please check your file format and try again.");
      }
    } catch (error) {
      console.error('Upload error:', error);
      setUploadStatus('error');
      setUploadMessage("Error generating conflict-free timetable. Please ensure your Excel file contains valid data and try again.");
    } finally {
      setIsUploading(false);
    }
  };

  const handleDragOver = (e) => {
    e.preventDefault();
  };

  const handleDrop = (e) => {
    e.preventDefault();
    const files = e.dataTransfer.files;
    if (files.length > 0) {
      const file = files[0];
      const event = { target: { files: [file] } };
      handleFileSelect(event);
    }
  };

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
      {/* Decorative Blobs */}
      <div style={{
        position: 'absolute',
        top: '8%',
        left: '8%',
        width: 'clamp(160px, 40vw, 260px)',
        height: 'clamp(160px, 40vw, 260px)',
        background: 'radial-gradient(circle, rgba(167, 139, 250, 0.13) 0%, transparent 70%)',
        borderRadius: '50%',
        animation: 'float 7s ease-in-out infinite, pulse 4s ease-in-out infinite',
        zIndex: 0,
      }} />
      <div style={{
        position: 'absolute',
        bottom: '10%',
        right: '10%',
        width: 'clamp(200px, 50vw, 320px)',
        height: 'clamp(200px, 50vw, 320px)',
        background: 'radial-gradient(circle, rgba(124, 58, 237, 0.13) 0%, transparent 70%)',
        borderRadius: '50%',
        animation: 'float 9s ease-in-out infinite reverse, pulse 4s ease-in-out infinite 2s',
        zIndex: 0,
      }} />

      {/* Header */}
      <header style={{
        width: '100%',
        padding: '1.2rem 1rem 1.1rem 1rem',
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
        {/* Back Button */}
        {onBack && (
          <button
            onClick={onBack}
            style={{
              position: 'absolute',
              top: '1.5rem',
              left: '1rem',
              background: 'rgba(56,189,248,0.10)',
              border: 'none',
              borderRadius: '50%',
              width: '2.75rem',
              height: '2.75rem',
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
            aria-label="Back"
          >
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="#38bdf8" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
              <path d="M15 18l-6-6 6-6" />
            </svg>
          </button>
        )}
        
        <h1 style={{
          fontSize: 'clamp(1rem, 4vw, 1.25rem)',
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
          paddingLeft: onBack ? '3rem' : '0',
          paddingRight: onBack ? '3rem' : '0',
        }}>
          Upload Timetable
        </h1>
        <p style={{ 
          fontSize: 'clamp(0.8rem, 3vw, 0.92rem)', 
          color: '#cbd5e1', 
          marginTop: '0.4rem', 
          fontWeight: 500, 
          letterSpacing: '0.3px', 
          animation: 'fadeInUp 1s ease-out 0.3s both',
          paddingLeft: onBack ? '3rem' : '0',
          paddingRight: onBack ? '3rem' : '0',
        }}>
          Upload your Excel file to generate a conflict-free timetable
        </p>
      </header>

      {/* Main Content */}
      <main style={{
        width: '100%',
        maxWidth: '800px',
        margin: '0 auto',
        padding: 'clamp(1rem, 4vw, 2rem)',
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        zIndex: 1,
        gap: 'clamp(1rem, 4vw, 2rem)',
      }}>
        {/* Upload Area */}
        <div style={{
          width: '100%',
          background: 'rgba(35,39,47,0.92)',
          borderRadius: 'clamp(12px, 3vw, 20px)',
          padding: 'clamp(2rem, 6vw, 3rem) clamp(1rem, 4vw, 2rem)',
          border: '2px dashed rgba(56,189,248,0.3)',
          textAlign: 'center',
          transition: 'all 0.3s ease',
          animation: 'fadeInUp 1s ease-out 0.5s both',
          cursor: 'pointer',
        }}
        onDragOver={handleDragOver}
        onDrop={handleDrop}
        onClick={() => document.getElementById('file-input').click()}
        >
          <FileSpreadsheet size="clamp(48px, 12vw, 64px)" color="#38bdf8" style={{ marginBottom: '1rem' }} />
          <h3 style={{ 
            color: '#38bdf8', 
            fontSize: 'clamp(1.2rem, 4vw, 1.5rem)', 
            fontWeight: 700, 
            marginBottom: '1rem' 
          }}>
            Upload Excel File
          </h3>
          <p style={{ 
            color: '#cbd5e1', 
            fontSize: 'clamp(0.9rem, 3vw, 1rem)', 
            marginBottom: 'clamp(1.5rem, 4vw, 2rem)',
            lineHeight: '1.5'
          }}>
            Drag and drop your Excel file here or click to browse
          </p>
          
          <input
            id="file-input"
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFileSelect}
            style={{ display: 'none' }}
          />
          
          {selectedFile && (
            <div style={{
              background: 'rgba(56,189,248,0.1)',
              borderRadius: 'clamp(8px, 2vw, 12px)',
              padding: 'clamp(0.75rem, 2vw, 1rem)',
              marginTop: '1rem',
              border: '1px solid rgba(56,189,248,0.3)',
            }}>
              <div style={{ 
                display: 'flex', 
                alignItems: 'center', 
                gap: '0.5rem', 
                justifyContent: 'center',
                flexWrap: 'wrap'
              }}>
                <FileSpreadsheet size="clamp(16px, 4vw, 20px)" color="#38bdf8" />
                <span style={{ 
                  color: '#38bdf8', 
                  fontWeight: 600,
                  fontSize: 'clamp(0.8rem, 3vw, 1rem)',
                  wordBreak: 'break-word',
                  textAlign: 'center'
                }}>
                  {selectedFile.name}
                </span>
              </div>
              <p style={{ 
                color: '#cbd5e1', 
                fontSize: 'clamp(0.8rem, 2.5vw, 0.9rem)', 
                marginTop: '0.5rem',
                textAlign: 'center'
              }}>
                Size: {(selectedFile.size / 1024 / 1024).toFixed(2)} MB
              </p>
            </div>
          )}
        </div>

        {/* Upload Status */}
        {uploadStatus && (
          <div style={{
            width: '100%',
            padding: 'clamp(0.75rem, 2vw, 1rem)',
            borderRadius: 'clamp(8px, 2vw, 12px)',
            display: 'flex',
            alignItems: 'center',
            gap: '0.5rem',
            animation: 'fadeInUp 0.5s ease-out',
            background: uploadStatus === 'success' 
              ? 'rgba(34,197,94,0.1)' 
              : 'rgba(239,68,68,0.1)',
            border: `1px solid ${
              uploadStatus === 'success' 
                ? 'rgba(34,197,94,0.3)' 
                : 'rgba(239,68,68,0.3)'
            }`,
            flexWrap: 'wrap',
            justifyContent: 'center',
          }}>
            {uploadStatus === 'success' ? (
              <CheckCircle size="clamp(16px, 4vw, 20px)" color="#22c55e" />
            ) : (
              <AlertCircle size="clamp(16px, 4vw, 20px)" color="#ef4444" />
            )}
            <span style={{ 
              color: uploadStatus === 'success' ? '#22c55e' : '#ef4444',
              fontWeight: 600,
              fontSize: 'clamp(0.8rem, 3vw, 1rem)',
              textAlign: 'center',
              wordBreak: 'break-word'
            }}>
              {uploadMessage}
            </span>
          </div>
        )}

        {/* Upload Buttons */}
        <div style={{
          display: 'flex',
          gap: 'clamp(0.75rem, 2vw, 1rem)',
          flexWrap: 'wrap',
          justifyContent: 'center',
          alignItems: 'center',
          width: '100%',
        }}>
          <button
            onClick={handleUpload}
            disabled={!selectedFile || isUploading}
            style={{
              background: selectedFile && !isUploading 
                ? 'linear-gradient(90deg, #38bdf8 0%, #0ea5e9 100%)' 
                : 'rgba(56,189,248,0.3)',
              color: '#fff',
              border: 'none',
              borderRadius: 'clamp(8px, 2vw, 12px)',
              padding: 'clamp(0.75rem, 3vw, 1rem) clamp(1.5rem, 4vw, 2rem)',
              fontWeight: 700,
              fontSize: 'clamp(0.9rem, 3vw, 1.1rem)',
              cursor: selectedFile && !isUploading ? 'pointer' : 'not-allowed',
              boxShadow: selectedFile && !isUploading 
                ? '0 4px 16px rgba(56,189,248,0.3)' 
                : 'none',
              transition: 'all 0.3s ease',
              display: 'flex',
              alignItems: 'center',
              gap: 'clamp(0.3rem, 1vw, 0.5rem)',
              animation: 'fadeInUp 1s ease-out 0.7s both',
              minHeight: 'clamp(2.5rem, 8vw, 3rem)',
              width: 'fit-content',
              maxWidth: '100%',
            }}
          >
            {isUploading ? (
              <>
                <Clock size="clamp(16px, 4vw, 20px)" />
                <span style={{ whiteSpace: 'nowrap' }}>Generating Conflict-Free Timetable...</span>
              </>
            ) : (
              <>
                <Upload size="clamp(16px, 4vw, 20px)" />
                <span style={{ whiteSpace: 'nowrap' }}>Generate Conflict-Free Timetable</span>
              </>
            )}
          </button>

          <button
            onClick={() => {
              // Handle manual timetable editing
              console.log('Edit timetable manually');
            }}
            style={{
              background: 'linear-gradient(90deg, #8b5cf6 0%, #7c3aed 100%)',
              color: '#fff',
              border: 'none',
              borderRadius: 'clamp(8px, 2vw, 12px)',
              padding: 'clamp(0.75rem, 3vw, 1rem) clamp(1.5rem, 4vw, 2rem)',
              fontWeight: 700,
              fontSize: 'clamp(0.9rem, 3vw, 1.1rem)',
              cursor: 'pointer',
              boxShadow: '0 4px 16px rgba(139,92,246,0.3)',
              transition: 'all 0.3s ease',
              display: 'flex',
              alignItems: 'center',
              gap: 'clamp(0.3rem, 1vw, 0.5rem)',
              animation: 'fadeInUp 1s ease-out 0.8s both',
              minHeight: 'clamp(2.5rem, 8vw, 3rem)',
              width: 'fit-content',
              maxWidth: '100%',
            }}
          >
            <Edit3 size="clamp(16px, 4vw, 20px)" />
            <span style={{ whiteSpace: 'nowrap' }}>Edit Timetable Manually</span>
          </button>
        </div>

        {/* Instructions */}
        <div style={{
          width: '100%',
          background: 'rgba(35,39,47,0.6)',
          borderRadius: 'clamp(12px, 3vw, 16px)',
          padding: 'clamp(1.5rem, 4vw, 2rem)',
          border: '1px solid rgba(56,189,248,0.2)',
          animation: 'fadeInUp 1s ease-out 0.9s both',
        }}>
          <h4 style={{ 
            color: '#38bdf8', 
            fontSize: 'clamp(1rem, 3vw, 1.2rem)', 
            fontWeight: 700, 
            marginBottom: 'clamp(0.75rem, 2vw, 1rem)' 
          }}>
            Instructions
          </h4>
          <ul style={{ 
            color: '#cbd5e1', 
            fontSize: 'clamp(0.85rem, 2.5vw, 0.95rem)', 
            lineHeight: '1.6',
            paddingLeft: 'clamp(1rem, 3vw, 1.5rem)'
          }}>
            <li>Upload an Excel file (.xlsx or .xls) containing your timetable data</li>
            <li>The file should include columns for: Teachers, Subjects, Classes, Time Slots</li>
            <li>Our AI will analyze the data and generate a conflict-free schedule</li>
            <li>The generated timetable will be optimized for all constraints</li>
            <li>You can download the final timetable once processing is complete</li>
          </ul>
        </div>
      </main>

      {/* Animations */}
      <style>{`
        @keyframes fadeInUp {
          from { opacity: 0; transform: translateY(30px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeInUpAdmin {
          from { opacity: 0; transform: translateY(40px); }
          to { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeInUpAdminTitle {
          from { opacity: 0; transform: translateY(60px) scale(0.95); }
          to { opacity: 1; transform: translateY(0) scale(1); }
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
      `}</style>
    </div>
  );
}

export default TimetableUpload; 