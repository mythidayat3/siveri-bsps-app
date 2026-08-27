import React, { useState } from 'react';
import { IconAlertTriangle, IconTrash } from './Icons';

const BACKEND_URL = import.meta.env.VITE_BACKEND_URL || 'http://127.0.0.1:8000';

export default function SettingsPanel({ stages, selectedStageId, onDeleteStage, onClearDatabase, showToast, isAdmin }) {
  const [changeTarget, setChangeTarget] = useState('balai_sul_3');
  const [newPassword, setNewPassword] = useState('');
  const [passLoading, setPassLoading] = useState(false);

  const handlePasswordSubmit = async (e) => {
    e.preventDefault();
    if (!newPassword || newPassword.length < 4) {
      showToast("Password minimal 4 karakter", "error");
      return;
    }
    setPassLoading(true);
    try {
      const res = await fetch(`${BACKEND_URL}/api/change-password`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username: changeTarget, new_password: newPassword })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || "Gagal memperbarui password");
      showToast(data.message || "Password berhasil diperbarui!");
      setNewPassword('');
    } catch (err) {
      showToast(err.message, "error");
    } finally {
      setPassLoading(false);
    }
  };

  const getSortedStages = () => {
    return [...stages].sort((a, b) => {
      const isPenggantiA = a.name.toLowerCase().includes('pengganti');
      const isPenggantiB = b.name.toLowerCase().includes('pengganti');
      if (isPenggantiA !== isPenggantiB) return isPenggantiA ? 1 : -1;
      const numA = parseInt(a.name.replace(/\D/g, '')) || 0;
      const numB = parseInt(b.name.replace(/\D/g, '')) || 0;
      return numA - numB;
    });
  };

  const handleCleanupOverrides = async () => {
    if (!window.confirm("Bersihkan data rekonsiliasi yang sudah tidak terkait dengan data verifikasi?")) {
      return;
    }
    try {
      const res = await fetch(`${BACKEND_URL}/api/database/cleanup-overrides`, {
        method: 'POST'
      });
      if (!res.ok) throw new Error("Gagal membersihkan data");
      const data = await res.json();
      showToast(data.message, 'success');
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  return (
    <div className="settings-container">
      <div className="settings-card">
        <h3>
          <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}>
            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
            <polyline points="7 10 12 15 17 10"></polyline>
            <line x1="12" y1="15" x2="12" y2="3"></line>
          </svg>
          Unduh Berkas Template Excel
        </h3>
        <p>Pastikan file Excel yang akan diunggah sesuai dengan struktur template berikut agar data dapat terbaca oleh sistem.</p>
        
        <div className="template-download-grid">
          <div className="template-download-card">
            <div className="template-info">
              <span className="template-name">Format INVERS (Rujukan)</span>
              <span className="template-desc">Wajib kolom: Nama, NIK, KK, Alamat, Desa, Kecamatan</span>
            </div>
            <a href={`${BACKEND_URL}/api/templates/download/invers`} download className="btn btn-secondary btn-sm">
              <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}>
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="7 10 12 15 17 10"></polyline>
                <line x1="12" y1="15" x2="12" y2="3"></line>
              </svg>
              Unduh (.xlsx)
            </a>
          </div>

          <div className="template-download-card">
            <div className="template-info">
              <span className="template-name">Format Hasil Verifikasi Lapangan</span>
              <span className="template-desc">Sheet 1: Lamp IIA (Lolos) &bull; Sheet 2: Lamp IIIA (Tidak Lolos)</span>
            </div>
            <a href={`${BACKEND_URL}/api/templates/download/verified`} download className="btn btn-secondary btn-sm">
              <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}>
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="7 10 12 15 17 10"></polyline>
                <line x1="12" y1="15" x2="12" y2="3"></line>
              </svg>
              Unduh (.xlsx)
            </a>
          </div>

          <div className="template-download-card">
            <div className="template-info">
              <span className="template-name">Format SK Dirjen</span>
              <span className="template-desc">Wajib kolom: Nama, NIK, KK, Alamat, Desa, Kecamatan, Kabupaten</span>
            </div>
            <a href={`${BACKEND_URL}/api/templates/download/sk_dirjen`} download className="btn btn-secondary btn-sm">
              <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}>
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="7 10 12 15 17 10"></polyline>
                <line x1="12" y1="15" x2="12" y2="3"></line>
              </svg>
              Unduh (.xlsx)
            </a>
          </div>

          <div className="template-download-card">
            <div className="template-info">
              <span className="template-name">Format Database Kode Desa</span>
              <span className="template-desc">Struktur: Kode Desa, Provinsi, Kab./Kota, Kecamatan, Desa/Kelurahan</span>
            </div>
            <a href={`${BACKEND_URL}/api/templates/download/village_codes`} download className="btn btn-secondary btn-sm">
              <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}>
                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                <polyline points="7 10 12 15 17 10"></polyline>
                <line x1="12" y1="15" x2="12" y2="3"></line>
              </svg>
              Unduh (.xlsx)
            </a>
          </div>
        </div>
      </div>

      {/* Hapus Tahap Tertentu (Admin Only) */}
      {isAdmin && (
        <div className="settings-card">
          <h3>
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle", color: "var(--danger)" }}>
              <polyline points="3 6 5 6 21 6"></polyline>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            </svg>
            Hapus Tahap INVERS
          </h3>
          <p>Pilih tahap yang ingin dihapus. Semua data terkait (INVERS, Berita Acara, Verifikasi, Rekonsiliasi) akan ikut terhapus secara permanen.</p>
          
          {stages.length === 0 ? (
            <div style={{ color: 'var(--text-muted)', fontSize: '0.9rem', fontStyle: 'italic' }}>
              Belum ada tahap INVERS yang tersedia.
            </div>
          ) : (
            <div className="stage-delete-list">
              {getSortedStages().map(stage => (
                <div 
                  key={stage.id} 
                  className={`stage-delete-item ${String(selectedStageId) === String(stage?.id) ? 'active' : ''}`}
                >
                  <div className="stage-delete-info">
                    <span className="stage-delete-name">{stage.name}</span>
                    <span className="stage-delete-meta">
                      {stage.record_count} CPB &bull; Rev. {stage.max_revision || 1}
                    </span>
                  </div>
                  <button 
                    className="btn btn-danger btn-sm"
                    onClick={() => onDeleteStage(stage.id, stage.name)}
                    title={`Hapus tahap ${stage.name} beserta semua data`}
                  >
                    <IconTrash /> Hapus
                  </button>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Cleanup Rekonsiliasi (Admin Only) */}
      {isAdmin && (
        <div className="settings-card">
          <h3>
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle", color: "var(--warning)" }}>
              <path d="M23 4v6h-6"></path>
              <path d="M1 20v-6h6"></path>
              <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path>
            </svg>
            Bersihkan Data Rekonsiliasi
          </h3>
          <p>Jika Anda melihat hasil rekonsiliasi yang masih muncul setelah menghapus batch, gunakan tombol ini untuk membersihkan data rekonsiliasi yang sudah tidak terkait.</p>
          <button 
            className="btn btn-secondary"
            onClick={handleCleanupOverrides}
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}>
              <path d="M23 4v6h-6"></path>
              <path d="M1 20v-6h6"></path>
              <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path>
            </svg>
            Bersihkan Rekonsiliasi orphaned
          </button>
        </div>
      )}

      {/* Pengaturan Password Akun (Admin Only) */}
      {isAdmin && (
        <div className="settings-card">
          <h3>
            <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}>
              <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
              <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
            </svg>
            Kelola Password Akun (Kredensial)
          </h3>
          <p>Ubah password untuk akun Admin (`yayatbalai`, `balaip3kp`) atau Viewer (`balai_sul_3`) sesuai kebutuhan Anda.</p>

          <form onSubmit={handlePasswordSubmit} style={{ maxWidth: '400px' }}>
            <div style={{ marginBottom: '12px' }}>
              <label style={{ display: 'block', fontSize: '0.8rem', fontWeight: '600', marginBottom: '4px' }}>Pilih Akun:</label>
              <select
                value={changeTarget}
                onChange={e => setChangeTarget(e.target.value)}
                style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e1', fontSize: '0.85rem' }}
              >
                <option value="balai_sul_3">Viewer (balai_sul_3)</option>
                <option value="yayatbalai">Admin 1 (yayatbalai)</option>
                <option value="balaip3kp">Admin 2 (balaip3kp)</option>
              </select>
            </div>

            <div style={{ marginBottom: '16px' }}>
              <label style={{ display: 'block', fontSize: '0.8rem', fontWeight: '600', marginBottom: '4px' }}>Password Baru:</label>
              <input
                type="password"
                value={newPassword}
                onChange={e => setNewPassword(e.target.value)}
                placeholder="Masukkan password baru..."
                required
                style={{ width: '100%', padding: '8px 12px', borderRadius: '6px', border: '1px solid #cbd5e1', fontSize: '0.85rem' }}
              />
            </div>

            <button type="submit" className="btn btn-primary btn-sm" disabled={passLoading}>
              {passLoading ? 'Menyimpan...' : 'Perbarui Password'}
            </button>
          </form>
        </div>
      )}

      {isAdmin && (
        <div className="settings-card danger-zone">
          <h3><IconAlertTriangle />Area Bahaya (Danger Zone)</h3>
          <p>Opsi ini akan menghapus seluruh data transaksi di dalam database, termasuk tahap verifikasi dan perbaikan rekonsiliasi yang sudah dikerjakan.</p>
          <button 
            className="btn btn-danger"
            onClick={onClearDatabase}
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}>
              <polyline points="3 6 5 6 21 6"></polyline>
              <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
            </svg>
            Kosongkan & Reset Seluruh Database
          </button>
        </div>
      )}
    </div>
  );
}
