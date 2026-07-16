import { IconAlertTriangle, IconTrash } from './Icons';

const BACKEND_URL = 'http://127.0.0.1:8000';

export default function SettingsPanel({ stages, selectedStageId, onDeleteStage, onClearDatabase, showToast }) {
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
        </div>
      </div>

      {/* Hapus Tahap Tertentu */}
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
                className={`stage-delete-item ${selectedStageId === stage.id.toString() ? 'active' : ''}`}
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

      {/* Cleanup Rekonsiliasi */}
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
    </div>
  );
}
