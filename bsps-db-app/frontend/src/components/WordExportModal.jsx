import { useState } from 'react';
import { IconFile } from './Icons';

export default function WordExportModal({ show, onClose, onSubmit }) {
  const [wordFormData, setWordFormData] = useState({
    nomor_ba: '',
    nomor_surat: '',
    tanggal_ba: '',
    lokasi_ba: '',
    no_surat_dirjen: '',
    tgl_surat_dirjen: '',
    hal_surat_dirjen: ''
  });

  if (!show) return null;

  const handleSubmit = (e) => {
    e.preventDefault();
    onSubmit(wordFormData);
  };

  return (
    <div className="modal-overlay">
      <form className="modal-content" onSubmit={handleSubmit}>
        <div className="modal-header">
          <h3><IconFile />Buat Dokumen Berita Acara & Surat Penyampaian</h3>
          <button 
            type="button" 
            className="modal-close"
            onClick={onClose}
          >
            &times;
          </button>
        </div>
        
        <div className="modal-body">
          <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
            Isi form berikut untuk mengganti kolom placeholder <code style={{color: 'var(--primary)'}}>[...]</code> yang ada di file template Word.
          </p>
          
          <div className="modal-form-grid">
            <div className="form-group">
              <label className="form-label">Nomor Berita Acara (BA)</label>
              <input 
                type="text" 
                placeholder="Contoh: 12/BA/PKP/2026"
                className="form-input"
                value={wordFormData.nomor_ba}
                onChange={e => setWordFormData({...wordFormData, nomor_ba: e.target.value})}
                required
              />
            </div>
            <div className="form-group">
              <label className="form-label">Nomor Surat Pengantar BA</label>
              <input 
                type="text" 
                placeholder="Contoh: 005/321/DPKP"
                className="form-input"
                value={wordFormData.nomor_surat}
                onChange={e => setWordFormData({...wordFormData, nomor_surat: e.target.value})}
                required
              />
            </div>
            <div className="form-group">
              <label className="form-label">Tanggal BA / Pengiriman</label>
              <input 
                type="text" 
                placeholder="Contoh: 07 Juli 2026"
                className="form-input"
                value={wordFormData.tanggal_ba}
                onChange={e => setWordFormData({...wordFormData, tanggal_ba: e.target.value})}
                required
              />
            </div>
            <div className="form-group">
              <label className="form-label">Tempat / Lokasi BA</label>
              <input 
                type="text" 
                placeholder="Contoh: Makassar"
                className="form-input"
                value={wordFormData.lokasi_ba}
                onChange={e => setWordFormData({...wordFormData, lokasi_ba: e.target.value})}
                required
              />
            </div>
          </div>

          <div style={{ borderTop: '1px solid var(--border)', paddingTop: '12px', marginTop: '6px' }}>
            <h4 style={{ fontSize: '0.9rem', marginBottom: '8px', color: 'var(--primary)' }}>Dasar Surat Dirjen (Rujukan)</h4>
            <div className="modal-form-grid">
              <div className="form-group">
                <label className="form-label">Nomor Surat Dirjen</label>
                <input 
                  type="text" 
                  placeholder="Contoh: RU.01.02-Dr/105"
                  className="form-input"
                  value={wordFormData.no_surat_dirjen}
                  onChange={e => setWordFormData({...wordFormData, no_surat_dirjen: e.target.value})}
                />
              </div>
              <div className="form-group">
                <label className="form-label">Tanggal Surat Dirjen</label>
                <input 
                  type="text" 
                  placeholder="Contoh: 15 Januari 2026"
                  className="form-input"
                  value={wordFormData.tgl_surat_dirjen}
                  onChange={e => setWordFormData({...wordFormData, tgl_surat_dirjen: e.target.value})}
                />
              </div>
            </div>
            <div className="form-group" style={{ marginTop: '8px' }}>
              <label className="form-label">Perihal Surat Dirjen</label>
              <input 
                type="text" 
                placeholder="Contoh: Penyampaian Data Calon Penerima Bantuan Swadaya..."
                className="form-input"
                value={wordFormData.hal_surat_dirjen}
                onChange={e => setWordFormData({...wordFormData, hal_surat_dirjen: e.target.value})}
              />
            </div>
          </div>
        </div>

        <div className="modal-footer">
          <button 
            type="button" 
            className="btn btn-secondary" 
            onClick={onClose}
          >
            Batal
          </button>
          <button type="submit" className="btn btn-primary">
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}>
              <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
              <polyline points="7 10 12 15 17 10"></polyline>
              <line x1="12" y1="15" x2="12" y2="3"></line>
            </svg>
            Buat ZIP Dokumen
          </button>
        </div>
      </form>
    </div>
  );
}
