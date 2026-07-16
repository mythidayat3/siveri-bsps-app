import { useState } from 'react';
import { IconLink } from './Icons';

export default function LinkToInversModal({ show, onClose, linkingRecord, recordsData, onLink }) {
  const [linkSearchTerm, setLinkSearchTerm] = useState('');

  if (!show || !linkingRecord) return null;

  const filteredRecords = recordsData?.invers_records?.filter(ir => {
    const term = linkSearchTerm.toLowerCase();
    return ir.nama.toLowerCase().includes(term) ||
      ir.no_ktp.includes(term) ||
      ir.no_kk.includes(term) ||
      (ir.desa_kelurahan || '').toLowerCase().includes(term);
  }).slice(0, 15) || [];

  return (
    <div className="modal-overlay">
      <div className="modal-content" style={{ width: '700px', maxWidth: '90%' }}>
        <div className="modal-header">
          <h3><IconLink />Pasangkan dengan Data INVERS Rujukan</h3>
          <button 
            type="button" 
            className="modal-close"
            onClick={onClose}
          >
            &times;
          </button>
        </div>
        
        <div className="modal-body">
          <div style={{ backgroundColor: '#f1f5f9', padding: '12px', borderRadius: '6px', marginBottom: '16px', fontSize: '0.85rem' }}>
            <strong>Data Lapangan yang Dipilih:</strong>
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px', marginTop: '6px' }}>
              <div>Nama: {linkingRecord.nama}</div>
              <div>NIK: {linkingRecord.no_ktp}</div>
              <div>No. KK: {linkingRecord.no_kk}</div>
              <div>Desa: {linkingRecord.desa_kelurahan}</div>
            </div>
          </div>
          
          <div className="form-group" style={{ marginBottom: '16px' }}>
            <label className="form-label">Cari Calon Penerima di Data INVERS</label>
            <input 
              type="text" 
              className="form-input" 
              placeholder="Ketik nama, NIK, KK, atau Desa untuk mencari..."
              value={linkSearchTerm}
              onChange={e => setLinkSearchTerm(e.target.value)}
              autoFocus
            />
          </div>
          
          <div style={{ maxHeight: '300px', overflowY: 'auto', border: '1px solid var(--border)', borderRadius: '6px' }}>
            <table className="data-table" style={{ fontSize: '0.85rem' }}>
              <thead>
                <tr>
                  <th>Nama</th>
                  <th>NIK</th>
                  <th>No. KK</th>
                  <th>Desa</th>
                  <th>Aksi</th>
                </tr>
              </thead>
              <tbody>
                {filteredRecords.map(ir => (
                  <tr key={ir.id}>
                    <td style={{ fontWeight: '600' }}>{ir.nama}</td>
                    <td className="mono-digit">{ir.no_ktp}</td>
                    <td className="mono-digit">{ir.no_kk}</td>
                    <td>{ir.desa_kelurahan}</td>
                    <td>
                      <button 
                        className="btn btn-primary btn-sm"
                        onClick={() => onLink(linkingRecord, ir)}
                      >
                        Pilih & Pasangkan
                      </button>
                    </td>
                  </tr>
                ))}
                {filteredRecords.length === 0 && (
                  <tr>
                    <td colSpan="5" style={{ textAlign: 'center', padding: '16px', color: 'var(--text-muted)' }}>
                      Tidak ditemukan data INVERS yang cocok.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
        
        <div className="modal-footer">
          <button 
            className="btn btn-secondary btn-sm" 
            onClick={onClose}
          >
            Batal
          </button>
        </div>
      </div>
    </div>
  );
}
