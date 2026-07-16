import { useState, useEffect } from 'react';
import { IconEdit } from './Icons';

export default function ManualEditModal({ show, onClose, record, onSave }) {
  const [editNama, setEditNama] = useState('');
  const [editKtp, setEditKtp] = useState('');
  const [editKk, setEditKk] = useState('');

  useEffect(() => {
    if (record) {
      setEditNama(record.nama);
      setEditKtp(record.no_ktp);
      setEditKk(record.no_kk);
    }
  }, [record]);

  if (!show || !record) return null;

  const handleSave = () => {
    onSave({
      ...record,
      nama: editNama,
      no_ktp: editKtp,
      no_kk: editKk
    });
  };

  return (
    <div className="modal-overlay">
      <div className="modal-content" style={{ width: '600px' }}>
        <div className="modal-header">
          <h3><IconEdit />Form Perbaikan Data Manual</h3>
          <button 
            type="button" 
            className="modal-close"
            onClick={onClose}
          >
            &times;
          </button>
        </div>
        <div className="modal-body">
          <div className="form-group">
            <label className="form-label">Perbaikan Nama</label>
            <input type="text" className="form-input" value={editNama} onChange={e => setEditNama(e.target.value)} />
          </div>
          <div className="form-group">
            <label className="form-label">Perbaikan NIK (KTP)</label>
            <input type="text" className="form-input" value={editKtp} onChange={e => setEditKtp(e.target.value)} />
          </div>
          <div className="form-group">
            <label className="form-label">Perbaikan No. KK</label>
            <input type="text" className="form-input" value={editKk} onChange={e => setEditKk(e.target.value)} />
          </div>
        </div>
        <div className="modal-footer">
          <button className="btn btn-secondary btn-sm" onClick={onClose}>Batal</button>
          <button className="btn btn-primary btn-sm" onClick={handleSave}>Terapkan Perbaikan</button>
        </div>
      </div>
    </div>
  );
}
