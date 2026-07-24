import React, { useState, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import logopkp from './assets/LOGOPKP.svg';
import SettingsPanel from './components/SettingsPanel';

const BACKEND_URL = 'http://127.0.0.1:8000';

// Beautiful SVG Flat Icons
const IconDashboard = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
    <path d="M3 13h8V3H3v10zm0 8h8v-6H3v6zm10 0h8V11h-8v10zm0-18v6h8V3h-8z"/>
  </svg>
);

const IconInvers = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
    <path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/>
  </svg>
);

const IconVerified = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
    <path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z"/>
  </svg>
);

const IconOverview = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
    <path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z"/>
  </svg>
);

const IconReconcile = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
    <path d="M19 8l-4 4h3c0 3.31-2.69 6-6 6-1.01 0-1.97-.25-2.8-.7l-1.46 1.46C8.97 19.54 10.43 20 12 20c4.42 0 8-3.58 8-8h3l-4-4zM6 12c0-3.31 2.69-6 6-6 1.01 0 1.97.25 2.8.7l1.46-1.46C15.03 4.46 13.57 4 12 4c-4.42 0-8 3.58-8 8H1l4 4 4-4H6z"/>
  </svg>
);

const IconSettings = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
    <path d="M19.14 12.94c.04-.3.06-.61.06-.94 0-.32-.02-.64-.07-.94l2.03-1.58c.18-.14.23-.41.12-.61l-1.92-3.32c-.12-.22-.37-.29-.59-.22l-2.39.96c-.5-.38-1.03-.7-1.62-.94l-.36-2.54c-.04-.24-.24-.41-.48-.41h-3.84c-.24 0-.43.17-.47.41l-.36 2.54c-.59.24-1.13.57-1.62.94l-2.39-.96c-.22-.08-.47 0-.59.22L2.74 8.87c-.12.21-.08.47.12.61l2.03 1.58c-.05.3-.09.63-.09.94s.02.64.07.94l-2.03 1.58c-.18.14-.23.41-.12.61l1.92 3.32c.12.22.37.29.59.22l2.39-.96c.5.38 1.03.7 1.62.94l.36 2.54c.05.24.24.41.48.41h3.84c.24 0 .44-.17.47-.41l.36-2.54c.59-.24 1.13-.56 1.62-.94l2.39.96c.22.08.47 0 .59-.22l1.92-3.32c.12-.22.07-.47-.12-.61l-2.01-1.58zM12 15.6c-1.98 0-3.6-1.62-3.6-3.6s1.62-3.6 3.6-3.6 3.6 1.62 3.6 3.6-1.62 3.6-3.6 3.6z"/>
  </svg>
);

const IconHelp = () => (
  <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
    <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 16h-2v-2h2v2zm1.07-7.75l-.9.92C12.45 11.9 12 12.5 12 14h-2v-.5c0-1.1.45-2.1 1.17-2.83l1.24-1.26c.37-.36.59-.86.59-1.41 0-1.1-.9-2-2-2s-2 .9-2 2H7c0-2.76 2.24-5 5-5s5 2.24 5 5c0 1.04-.42 1.99-1.07 2.75z"/>
  </svg>
);

// Expandable Kabupaten Row for "Rekap Per Kabupaten" tab
function KabupatenPengusulRow({ kab }) {
  const [isExpanded, setIsExpanded] = useState(false);
  return (
    <div style={{ display: 'flex', flexDirection: 'column' }}>
      <div 
        className={`tree-row ${isExpanded ? 'active-row' : ''}`}
        style={{ 
          display: 'grid',
          gridTemplateColumns: 'minmax(200px, 2fr) repeat(5, 90px)',
          alignItems: 'center',
          padding: '12px 14px',
          borderBottom: '1px solid #e9ecef',
          cursor: 'pointer'
        }}
        onClick={() => setIsExpanded(!isExpanded)}
      >
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          <span 
            className={`tree-arrow ${isExpanded ? 'expanded' : ''}`}
            style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '20px', height: '20px' }}
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 18 15 12 9 6"></polyline></svg>
          </span>
          <span style={{ fontWeight: '700', fontSize: '0.9rem' }}>{kab.name}</span>
        </div>
        <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem' }}>{kab.cpb}</div>
        <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem', color: 'var(--success)' }}>{kab.lolos}</div>
        <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem', color: '#856404' }}>{kab.tidak_lolos}</div>
        <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem', color: 'var(--danger)' }}>{kab.belum_verifikasi}</div>
        <div style={{ textAlign: 'center' }}>
          <span 
            style={{ 
              cursor: 'pointer', fontSize: '0.75rem', color: 'var(--primary)',
              fontWeight: '600', padding: '2px 8px', borderRadius: '4px',
              backgroundColor: isExpanded ? 'var(--primary-light)' : 'transparent'
            }}
          >
            {isExpanded ? 'Tutup' : 'Buka'}
          </span>
        </div>
      </div>
      {isExpanded && kab.children && kab.children.length > 0 && (
        <div style={{ marginLeft: '20px' }}>
          {kab.children.map((pengusul, idx) => (
            <div 
              key={idx}
              style={{ 
                display: 'grid',
                gridTemplateColumns: 'minmax(200px, 2fr) repeat(5, 90px)',
                alignItems: 'center',
                padding: '10px 14px',
                borderBottom: '1px dashed #e9ecef',
                backgroundColor: '#fafbfc'
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px', paddingLeft: '20px' }}>
                <span style={{ borderLeft: '2px dashed #dee2e6', marginLeft: '9px', height: '100%', display: 'inline-block' }}></span>
                <span style={{ fontWeight: '500', fontSize: '0.85rem' }}>{pengusul.name}</span>
              </div>
              <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem' }}>{pengusul.cpb}</div>
              <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem', color: 'var(--success)' }}>{pengusul.lolos}</div>
              <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem', color: '#856404' }}>{pengusul.tidak_lolos}</div>
              <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem', color: 'var(--danger)' }}>{pengusul.belum_verifikasi}</div>
              <div></div>
            </div>
          ))}
        </div>
      )}
      {isExpanded && (!kab.children || kab.children.length === 0) && (
        <div style={{ padding: '12px 14px 12px 48px', color: 'var(--text-muted)', fontStyle: 'italic', fontSize: '0.85rem', backgroundColor: '#fafbfc', borderBottom: '1px solid #e9ecef' }}>
          Tidak ada data pengusul untuk kabupaten ini.
        </div>
      )}
    </div>
  );
}

// Recursive Tree Node Component for Pengusul Center (Table-like)
function TreeNode({ node, level = 0, columns = 5 }) {
  const [isExpanded, setIsExpanded] = useState(false);
  const hasChildren = node.children && node.children.length > 0;

  return (
    <div style={{ display: 'flex', flexDirection: 'column' }}>
      <div 
        className={`tree-row ${isExpanded ? 'active-row' : ''}`}
        style={{ 
          paddingLeft: `${14 + level * 20}px`,
          display: 'grid',
          gridTemplateColumns: `minmax(200px, 2fr) repeat(${columns}, 90px)`,
          alignItems: 'center',
          gap: '0',
          borderBottom: '1px solid #e9ecef'
        }}
      >
        <div className="tree-node-title" style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
          {hasChildren ? (
            <span 
              className={`tree-arrow ${isExpanded ? 'expanded' : ''}`} 
              style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', width: '20px', height: '20px', cursor: 'pointer' }}
              onClick={() => setIsExpanded(!isExpanded)}
            >
              <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 18 15 12 9 6"></polyline></svg>
            </span>
          ) : (
            <span style={{ width: '20px', display: 'inline-block', borderLeft: '2px dashed #dee2e6', marginLeft: '9px', height: '100%' }}></span>
          )}
          <span style={{ fontWeight: level === 0 ? '700' : level === 1 ? '600' : '500', fontSize: level === 0 ? '0.9rem' : '0.85rem' }}>{node.name}</span>
        </div>
        <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem' }}>{node.cpb}</div>
        <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem', color: 'var(--success)' }}>{node.lolos}</div>
        <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem', color: '#856404' }}>{node.tidak_lolos}</div>
        <div style={{ textAlign: 'center', fontWeight: '600', fontSize: '0.85rem', color: 'var(--danger)' }}>{node.belum_verifikasi}</div>
        <div style={{ textAlign: 'center' }}>
          {hasChildren && (
            <span 
              onClick={() => setIsExpanded(!isExpanded)}
              style={{ 
                cursor: 'pointer', 
                fontSize: '0.75rem', 
                color: 'var(--primary)',
                fontWeight: '600',
                padding: '2px 8px',
                borderRadius: '4px',
                backgroundColor: isExpanded ? 'var(--primary-light)' : 'transparent'
              }}
            >
              {isExpanded ? 'Tutup' : 'Buka'}
            </span>
          )}
        </div>
      </div>
      {hasChildren && isExpanded && (
        <div className="tree-children" style={{ marginLeft: `${level === 0 ? 12 : 20}px` }}>
          {node.children.map((child, idx) => (
            <TreeNode key={idx} node={child} level={level + 1} columns={columns} />
          ))}
        </div>
      )}
    </div>
  );
}

function App() {
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
  const [activeTab, setActiveTab] = useState('dashboard');
  const [stages, setStages] = useState([]);
  const [selectedStageId, setSelectedStageId] = useState('');
  const [stageSummary, setStageSummary] = useState(null);
  const [darkMode, setDarkMode] = useState(() => localStorage.getItem('darkMode') === 'true');
  const [recordsData, setRecordsData] = useState(null);
  
  // Drag and Drop Batch Reordering
  const [draggedBatchIndex, setDraggedBatchIndex] = useState(null);
  const [dragOverBatchIndex, setDragOverBatchIndex] = useState(null);

  const handleBatchDragStart = (e, index) => {
    setDraggedBatchIndex(index);
    e.dataTransfer.effectAllowed = 'move';
    e.dataTransfer.setData('text/plain', index);
  };

  const handleBatchDragOver = (e, index) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = 'move';
    if (dragOverBatchIndex !== index) {
      setDragOverBatchIndex(index);
    }
  };

  const handleBatchDrop = async (e, dropIndex) => {
    e.preventDefault();
    if (draggedBatchIndex === null || draggedBatchIndex === dropIndex) {
      setDraggedBatchIndex(null);
      setDragOverBatchIndex(null);
      return;
    }

    const currentBatches = [...(stageSummary?.batches || [])];
    const [draggedItem] = currentBatches.splice(draggedBatchIndex, 1);
    currentBatches.splice(dropIndex, 0, draggedItem);

    setStageSummary(prev => ({ ...prev, batches: currentBatches }));
    setDraggedBatchIndex(null);
    setDragOverBatchIndex(null);

    const orders = currentBatches.map((b, idx) => ({ id: b.id, sort_order: idx + 1 }));

    try {
      const res = await fetch(`${BACKEND_URL}/api/verified/batches/reorder`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ orders })
      });
      if (res.ok) {
        showToast('Urutan Berita Acara berhasil diperbarui', 'success');
        if (typeof fetchRekapBatchBA === 'function') fetchRekapBatchBA();
      } else {
        showToast('Gagal menyimpan urutan batch', 'error');
        if (selectedStageId) fetchStageSummary(selectedStageId);
      }
    } catch (err) {
      showToast('Terjadi kesalahan saat menyimpan urutan batch', 'error');
      if (selectedStageId) fetchStageSummary(selectedStageId);
    }
  };
  
  // Dashboard & Rekap Center Data
  const [overviewStats, setOverviewStats] = useState(null);
  const [overviewTables, setOverviewTables] = useState(null);
  const [pengusulTree, setPengusulTree] = useState([]);
  
  // State Unggah & Modals
  const [uploadType, setUploadType] = useState(''); // 'invers' atau 'verified'
  const [stageNameInput, setStageNameInput] = useState('');
  const [batchNameInput, setBatchNameInput] = useState('Berita Acara Pertama');
  const [selectedFile, setSelectedFile] = useState(null);
  const [isDragOver, setIsDragOver] = useState(false);
  
  // State Pencarian & Filter
  const [searchTerm, setSearchTerm] = useState('');
  const [errorFilter, setErrorFilter] = useState('ALL');
  
  // State Filter Tabel INVERS
  const [inversKabFilter, setInversKabFilter] = useState('ALL');
  const [inversDesaFilter, setInversDesaFilter] = useState('ALL');
  const [inversStatusFilter, setInversStatusFilter] = useState('ALL');
  
  // State Filter Tabel Terverifikasi
  const [verifiedKabFilter, setVerifiedKabFilter] = useState('ALL');
  const [verifiedDesaFilter, setVerifiedDesaFilter] = useState('ALL');
  const [verifiedHasilFilter, setVerifiedHasilFilter] = useState('ALL');

  // State Filter Tabel Berita Acara
  const [baSearchTerm, setBaSearchTerm] = useState('');
  const [baFilterStatus, setBaFilterStatus] = useState('all');
  const [verifiedBatchFilter, setVerifiedBatchFilter] = useState('ALL');
  
  // State Edit Formulir (Rekonsiliasi Manual)
  const [editingRecord, setEditingRecord] = useState(null);
  const [editNama, setEditNama] = useState('');
  const [editKtp, setEditKtp] = useState('');
  const [editKk, setEditKk] = useState('');

  // State Invers Manual Pairs (rekonsiliasi invers ↔ verified)
  const [unmatchedInvers, setUnmatchedInvers] = useState([]);
  const [manualPairs, setManualPairs] = useState([]);
  const [unmatchedVerified, setUnmatchedVerified] = useState([]);
  const [showPairModal, setShowPairModal] = useState(false);
  const [pairingInvers, setPairingInvers] = useState(null);
  const [pairSearchTerm, setPairSearchTerm] = useState('');
  const [pairSearchResults, setPairSearchResults] = useState([]);
  const [autoPairing, setAutoPairing] = useState(false);
  const [suggestedPairs, setSuggestedPairs] = useState([]);
  const [selectedSuggestions, setSelectedSuggestions] = useState(new Set());
  const [batchPairing, setBatchPairing] = useState(false);
  const [inlineSuggestions, setInlineSuggestions] = useState({});
  const [skippedNiks, setSkippedNiks] = useState(new Set());
  const [sectionCollapsed, setSectionCollapsed] = useState({
    belumVerif: true, pasangan: true, verifiedBelum: true
  });
  const [sectionPage, setSectionPage] = useState({
    belumVerif: 1, pasangan: 1, verifiedBelum: 1
  });

  // State Ekspor Word (Berita Acara)
  const [showWordModal, setShowWordModal] = useState(false);
  const [selectedBatchIdForWord, setSelectedBatchIdForWord] = useState(null);
  const [inversPage, setInversPage] = useState(1);
  const [verifiedPage, setVerifiedPage] = useState(1);
  const ITEMS_PER_PAGE = 30;

  // State Pasangkan Manual ke INVERS
  const [showLinkModal, setShowLinkModal] = useState(false);
  const [linkingRecord, setLinkingRecord] = useState(null);
  const [linkSearchTerm, setLinkSearchTerm] = useState('');

  const [wordFormData, setWordFormData] = useState({
    nomor_ba: '',
    nomor_surat: '',
    tanggal_ba: '',
    lokasi_ba: '',
    no_surat_dirjen: '',
    tgl_surat_dirjen: '',
    hal_surat_dirjen: ''
  });

  // Word form template save/load (localStorage)
  const getTemplateKey = (stageId, batchId) => `bsps_ba_template_${stageId}_${batchId}`;

  const getSavedTemplate = () => {
    if (!selectedStageId || !selectedBatchIdForWord) return null;
    try {
      const saved = localStorage.getItem(getTemplateKey(selectedStageId, selectedBatchIdForWord));
      return saved ? JSON.parse(saved) : null;
    } catch { return null; }
  };

  const handleSaveTemplate = async () => {
    if (!selectedStageId || !selectedBatchIdForWord) return;
    try {
      localStorage.setItem(getTemplateKey(selectedStageId, selectedBatchIdForWord), JSON.stringify(wordFormData));
      
      const res = await fetch(`${BACKEND_URL}/api/verified/batch/${selectedBatchIdForWord}/save-metadata`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          nomor_ba: wordFormData.nomor_ba,
          tanggal_ba: wordFormData.tanggal_ba
        })
      });
      
      if (!res.ok) {
        throw new Error("Gagal mengirim metadata ke server");
      }
      
      showToast("Template form dan metadata Berita Acara berhasil disimpan!");
      fetchRekapBatchBA();
    } catch (err) {
      showToast("Gagal menyimpan template: " + err.message, "error");
    }
  };

  const handleLoadTemplate = () => {
    const saved = getSavedTemplate();
    if (saved) {
      setWordFormData(saved);
      showToast("Template form berhasil dimuat!");
    } else {
      showToast("Belum ada template tersimpan untuk batch ini", "error");
    }
  };

  const handleDeleteTemplate = () => {
    if (!selectedStageId || !selectedBatchIdForWord) return;
    if (!window.confirm("Hapus template form yang tersimpan untuk batch ini?")) return;
    try {
      localStorage.removeItem(getTemplateKey(selectedStageId, selectedBatchIdForWord));
      showToast("Template form berhasil dihapus!");
    } catch {
      showToast("Gagal menghapus template", "error");
    }
  };

  // State Overview Subtabs
  const [overviewSubTab, setOverviewSubTab] = useState('summary_table'); // 'summary_table', 'tree_pengusul', atau 'kabupaten_pengusul'
  const [selectedPengusulFilter, setSelectedPengusulFilter] = useState('ALL');
  const [kabPengusulTree, setKabPengusulTree] = useState([]);

  // State Rekap Keseluruhan
  const [rekapData, setRekapData] = useState(null);
  const [rekapUnggahanData, setRekapUnggahanData] = useState(null);
  const [rekapLoading, setRekapLoading] = useState(false);

  // State Rekap Batch Berita Acara
  const [rekapBatchData, setRekapBatchData] = useState(null);
  const [rekapBatchLoading, setRekapBatchLoading] = useState(false);

  // State Pencarian Global
  const [globalSearchQuery, setGlobalSearchQuery] = useState('');
  const [globalFilterKab, setGlobalFilterKab] = useState('ALL');
  const [globalFilterKec, setGlobalFilterKec] = useState('ALL');
  const [globalFilterDesa, setGlobalFilterDesa] = useState('ALL');
  const [globalFilterStatus, setGlobalFilterStatus] = useState('ALL');
  const [globalFilterTahap, setGlobalFilterTahap] = useState('ALL');
  const [globalFilterType, setGlobalFilterType] = useState('all');
  const [globalPage, setGlobalPage] = useState(1);
  const [globalData, setGlobalData] = useState(null);
  const [globalLoading, setGlobalLoading] = useState(false);

  // State Bulk Operations
  const [selectedRecordIds, setSelectedRecordIds] = useState(new Set());
  const [selectedMismatchNiks, setSelectedMismatchNiks] = useState(new Set());

  // State Batch Kabupaten Breakdown
  const [expandedBatchId, setExpandedBatchId] = useState(null);
  const [batchBreakdownCache, setBatchBreakdownCache] = useState({});

  // State Excel Preview Modal
  const [showPreviewModal, setShowPreviewModal] = useState(false);
  const [selectedBatchIdForPreview, setSelectedBatchIdForPreview] = useState(null);
  const [previewData, setPreviewData] = useState(null);
  const [previewTab, setPreviewTab] = useState('lolos');
  const [previewSearchTerm, setPreviewSearchTerm] = useState('');

  // State Rename Modal (Stage & Batch)
  const [showRenameModal, setShowRenameModal] = useState(false);
  const [renameTarget, setRenameTarget] = useState({ type: '', id: null, currentName: '' });
  const [newRenameName, setNewRenameName] = useState('');
  const [renameLoading, setRenameLoading] = useState(false);

  const openRenameStageModal = (stageId, currentName) => {
    setRenameTarget({ type: 'stage', id: stageId, currentName });
    setNewRenameName(currentName);
    setShowRenameModal(true);
  };

  const openRenameBatchModal = (batchId, currentName) => {
    setRenameTarget({ type: 'batch', id: batchId, currentName });
    setNewRenameName(currentName);
    setShowRenameModal(true);
  };

  const handleSaveRename = async () => {
    const trimmed = newRenameName.trim();
    if (!trimmed) {
      showToast("Nama tidak boleh kosong", "error");
      return;
    }
    if (trimmed === renameTarget.currentName) {
      setShowRenameModal(false);
      return;
    }

    setRenameLoading(true);
    try {
      if (renameTarget.type === 'stage') {
        const res = await fetch(`${BACKEND_URL}/api/stage/${renameTarget.id}/rename`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: trimmed })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || "Gagal mengubah nama tahap");
        showToast(data.message || "Nama Tahap berhasil diperbarui!");
        fetchStages();
        fetchStageData(renameTarget.id);
        if (activeTab === 'rekap-keseluruhan') fetchRekapKeseluruhan();
        if (activeTab === 'rekap-batch-ba') fetchRekapBatchBA();
      } else if (renameTarget.type === 'batch') {
        const res = await fetch(`${BACKEND_URL}/api/verified/batch/${renameTarget.id}/rename`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ name: trimmed })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || "Gagal mengubah nama batch");
        showToast(data.message || "Nama Batch berhasil diperbarui!");
        fetchStageData(selectedStageId);
        if (activeTab === 'rekap-batch-ba') fetchRekapBatchBA();
      }
      setShowRenameModal(false);
    } catch (err) {
      showToast(err.message, "error");
    } finally {
      setRenameLoading(false);
    }
  };

  // State SK Dirjen
  const [skDirjenActiveSubTab, setSkDirjenActiveSubTab] = useState('daftar-pb');
  const [skDirjenSubmenuOpen, setSkDirjenSubmenuOpen] = useState(false);
  const [skDirjenBatches, setSkDirjenBatches] = useState([]);
  const [skDirjenSelectedBatch, setSkDirjenSelectedBatch] = useState('all');
  const [skDirjenRecords, setSkDirjenRecords] = useState([]);
  const [skDirjenFilterKab, setSkDirjenFilterKab] = useState('');
  const [skDirjenFilterKec, setSkDirjenFilterKec] = useState('');
  const [skDirjenFilterDesa, setSkDirjenFilterDesa] = useState('');
  const [skDirjenFilterTahap, setSkDirjenFilterTahap] = useState('');
  const [skDirjenFilterAsalBatch, setSkDirjenFilterAsalBatch] = useState('');
  const [skDirjenFilterStatus, setSkDirjenFilterStatus] = useState('');
  const [skDirjenRekapPerTahap, setSkDirjenRekapPerTahap] = useState(null);
  const [skDirjenRekapPerKab, setSkDirjenRekapPerKab] = useState(null);
  const [skDirjenSelectedKabForRekap, setSkDirjenSelectedKabForRekap] = useState(null);
  const [skDirjenApprovalRecord, setSkDirjenApprovalRecord] = useState(null);
  const [skDirjenUploading, setSkDirjenUploading] = useState(false);
  const [skDirjenStageName, setSkDirjenStageName] = useState('');
  const [skDirjenPairingRecord, setSkDirjenPairingRecord] = useState(null);
  const [skDirjenPairSearchTerm, setSkDirjenPairSearchTerm] = useState('');
  const [skDirjenPairSearchResults, setSkDirjenPairSearchResults] = useState([]);
  const [skDirjenSearchTerm, setSkDirjenSearchTerm] = useState('');
  const [skDirjenDebouncedSearch, setSkDirjenDebouncedSearch] = useState('');
  const [sumberSelisihPopup, setSumberSelisihPopup] = useState(null);
  const [expandedSumber, setExpandedSumber] = useState(null);
  const [sumberDetail, setSumberDetail] = useState([]);
  const [sumberDetailLoading, setSumberDetailLoading] = useState(false);

  // State Notifikasi Toast
  const [toast, setToast] = useState(null);

  const showToast = (message, type = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 4000);
  };

  // Ambil data tahap awal
  const fetchStages = useCallback(async () => {
    try {
      const res = await fetch(`${BACKEND_URL}/api/stages`);
      if (!res.ok) throw new Error("Gagal mengambil data tahap");
      const data = await res.json();
      setStages(data);
      if (data.length > 0 && !selectedStageId) {
        setSelectedStageId(data[0].id.toString());
      }
    } catch (err) {
      showToast(err.message, 'error');
    }
  }, [selectedStageId]);

  useEffect(() => {
    fetchStages();
  }, [fetchStages]);

  // Dark Mode toggle
  useEffect(() => {
    document.documentElement.classList.toggle('dark', darkMode);
    localStorage.setItem('darkMode', darkMode);
  }, [darkMode]);

  // Ambil ringkasan dan data rekaman ketika tahap berubah
  useEffect(() => {
    setInversPage(1);
    setVerifiedPage(1);
    if (selectedStageId) {
      fetchStageData(selectedStageId);
      fetchUnmatchedInvers(selectedStageId);
      fetchManualPairs(selectedStageId);
      fetchUnmatchedVerified(selectedStageId);
    } else {
      setStageSummary(null);
      setRecordsData(null);
      setOverviewStats(null);
      setOverviewTables(null);
      setPengusulTree([]);
      setKabPengusulTree([]);
      setUnmatchedInvers([]);
      setManualPairs([]);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [selectedStageId]);

  // Reset halaman ketika keyword pencarian berubah
  useEffect(() => {
    setInversPage(1);
    setVerifiedPage(1);
  }, [searchTerm]);

  // Auto-load saved template when Word modal opens with a batch
  useEffect(() => {
    if (showWordModal && selectedBatchIdForWord && selectedStageId) {
      const saved = getSavedTemplate();
      const currentBatch = stageSummary?.batches?.find(b => b.id === selectedBatchIdForWord);
      
      const defaultData = {
        nomor_ba: currentBatch?.nomor_ba || '',
        nomor_surat: '',
        tanggal_ba: currentBatch?.tanggal_ba || '',
        lokasi_ba: '',
        no_surat_dirjen: '',
        tgl_surat_dirjen: '',
        hal_surat_dirjen: ''
      };
      
      if (saved) {
        setWordFormData({
          ...defaultData,
          ...saved,
          nomor_ba: currentBatch?.nomor_ba || saved.nomor_ba || '',
          tanggal_ba: currentBatch?.tanggal_ba || saved.tanggal_ba || ''
        });
      } else {
        setWordFormData(defaultData);
      }
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [showWordModal, selectedBatchIdForWord, selectedStageId, stageSummary]);

  const fetchStageData = async (stageId) => {
    setBatchBreakdownCache({});
    setExpandedBatchId(null);
    try {
      // 1. Ambil Ringkasan
      const sumRes = await fetch(`${BACKEND_URL}/api/stage/${stageId}/summary`);
      if (!sumRes.ok) throw new Error("Gagal mengambil ringkasan data");
      const sumData = await sumRes.json();
      setStageSummary(sumData);

      // 2. Ambil Rekaman Data
      const recRes = await fetch(`${BACKEND_URL}/api/stage/${stageId}/records`);
      if (!recRes.ok) throw new Error("Gagal mengambil rekaman data");
      const recData = await recRes.json();
      setRecordsData(recData);

      // 3. Ambil Dashboard Stats (Progress Bar)
      const statsRes = await fetch(`${BACKEND_URL}/api/stage/${stageId}/overview-stats`);
      if (statsRes.ok) {
        const statsData = await statsRes.json();
        setOverviewStats(statsData);
      }

      // 4. Ambil Summary Tables
      const tablesRes = await fetch(`${BACKEND_URL}/api/stage/${stageId}/overview-tables`);
      if (tablesRes.ok) {
        const tablesData = await tablesRes.json();
        setOverviewTables(tablesData);
      }

      // 5. Ambil Pengusul Accordion Tree
      const treeRes = await fetch(`${BACKEND_URL}/api/stage/${stageId}/pengusul-tree`);
      if (treeRes.ok) {
        const treeData = await treeRes.json();
        setPengusulTree(treeData);
      }

      // 6. Ambil Kabupaten > Pengusul Tree
      const kabTreeRes = await fetch(`${BACKEND_URL}/api/stage/${stageId}/kabupaten-pengusul-tree`);
      if (kabTreeRes.ok) {
        const kabTreeData = await kabTreeRes.json();
        setKabPengusulTree(kabTreeData);
      }
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const fetchRekapKeseluruhan = async () => {
    setRekapLoading(true);
    try {
      const res = await fetch(`${BACKEND_URL}/api/rekap-keseluruhan?published_only=1`);
      if (!res.ok) throw new Error("Gagal mengambil data rekap keseluruhan");
      const data = await res.json();
      setRekapData(data);
    } catch (err) {
      showToast(err.message, 'error');
    } finally {
      setRekapLoading(false);
    }
  };

  const fetchRekapUnggahan = async () => {
    setRekapLoading(true);
    try {
      const res = await fetch(`${BACKEND_URL}/api/rekap-keseluruhan`);
      if (!res.ok) throw new Error("Gagal mengambil data rekap unggahan");
      const data = await res.json();
      setRekapUnggahanData(data);
    } catch (err) {
      showToast(err.message, 'error');
    } finally {
      setRekapLoading(false);
    }
  };

  const fetchRekapBatchBA = async () => {
    setRekapBatchLoading(true);
    try {
      const res = await fetch(`${BACKEND_URL}/api/rekap-batch-ba?published_only=1`);
      if (!res.ok) throw new Error("Gagal mengambil data rekap batch berita acara");
      const data = await res.json();
      setRekapBatchData(data);
    } catch (err) {
      showToast(err.message, 'error');
    } finally {
      setRekapBatchLoading(false);
    }
  };

  // Navigate to Global Search with pre-applied filters (used by Rekap table clicks)
  const navigateToData = (type, { kab, tahap, status } = {}) => {
    setActiveTab('global-search');
    setGlobalFilterKab(kab || 'ALL');
    setGlobalFilterTahap(tahap ? String(tahap) : 'ALL');
    setGlobalFilterStatus(status || 'ALL');
    if (type === 'verified') setGlobalFilterType('verified');
    else if (type === 'invers') setGlobalFilterType('invers');
    else setGlobalFilterType('all');
    setGlobalFilterKec('ALL');
    setGlobalFilterDesa('ALL');
    setGlobalSearchQuery('');
    setGlobalPage(1);
  };

  const fetchGlobalSearch = useCallback(async (overridePage) => {
    setGlobalLoading(true);
    try {
      const params = new URLSearchParams();
      if (globalSearchQuery) params.set('q', globalSearchQuery);
      if (globalFilterKab !== 'ALL') params.set('kabupaten', globalFilterKab);
      if (globalFilterKec !== 'ALL') params.set('kecamatan', globalFilterKec);
      if (globalFilterDesa !== 'ALL') params.set('desa', globalFilterDesa);
      if (globalFilterStatus !== 'ALL') params.set('status', globalFilterStatus);
      if (globalFilterTahap !== 'ALL') params.set('tahap', globalFilterTahap);
      if (globalFilterType !== 'all') params.set('record_type', globalFilterType);
      params.set('page', overridePage || globalPage);
      params.set('limit', '30');
      const res = await fetch(`${BACKEND_URL}/api/global-search?${params}`);
      if (!res.ok) throw new Error("Gagal mengambil data pencarian global");
      const data = await res.json();
      setGlobalData(data);
    } catch (err) {
      showToast(err.message, 'error');
    } finally {
      setGlobalLoading(false);
    }
  }, [globalSearchQuery, globalFilterKab, globalFilterKec, globalFilterDesa,
      globalFilterStatus, globalFilterTahap, globalFilterType, globalPage]);

  // Debounced search for global search input
  useEffect(() => {
    const timer = setTimeout(() => {
      if (activeTab === 'global-search') {
        setGlobalPage(1);
        fetchGlobalSearch(1);
      }
    }, 400);
    return () => clearTimeout(timer);
  }, [globalSearchQuery, activeTab]);

  // Fetch when filters change (not search query - handled by debounce)
  useEffect(() => {
    if (activeTab === 'global-search') {
      setGlobalPage(1);
      fetchGlobalSearch(1);
    }
  }, [globalFilterKab, globalFilterKec, globalFilterDesa, globalFilterStatus, globalFilterTahap, globalFilterType]);

  // Fetch when page changes (pagination)
  useEffect(() => {
    if (activeTab === 'global-search' && globalPage > 0) {
      fetchGlobalSearch();
    }
  }, [globalPage]);

  // Debounce search for SK Dirjen
  useEffect(() => {
    const timer = setTimeout(() => {
      setSkDirjenDebouncedSearch(skDirjenSearchTerm);
    }, 300);
    return () => clearTimeout(timer);
  }, [skDirjenSearchTerm]);

  useEffect(() => {
    if (activeTab === 'sk-dirjen' && skDirjenSelectedBatch) {
      if (skDirjenSelectedBatch === 'all') {
        fetchSkDirjenAllRecords({
          q: skDirjenDebouncedSearch,
          kabupaten: skDirjenFilterKab,
          kecamatan: skDirjenFilterKec,
          desa: skDirjenFilterDesa,
          tahap: skDirjenFilterTahap,
          asal_batch: skDirjenFilterAsalBatch,
          status: skDirjenFilterStatus
        });
      } else {
        fetchSkDirjenRecords(skDirjenSelectedBatch, {
          q: skDirjenDebouncedSearch,
          kabupaten: skDirjenFilterKab,
          kecamatan: skDirjenFilterKec,
          desa: skDirjenFilterDesa,
          tahap: skDirjenFilterTahap,
          asal_batch: skDirjenFilterAsalBatch,
          status: skDirjenFilterStatus
        });
      }
    }
  }, [skDirjenDebouncedSearch]);

  const handleStageSelect = (e) => {
    setSelectedStageId(e.target.value);
  };

  const handleFileChange = (e) => {
    setSelectedFile(e.target.files[0]);
  };

  // Client-side schema validation before uploading Excel file
  const validateFileSchema = (file, type) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          
          if (type === 'invers') {
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            
            let hasNama = false;
            let hasNik = false;
            let hasKk = false;
            
            for (const row of json) {
              if (!Array.isArray(row)) continue;
              const cells = row.map(c => String(c || '').trim().replace('\n', ' ').toUpperCase());
              if (cells.includes('NAMA') && cells.some(c => c.includes('NIK') || c.includes('KTP') || c.includes('KARTU'))) {
                hasNama = true;
                hasNik = cells.some(c => c.includes('NIK') || c.includes('KTP') || c.includes('KARTU'));
                hasKk = cells.some(c => c.includes('KK') || c.includes('KELUARGA'));
                break;
              }
            }
            
            if (!hasNama || !hasNik || !hasKk) {
              resolve({ valid: false, error: 'Format kolom salah. Kolom wajib: NAMA, NO KTP/NIK, dan NO KK tidak terdeteksi.' });
              return;
            }
            resolve({ valid: true });
          } else {
            // type === 'verified'
            if (workbook.SheetNames.length < 2) {
              resolve({ valid: false, error: 'File verifikasi harus memiliki minimal 2 Sheet (Sheet 1: Lamp. IIA, Sheet 2: Lamp. IIIA)' });
              return;
            }
            // Sheet 1: Lamp IIA (Lolos)
            const sheet1 = workbook.Sheets[workbook.SheetNames[0]];
            const json1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
            let hasNama1 = false, hasNik1 = false, hasKk1 = false;
            for (const row of json1) {
              if (!Array.isArray(row)) continue;
              const cells = row.map(c => String(c || '').trim().replace('\n', ' ').toUpperCase());
              if (cells.includes('NAMA') && cells.some(c => c.includes('NIK') || c.includes('KTP') || c.includes('KARTU'))) {
                hasNama1 = true;
                hasNik1 = true;
                hasKk1 = cells.some(c => c.includes('KK') || c.includes('KELUARGA'));
                break;
              }
            }
            
            // Sheet 2: Lamp IIIA (Tidak Lolos / Pengganti)
            const sheet2 = workbook.Sheets[workbook.SheetNames[1]];
            const json2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });
            let hasNama2 = false;
            for (const row of json2) {
              if (!Array.isArray(row)) continue;
              const cells = row.map(c => String(c || '').trim().replace('\n', ' ').toUpperCase());
              if (cells.includes('NAMA') || cells.some(c => c.includes('NAMA (PENGGANTI)'))) {
                hasNama2 = true;
                break;
              }
            }
            
            if (!hasNama1 || !hasNik1 || !hasKk1) {
              resolve({ valid: false, error: 'Format kolom Sheet 1 (Lamp. IIA) tidak sesuai. Harus ada kolom NAMA, NIK, dan KK.' });
              return;
            }
            if (!hasNama2) {
              resolve({ valid: false, error: 'Format kolom Sheet 2 (Lamp. IIIA) tidak sesuai. Baris kolom nama tidak terdeteksi.' });
              return;
            }
            resolve({ valid: true });
          }
        } catch (err) {
          resolve({ valid: false, error: 'Gagal membaca skema berkas Excel: ' + err.message });
        }
      };
      reader.onerror = () => resolve({ valid: false, error: 'Gagal memuat berkas.' });
      reader.readAsArrayBuffer(file);
    });
  };

  const handleUploadSubmit = async (e) => {
    e.preventDefault();
    if (!selectedFile) {
      showToast("Silakan pilih file terlebih dahulu untuk diunggah", "error");
      return;
    }

    // Lakukan validasi skema sebelum upload
    showToast("Memvalidasi skema kolom berkas...", "info");
    const validationResult = await validateFileSchema(selectedFile, uploadType);
    if (!validationResult.valid) {
      showToast(validationResult.error, "error");
      return;
    }

    const formData = new FormData();
    formData.append('file', selectedFile);

    try {
      if (uploadType === 'invers') {
        if (!stageNameInput.trim()) {
          showToast("Silakan masukkan nama tahap", "error");
          return;
        }
        formData.append('stage_name', stageNameInput);
        const res = await fetch(`${BACKEND_URL}/api/invers/upload`, {
          method: 'POST',
          body: formData
        });
        if (!res.ok) {
          const errData = await res.json();
          throw new Error(errData.detail || "Gagal mengunggah file");
        }
        const data = await res.json();
        showToast(`Berhasil mengunggah INVERS Tahap: ${stageNameInput}. Ditambahkan ${data.inserted_records} data.`);
        setStageNameInput('');
        setSelectedFile(null);
        setUploadType('');
        await fetchStages();
        setSelectedStageId(data.stage_id.toString());
      } else {
        if (!selectedStageId) {
          showToast("Silakan pilih target Tahap INVERS terlebih dahulu", "error");
          return;
        }
        if (!batchNameInput.trim()) {
          showToast("Silakan masukkan nama Berita Acara / Batch", "error");
          return;
        }
        formData.append('stage_id', selectedStageId);
        formData.append('batch_name', batchNameInput);
        
        const res = await fetch(`${BACKEND_URL}/api/verified/upload`, {
          method: 'POST',
          body: formData
        });
        if (!res.ok) {
          const errData = await res.json();
          throw new Error(errData.detail || "Gagal mengunggah file");
        }
        const data = await res.json();
        const stats = data.stats;
        showToast(
          `Berhasil unggah batch: ${batchNameInput}. ` +
          `Lolos: ${stats.lolos_added} ditambahkan. ` +
          `Tidak Lolos: ${stats.tidak_lolos_added} ditambahkan.`
        );
        setSelectedFile(null);
        setUploadType('');
        fetchStageData(selectedStageId);
      }
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleDeleteBatch = async (batchId) => {
    if (!window.confirm("Apakah Anda yakin ingin menghapus batch ini? Semua data verifikasi dan hasil rekonsiliasi terkait akan ikut terhapus.")) {
      return;
    }
    try {
      const res = await fetch(`${BACKEND_URL}/api/verified/batch/${batchId}/delete`, {
        method: 'POST'
      });
      if (!res.ok) throw new Error("Gagal menghapus data");
      const data = await res.json();
      showToast(data.message || "Berita Acara berhasil dihapus");
      fetchStageData(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleTogglePublished = async (batchId) => {
    try {
      const res = await fetch(`${BACKEND_URL}/api/verified/batch/${batchId}/toggle-published`, {
        method: 'POST'
      });
      if (!res.ok) throw new Error("Gagal mengubah status");
      const data = await res.json();
      showToast(data.is_published ? "Berita Acara ditandai sudah terbit" : "Berita Acara ditandai belum terbit");
      fetchStageData(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const toggleBatchBreakdown = async (batchId) => {
    if (expandedBatchId === batchId) {
      setExpandedBatchId(null);
      return;
    }
    setExpandedBatchId(batchId);
    if (!batchBreakdownCache[batchId]) {
      try {
        const res = await fetch(`${BACKEND_URL}/api/batch/${batchId}/kabupaten-breakdown`);
        if (!res.ok) throw new Error("Gagal memuat data kabupaten");
        const data = await res.json();
        setBatchBreakdownCache(prev => ({ ...prev, [batchId]: data.breakdown }));
      } catch (err) {
        showToast(err.message, 'error');
      }
    }
  };

  const openPreview = async (batchId) => {
    setSelectedBatchIdForPreview(batchId);
    setPreviewTab('lolos');
    setPreviewSearchTerm('');
    setPreviewData(null);
    setShowPreviewModal(true);
    try {
      const res = await fetch(`${BACKEND_URL}/api/batch/${batchId}/export-preview`);
      if (!res.ok) throw new Error("Gagal memuat data preview");
      const data = await res.json();
      setPreviewData(data);
    } catch (err) {
      showToast(err.message, 'error');
      setShowPreviewModal(false);
    }
  };

  const closePreview = () => {
    setShowPreviewModal(false);
    setSelectedBatchIdForPreview(null);
    setPreviewData(null);
    setPreviewSearchTerm('');
  };

  const filterPreviewRecords = (records) => {
    if (!previewSearchTerm.trim()) return records;
    const term = previewSearchTerm.toLowerCase();
    return records.filter(r => {
      const searchFields = [
        r.nama, r.no_ktp, r.no_kk, r.alamat, r.desa_kelurahan,
        r.kecamatan, r.kabupaten_kota, r.jenis_kelamin, r.tahap, r.tanggal,
        r.alasan_tidak_lolos, r.keterangan,
        r.nama_pengganti, r.no_ktp_pengganti, r.no_kk_pengganti,
        r.alamat_pengganti, r.desa_pengganti, r.kec_pengganti, r.kab_pengganti
      ];
      return searchFields.some(f => f && f.toLowerCase().includes(term));
    });
  };

  // ============================================================
  // SK DIRJEN HANDLERS
  // ============================================================

  const fetchSkDirjenBatches = async () => {
    try {
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/batches`);
      const data = await res.json();
      setSkDirjenBatches(data.batches);
      if (data.batches.length > 0 && skDirjenSelectedBatch === 'all' && skDirjenRecords.length === 0) {
        fetchSkDirjenAllRecords();
      }
    } catch (err) {
      showToast("Gagal memuat batch SK Dirjen: " + err.message, 'error');
    }
  };

  const fetchSkDirjenRecords = async (batchId, filters = {}) => {
    try {
      const params = new URLSearchParams();
      if (filters.q) params.append('q', filters.q);
      if (filters.kabupaten) params.append('kabupaten', filters.kabupaten);
      if (filters.kecamatan) params.append('kecamatan', filters.kecamatan);
      if (filters.desa) params.append('desa', filters.desa);
      if (filters.tahap) params.append('tahap', filters.tahap);
      if (filters.asal_batch) params.append('asal_batch', filters.asal_batch);
      if (filters.status) params.append('status', filters.status);
      const qs = params.toString();
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/${batchId}/records${qs ? '?' + qs : ''}`);
      const data = await res.json();
      setSkDirjenRecords(data.records);
    } catch (err) {
      showToast("Gagal memuat data SK Dirjen: " + err.message, 'error');
    }
  };

  const fetchSkDirjenAllRecords = async (filters = {}) => {
    try {
      const params = new URLSearchParams();
      if (filters.q) params.append('q', filters.q);
      if (filters.kabupaten) params.append('kabupaten', filters.kabupaten);
      if (filters.kecamatan) params.append('kecamatan', filters.kecamatan);
      if (filters.desa) params.append('desa', filters.desa);
      if (filters.tahap) params.append('tahap', filters.tahap);
      if (filters.asal_batch) params.append('asal_batch', filters.asal_batch);
      if (filters.status) params.append('status', filters.status);
      const qs = params.toString();
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/all-records${qs ? '?' + qs : ''}`);
      const data = await res.json();
      setSkDirjenRecords(data.records);
    } catch (err) {
      showToast("Gagal memuat data SK Dirjen: " + err.message, 'error');
    }
  };

  const fetchSkDirjenWithFilters = (filters = {}) => {
    if (skDirjenSelectedBatch === 'all') {
      fetchSkDirjenAllRecords({ q: skDirjenDebouncedSearch, ...filters });
    } else if (skDirjenSelectedBatch) {
      fetchSkDirjenRecords(skDirjenSelectedBatch, { q: skDirjenDebouncedSearch, ...filters });
    }
  };

  const handleSkDirjenUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!skDirjenStageName.trim()) {
      showToast("Masukkan nama Tahap SK Dirjen terlebih dahulu", 'error');
      return;
    }
    setSkDirjenUploading(true);
    try {
      const formData = new FormData();
      formData.append('file', file);
      formData.append('stage_name', skDirjenStageName.trim());
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/upload`, { method: 'POST', body: formData });
      if (!res.ok) {
        const err = await res.json();
        throw new Error(err.detail || 'Upload gagal');
      }
      const data = await res.json();
      showToast(`Upload berhasil! ${data.inserted_records} data terproses dari ${data.total_rows} baris.`, 'success');
      setSkDirjenStageName('');
      e.target.value = '';
      fetchSkDirjenBatches();
    } catch (err) {
      showToast(err.message, 'error');
    } finally {
      setSkDirjenUploading(false);
    }
  };

  const handleSkDirjenApprove = async (recordId) => {
    try {
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/record/${recordId}/approve`, { method: 'POST' });
      if (!res.ok) throw new Error('Gagal menyetujui');
      showToast("Persetujuan berhasil disimpan", 'success');
      setSkDirjenApprovalRecord(null);
      if (skDirjenSelectedBatch) {
        const filters = { kabupaten: skDirjenFilterKab, kecamatan: skDirjenFilterKec, desa: skDirjenFilterDesa, tahap: skDirjenFilterTahap, asal_batch: skDirjenFilterAsalBatch, status: skDirjenFilterStatus };
        if (skDirjenSelectedBatch === 'all') {
          fetchSkDirjenAllRecords({ q: skDirjenDebouncedSearch, ...filters });
        } else {
          fetchSkDirjenRecords(skDirjenSelectedBatch, { q: skDirjenDebouncedSearch, ...filters });
        }
      }
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleSkDirjenReject = async (recordId) => {
    try {
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/record/${recordId}/reject`, { method: 'POST' });
      if (!res.ok) throw new Error('Gagal menolak');
      showToast("Penolakan berhasil disimpan", 'success');
      setSkDirjenApprovalRecord(null);
      if (skDirjenSelectedBatch) {
        const filters = { kabupaten: skDirjenFilterKab, kecamatan: skDirjenFilterKec, desa: skDirjenFilterDesa, tahap: skDirjenFilterTahap, asal_batch: skDirjenFilterAsalBatch, status: skDirjenFilterStatus };
        if (skDirjenSelectedBatch === 'all') {
          fetchSkDirjenAllRecords({ q: skDirjenDebouncedSearch, ...filters });
        } else {
          fetchSkDirjenRecords(skDirjenSelectedBatch, { q: skDirjenDebouncedSearch, ...filters });
        }
      }
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleSkDirjenDeleteBatch = async (batchId, batchName) => {
    if (!window.confirm(`Hapus tahap SK Dirjen "${batchName}" beserta SEMUA data terkait (records, matches, persetujuan, penolakan)?\n\nTindakan ini tidak dapat dibatalkan.`)) return;
    try {
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/batch/${batchId}`, { method: 'DELETE' });
      if (!res.ok) throw new Error('Gagal menghapus batch');
      showToast("Batch SK Dirjen berhasil dihapus", 'success');
      setSkDirjenSelectedBatch('all');
      setSkDirjenRecords([]);
      fetchSkDirjenBatches();
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleSkDirjenSearchVerified = async (term) => {
    setSkDirjenPairSearchTerm(term);
    if (term.length < 2) {
      setSkDirjenPairSearchResults([]);
      return;
    }
    try {
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/search-verified?q=${encodeURIComponent(term)}`);
      const data = await res.json();
      setSkDirjenPairSearchResults(data.records || []);
    } catch (err) {
      setSkDirjenPairSearchResults([]);
    }
  };

  const handleSkDirjenPair = async (skRecordId, verifiedRecordId) => {
    try {
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/pair/${skRecordId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ verified_record_id: verifiedRecordId })
      });
      if (!res.ok) throw new Error('Gagal memasangkan data');
      showToast("Berhasil memasangkan data!", 'success');
      setSkDirjenPairingRecord(null);
      setSkDirjenPairSearchTerm('');
      setSkDirjenPairSearchResults([]);
      if (skDirjenSelectedBatch) {
        const filters = { kabupaten: skDirjenFilterKab, kecamatan: skDirjenFilterKec, desa: skDirjenFilterDesa, tahap: skDirjenFilterTahap, asal_batch: skDirjenFilterAsalBatch, status: skDirjenFilterStatus };
        if (skDirjenSelectedBatch === 'all') {
          fetchSkDirjenAllRecords({ q: skDirjenDebouncedSearch, ...filters });
        } else {
          fetchSkDirjenRecords(skDirjenSelectedBatch, { q: skDirjenDebouncedSearch, ...filters });
        }
      }
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const fetchSkDirjenRekapPerTahap = async () => {
    try {
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/rekap-per-tahap`);
      const data = await res.json();
      setSkDirjenRekapPerTahap(data);
    } catch (err) {
      showToast("Gagal memuat rekap: " + err.message, 'error');
    }
  };

  const fetchSkDirjenRekapPerKab = async (batchId) => {
    try {
      const url = batchId === 'all'
        ? `${BACKEND_URL}/api/sk-dirjen/rekap-per-kabupaten/all`
        : `${BACKEND_URL}/api/sk-dirjen/rekap-per-kabupaten/${batchId}`;
      const res = await fetch(url);
      const data = await res.json();
      setSkDirjenRekapPerKab(data);
    } catch (err) {
      showToast("Gagal memuat rekap kabupaten: " + err.message, 'error');
    }
  };

  const fetchSumberSelisihDetail = async (kab, batchName, stageName, idx) => {
    if (expandedSumber?.kab === kab && expandedSumber?.idx === idx) {
      setExpandedSumber(null);
      return;
    }
    setExpandedSumber({ kab, idx });
    setSumberDetailLoading(true);
    try {
      const params = new URLSearchParams({ kabupaten: kab, batch_name: batchName, stage_name: stageName });
      const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/sumber-selisih-detail?${params}`);
      const data = await res.json();
      setSumberDetail(data.records || []);
    } catch (err) {
      setSumberDetail([]);
    } finally {
      setSumberDetailLoading(false);
    }
  };

  const handleExport = (batchId = null) => {
    if (!selectedStageId) return;
    const url = batchId 
      ? `${BACKEND_URL}/api/stage/${selectedStageId}/export?batch_id=${batchId}`
      : `${BACKEND_URL}/api/stage/${selectedStageId}/export`;
    window.open(url, '_blank');
  };

  const handleExportRekapKeseluruhan = () => {
    window.open(`${BACKEND_URL}/api/rekap-keseluruhan/export`, '_blank');
  };

  const handleExportRekapBatchBA = () => {
    window.open(`${BACKEND_URL}/api/rekap-batch-ba/export`, '_blank');
  };

  const handleExportFilteredInvers = () => {
    const data = getFilteredInvers();
    if (data.length === 0) { showToast("Tidak ada data untuk diekspor", "error"); return; }
    const rows = data.map(r => ({
      "No. Urut": r.no_urut,
      "Nama": r.nama,
      "NIK / No. KTP": r.no_ktp,
      "No. KK": r.no_kk,
      "Jenis Kelamin": r.jenis_kelamin,
      "Desa/Kelurahan": r.desa_kelurahan,
      "Kecamatan": r.kecamatan,
      "Kabupaten/Kota": r.kabupaten_kota,
      "Status Verifikasi": verifiedNikSet.has(r.no_ktp?.trim()) ? 'Terverifikasi' : 'Belum Terverifikasi'
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data INVERS");
    XLSX.writeFile(wb, `Data_INVERS_Terfilter_${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast(`Berhasil mengekspor ${rows.length} data INVERS`);
  };

  const handleExportFilteredVerified = () => {
    const data = getFilteredVerified();
    if (data.length === 0) { showToast("Tidak ada data untuk diekspor", "error"); return; }
    const rows = data.map(r => ({
      "Status": r.status,
      "Nama": r.nama,
      "NIK / No. KTP": r.no_ktp,
      "No. KK": r.no_kk,
      "Berita Acara": r.batch_name,
      "Kabupaten/Kota": r.kabupaten_kota || r.expected_invers?.kabupaten_kota || '-',
      "Desa/Kelurahan": r.desa_kelurahan || r.expected_invers?.desa_kelurahan || '-',
      "Hasil Analisis": r.has_error ? (r.override ? 'Rekonsiliasi Selesai' : 'Error (Butuh Rekonsiliasi)') : '-',
      "CPB Pengganti (Jika Ada)": r.override?.corrected_no_ktp || '-'
    }));
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data Terverifikasi");
    XLSX.writeFile(wb, `Data_Terverifikasi_Terfilter_${new Date().toISOString().slice(0,10)}.xlsx`);
    showToast(`Berhasil mengekspor ${rows.length} data terverifikasi`);
  };

  const handleReconcileOverride = async (record, type) => {
    const formData = new FormData();
    formData.append('stage_id', selectedStageId);
    formData.append('original_no_ktp', record.no_ktp);
    formData.append('override_type', type);

    if (type === 'MANUAL_EDIT') {
      formData.append('corrected_nama', editNama);
      formData.append('corrected_no_ktp', editKtp);
      formData.append('corrected_no_kk', editKk);
    }

    try {
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/override`, {
        method: 'POST',
        body: formData
      });
      if (!res.ok) throw new Error("Gagal menyimpan data perbaikan");
      showToast("Perbaikan rekonsiliasi berhasil diterapkan");
      setEditingRecord(null);
      fetchStageData(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleLinkToInvers = async (record, inversItem) => {
    const formData = new FormData();
    formData.append('stage_id', selectedStageId);
    formData.append('original_no_ktp', record.no_ktp);
    formData.append('override_type', 'MANUAL_EDIT');
    formData.append('corrected_nama', inversItem.nama);
    formData.append('corrected_no_ktp', inversItem.no_ktp);
    formData.append('corrected_no_kk', inversItem.no_kk);

    try {
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/override`, {
        method: 'POST',
        body: formData
      });
      if (!res.ok) throw new Error("Gagal memasangkan data");
      showToast("Berhasil memasangkan data dengan INVERS!");
      setShowLinkModal(false);
      setLinkingRecord(null);
      fetchStageData(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleDeleteVerifiedRecord = async (recordId, name) => {
    if (!window.confirm(`Apakah Anda yakin ingin menghapus data lapangan '${name}' dari database hasil verifikasi? Tindakan ini tidak dapat dibatalkan.`)) {
      return;
    }
    try {
      const res = await fetch(`${BACKEND_URL}/api/verified/record/${recordId}/delete`, {
        method: 'POST'
      });
      if (!res.ok) throw new Error("Gagal menghapus data lapangan");
      showToast("Data lapangan berhasil dihapus dari database!");
      fetchStageData(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  // Bulk Delete Handler
  const handleBulkDelete = async () => {
    if (selectedRecordIds.size === 0) {
      showToast("Tidak ada data yang dipilih untuk dihapus", "error");
      return;
    }
    if (!window.confirm(`Apakah Anda yakin ingin menghapus ${selectedRecordIds.size} data lapangan yang dipilih? Tindakan ini tidak dapat dibatalkan.`)) {
      return;
    }
    try {
      const formData = new FormData();
      formData.append('record_ids', Array.from(selectedRecordIds).join(','));
      const res = await fetch(`${BACKEND_URL}/api/verified/records/bulk-delete`, {
        method: 'POST',
        body: formData
      });
      if (!res.ok) throw new Error("Gagal menghapus data");
      const data = await res.json();
      showToast(data.message || `${selectedRecordIds.size} data berhasil dihapus`);
      setSelectedRecordIds(new Set());
      fetchStageData(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  // Toggle record selection for bulk delete
  const toggleRecordSelection = (recordId) => {
    setSelectedRecordIds(prev => {
      const next = new Set(prev);
      if (next.has(recordId)) {
        next.delete(recordId);
      } else {
        next.add(recordId);
      }
      return next;
    });
  };

  // Toggle all visible records
  const toggleAllRecordSelection = () => {
    const filtered = getFilteredVerified();
    if (selectedRecordIds.size === filtered.length) {
      setSelectedRecordIds(new Set());
    } else {
      setSelectedRecordIds(new Set(filtered.map(r => r.id)));
    }
  };

  // Bulk Reconcile Handler
  const handleBulkReconcile = async (overrideType) => {
    if (selectedMismatchNiks.size === 0) {
      showToast("Tidak ada data yang dipilih untuk direkonsiliasi", "error");
      return;
    }
    const label = overrideType === 'ACCEPT_VERIFIED' ? 'Menyetujui' : 'Menolak';
    if (!window.confirm(`Apakah Anda yakin ingin ${label} ${selectedMismatchNiks.size} data yang dipilih?`)) {
      return;
    }
    try {
      const formData = new FormData();
      formData.append('stage_id', selectedStageId);
      formData.append('niks', Array.from(selectedMismatchNiks).join(','));
      formData.append('override_type', overrideType);
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/bulk-override`, {
        method: 'POST',
        body: formData
      });
      if (!res.ok) throw new Error("Gagal melakukan rekonsiliasi bulk");
      const data = await res.json();
      showToast(data.message || `${selectedMismatchNiks.size} data berhasil direkonsiliasi`);
      setSelectedMismatchNiks(new Set());
      fetchStageData(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleBulkDeleteFromRecon = async () => {
    if (selectedMismatchNiks.size === 0) {
      showToast("Tidak ada data yang dipilih untuk dihapus", "error");
      return;
    }
    if (!window.confirm(`Apakah Anda yakin ingin menghapus ${selectedMismatchNiks.size} data lapangan yang dipilih dari database? Tindakan ini tidak dapat dibatalkan.`)) {
      return;
    }
    try {
      const formData = new FormData();
      formData.append('niks', Array.from(selectedMismatchNiks).join(','));
      if (selectedStageId) formData.append('stage_id', selectedStageId);
      const res = await fetch(`${BACKEND_URL}/api/verified/records/bulk-delete-by-nik`, {
        method: 'POST',
        body: formData
      });
      if (!res.ok) throw new Error("Gagal menghapus data");
      const data = await res.json();
      showToast(data.message || `${selectedMismatchNiks.size} data berhasil dihapus`);
      setSelectedMismatchNiks(new Set());
      fetchStageData(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const toggleMismatchSelection = (nik) => {
    setSelectedMismatchNiks(prev => {
      const next = new Set(prev);
      if (next.has(nik)) {
        next.delete(nik);
      } else {
        next.add(nik);
      }
      return next;
    });
  };

  // Toggle all visible mismatches
  const toggleAllMismatchSelection = () => {
    const filtered = getFilteredMismatches();
    if (selectedMismatchNiks.size === filtered.length) {
      setSelectedMismatchNiks(new Set());
    } else {
      setSelectedMismatchNiks(new Set(filtered.map(r => r.no_ktp)));
    }
  };

  // --- Invers Manual Pairs ---
  const fetchUnmatchedInvers = async (stageId) => {
    if (!stageId) return;
    try {
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/unmatched-invers/${stageId}`);
      const data = await res.json();
      setUnmatchedInvers(data.records || []);
    } catch (err) {
      showToast("Gagal memuat data belum terverifikasi: " + err.message, 'error');
    }
  };

  const fetchManualPairs = async (stageId) => {
    if (!stageId) return;
    try {
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/manual-pairs/${stageId}`);
      const data = await res.json();
      setManualPairs(data.pairs || []);
    } catch (err) {
      showToast("Gagal memuat data pasangan: " + err.message, 'error');
    }
  };

  const fetchUnmatchedVerified = async (stageId) => {
    if (!stageId) return;
    try {
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/unmatched-verified/${stageId}`);
      const data = await res.json();
      setUnmatchedVerified(data.records || []);
    } catch (err) {
      showToast("Gagal memuat data terverifikasi belum dipasangkan: " + err.message, 'error');
    }
  };

  const handleOpenPairModal = (inversRecord) => {
    setPairingInvers(inversRecord);
    const nama = inversRecord.nama || '';
    setPairSearchTerm(nama);
    setPairSearchResults([]);
    setShowPairModal(true);
    if (nama.length >= 2) handleSearchForPairing(nama, inversRecord.desa_kelurahan);
  };

  const handleAutoPairNIK = async () => {
    if (!selectedStageId) return;
    setAutoPairing(true);
    try {
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/auto-pair-nik/${selectedStageId}`, { method: 'POST' });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Gagal auto-pair');
      showToast(`Berhasil auto-pair ${data.paired_count} data. ${data.no_match} data belum cocok.`, 'success');
      fetchUnmatchedInvers(selectedStageId);
      fetchManualPairs(selectedStageId);
      fetchUnmatchedVerified(selectedStageId);
      fetchSuggestedPairs(selectedStageId);
    } catch (err) {
      showToast("Gagal auto-pair: " + err.message, 'error');
    } finally {
      setAutoPairing(false);
    }
  };

  const fetchSuggestedPairs = async (stageId) => {
    if (!stageId) return;
    try {
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/suggest-pairs/${stageId}`);
      const data = await res.json();
      setSuggestedPairs(data.suggestions || []);
      setSelectedSuggestions(new Set((data.suggestions || []).map((_, i) => i)));
    } catch (err) {
      setSuggestedPairs([]);
    }
  };

  const handleBatchPair = async () => {
    if (!selectedStageId || selectedSuggestions.size === 0) return;
    setBatchPairing(true);
    try {
      const pairs = suggestedPairs
        .filter((_, i) => selectedSuggestions.has(i))
        .map(s => ({
          invers_nik: s.invers.no_ktp,
          invers_nama: s.invers.nama,
          invers_kabupaten: s.invers.kabupaten_kota || '',
          verified_record_id: s.verified.id
        }));
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/batch-pair`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ stage_id: selectedStageId, pairs })
      });
      const data = await res.json();
      if (!res.ok) throw new Error(data.detail || 'Gagal batch pair');
      showToast(data.message || `Berhasil memasangkan ${data.paired_count} data`);
      setSuggestedPairs([]);
      setSelectedSuggestions(new Set());
      fetchUnmatchedInvers(selectedStageId);
      fetchManualPairs(selectedStageId);
      fetchUnmatchedVerified(selectedStageId);
    } catch (err) {
      showToast("Gagal batch pair: " + err.message, 'error');
    } finally {
      setBatchPairing(false);
    }
  };

  const fetchInlineSuggestions = async (records) => {
    if (!records || records.length === 0) return;
    const newSuggestions = {};
    for (const ir of records) {
      try {
        const res = await fetch(`${BACKEND_URL}/api/sk-dirjen/search-verified?q=${encodeURIComponent(ir.nama)}&desa=${encodeURIComponent(ir.desa_kelurahan || '')}`);
        const data = await res.json();
        if (data.records && data.records.length > 0) {
          newSuggestions[ir.no_ktp] = data.records[0];
        }
      } catch (err) { /* skip */ }
    }
    setInlineSuggestions(prev => ({ ...prev, ...newSuggestions }));
  };

  const handleInlinePair = async (ir, verifiedRecord) => {
    if (!selectedStageId) return;
    try {
      const formData = new FormData();
      formData.append('stage_id', selectedStageId);
      formData.append('invers_nik', ir.no_ktp);
      formData.append('invers_nama', ir.nama);
      formData.append('invers_kabupaten', ir.kabupaten_kota || '');
      formData.append('verified_record_id', verifiedRecord.id);
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/pair-invers`, { method: 'POST', body: formData });
      if (!res.ok) throw new Error('Gagal memasangkan data');
      const data = await res.json();
      showToast(data.message || 'Berhasil memasangkan data');
      fetchUnmatchedInvers(selectedStageId);
      fetchManualPairs(selectedStageId);
      fetchUnmatchedVerified(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleSearchForPairing = async (term, desaFilter = null) => {
    setPairSearchTerm(term);
    if (term.length < 2) { setPairSearchResults([]); return; }
    try {
      let url = `${BACKEND_URL}/api/sk-dirjen/search-verified?q=${encodeURIComponent(term)}`;
      if (desaFilter && desaFilter.trim()) {
        url += `&desa=${encodeURIComponent(desaFilter)}`;
      }
      const res = await fetch(url);
      const data = await res.json();
      setPairSearchResults(data.records || []);
    } catch (err) {
      showToast("Gagal mencari data: " + err.message, 'error');
    }
  };

  const handlePairInvers = async (verifiedRecordId) => {
    if (!pairingInvers || !selectedStageId) return;
    try {
      const formData = new FormData();
      formData.append('stage_id', selectedStageId);
      formData.append('invers_nik', pairingInvers.no_ktp);
      formData.append('invers_nama', pairingInvers.nama);
      formData.append('invers_kabupaten', pairingInvers.kabupaten_kota || '');
      formData.append('verified_record_id', verifiedRecordId);
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/pair-invers`, {
        method: 'POST', body: formData
      });
      if (!res.ok) throw new Error('Gagal memasangkan data');
      const data = await res.json();
      showToast(data.message || 'Berhasil memasangkan data');
      setShowPairModal(false);
      setPairingInvers(null);
      fetchUnmatchedInvers(selectedStageId);
      fetchManualPairs(selectedStageId);
      fetchUnmatchedVerified(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleUnpairInvers = async (pairId) => {
    if (!window.confirm('Yakin ingin menghapus pasangan ini?')) return;
    try {
      const res = await fetch(`${BACKEND_URL}/api/reconciliation/unpair-invers/${pairId}`, { method: 'DELETE' });
      if (!res.ok) throw new Error('Gagal menghapus pasangan');
      showToast('Pasangan berhasil dihapus');
      fetchUnmatchedInvers(selectedStageId);
      fetchManualPairs(selectedStageId);
      fetchUnmatchedVerified(selectedStageId);
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleExportPengusulTree = () => {
    if (!selectedStageId) return;
    window.open(`${BACKEND_URL}/api/stage/${selectedStageId}/pengusul-tree/export`, '_blank');
  };

  const handleStartManualEdit = (record) => {
    setEditingRecord(record);
    setEditNama(record.nama);
    setEditKtp(record.no_ktp);
    setEditKk(record.no_kk);
  };

  const handleClearDatabase = async () => {
    if (!window.confirm("PERINGATAN! Ini akan menghapus seluruh data tahap, invers, verifikasi, dan perbaikan rekonsiliasi di database. Apakah Anda yakin?")) {
      return;
    }
    try {
      const res = await fetch(`${BACKEND_URL}/api/database/clear`, {
        method: 'POST'
      });
      if (!res.ok) throw new Error("Gagal mengosongkan database");
      showToast("Database berhasil disetel ulang (bersih)!");
      setSelectedStageId('');
      setStages([]);
      setStageSummary(null);
      setRecordsData(null);
      setOverviewStats(null);
      setOverviewTables(null);
      setPengusulTree([]);
      setKabPengusulTree([]);
      await fetchStages();
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleDeleteStage = async (stageId, stageName) => {
    if (!window.confirm(`PERINGATAN! Anda akan menghapus tahap "${stageName}" beserta SEMUA data terkait (INVERS, Berita Acara, Verifikasi, Rekonsiliasi). Tindakan ini tidak dapat dibatalkan.\n\nApakah Anda yakin?`)) {
      return;
    }
    try {
      const res = await fetch(`${BACKEND_URL}/api/stage/${stageId}/delete`, {
        method: 'POST'
      });
      if (!res.ok) {
        const errData = await res.json();
        throw new Error(errData.detail || "Gagal menghapus tahap");
      }
      showToast(`Tahap "${stageName}" berhasil dihapus beserta semua data terkait`);
      
      // Reset selection jika tahap yang dihapus adalah yang sedang aktif
      if (selectedStageId === stageId.toString()) {
        setSelectedStageId('');
        setStageSummary(null);
        setRecordsData(null);
        setOverviewStats(null);
        setOverviewTables(null);
        setPengusulTree([]);
      }
      
      await fetchStages();
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handleWordExportSubmit = async (e) => {
    e.preventDefault();
    try {
      const formData = new FormData();
      formData.append('stage_id', selectedStageId);
      if (selectedBatchIdForWord) {
        formData.append('batch_id', selectedBatchIdForWord);
      }
      formData.append('nomor_ba', wordFormData.nomor_ba);
      formData.append('nomor_surat', wordFormData.nomor_surat);
      formData.append('tanggal_ba', wordFormData.tanggal_ba);
      formData.append('lokasi_ba', wordFormData.lokasi_ba);
      formData.append('no_surat_dirjen', wordFormData.no_surat_dirjen);
      formData.append('tgl_surat_dirjen', wordFormData.tgl_surat_dirjen);
      formData.append('hal_surat_dirjen', wordFormData.hal_surat_dirjen);

      showToast("Sedang menyusun surat Word...", "info");
      const res = await fetch(`${BACKEND_URL}/api/export/docx`, {
        method: 'POST',
        body: formData
      });
      if (!res.ok) throw new Error("Gagal mengekspor berkas Word");

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      
      const batchName = stageSummary?.batches?.find(b => b.id === selectedBatchIdForWord)?.name || '';
      const nameSuffix = (batchName || selectedStageName).replace(/\s+/g, '_');
      
      a.download = `DRAFT_SURAT_BA_${nameSuffix}.zip`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      
      setShowWordModal(false);
      setSelectedBatchIdForWord(null);
      showToast("Berkas BA & Surat Penyampaian berhasil diunduh!");
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const handlePdfExportSubmit = async (e) => {
    e.preventDefault();
    try {
      const formData = new FormData();
      formData.append('stage_id', selectedStageId);
      if (selectedBatchIdForWord) {
        formData.append('batch_id', selectedBatchIdForWord);
      }
      formData.append('nomor_ba', wordFormData.nomor_ba);
      formData.append('nomor_surat', wordFormData.nomor_surat);
      formData.append('tanggal_ba', wordFormData.tanggal_ba);
      formData.append('lokasi_ba', wordFormData.lokasi_ba);
      formData.append('no_surat_dirjen', wordFormData.no_surat_dirjen);
      formData.append('tgl_surat_dirjen', wordFormData.tgl_surat_dirjen);
      formData.append('hal_surat_dirjen', wordFormData.hal_surat_dirjen);

      showToast("Sedang menyusun surat PDF...", "info");
      const res = await fetch(`${BACKEND_URL}/api/export/pdf`, {
        method: 'POST',
        body: formData
      });
      if (!res.ok) throw new Error("Gagal mengekspor berkas PDF");

      const blob = await res.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      
      const batchName = stageSummary?.batches?.find(b => b.id === selectedBatchIdForWord)?.name || '';
      const nameSuffix = (batchName || selectedStageName).replace(/\s+/g, '_');
      
      a.download = `DRAFT_SURAT_BA_${nameSuffix}.zip`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
      
      setShowWordModal(false);
      setSelectedBatchIdForWord(null);
      showToast("Berkas BA PDF & Surat Penyampaian berhasil diunduh!");
    } catch (err) {
      showToast(err.message, 'error');
    }
  };

  const selectedStageName = stages.find(s => s.id.toString() === selectedStageId)?.name || 'Tidak ada';

  // Filter pencarian data
  const getFilteredInvers = () => {
    if (!recordsData || !recordsData.invers_records) return [];
    const verifiedNiks = new Set(
      (recordsData.verified_records || []).map(v => v.no_ktp?.trim())
    );
    return recordsData.invers_records.filter(r => {
      const matchSearch = r.nama.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.no_ktp.includes(searchTerm) ||
        r.no_kk.includes(searchTerm);
      const matchKab = inversKabFilter === 'ALL' || r.kabupaten_kota === inversKabFilter;
      const matchDesa = inversDesaFilter === 'ALL' || r.desa_kelurahan === inversDesaFilter;
      const isVerified = verifiedNiks.has(r.no_ktp?.trim());
      const matchStatus = inversStatusFilter === 'ALL' ||
        (inversStatusFilter === 'VERIFIED' && isVerified) ||
        (inversStatusFilter === 'NOT_VERIFIED' && !isVerified);
      return matchSearch && matchKab && matchDesa && matchStatus;
    });
  };

  const getFilteredVerified = () => {
    if (!recordsData || !recordsData.verified_records) return [];
    return recordsData.verified_records.filter(r => {
      const matchSearch = r.nama.toLowerCase().includes(searchTerm.toLowerCase()) ||
        r.no_ktp.includes(searchTerm) ||
        r.no_kk.includes(searchTerm);
      const kabVal = r.kabupaten_kota || r.expected_invers?.kabupaten_kota || '';
      const desaVal = r.desa_kelurahan || r.expected_invers?.desa_kelurahan || '';
      const matchKab = verifiedKabFilter === 'ALL' || kabVal === verifiedKabFilter;
      const matchDesa = verifiedDesaFilter === 'ALL' || desaVal === verifiedDesaFilter;
      const matchBatch = verifiedBatchFilter === 'ALL' || r.batch_name === verifiedBatchFilter;
      let matchHasil = true;
      if (verifiedHasilFilter !== 'ALL') {
        if (verifiedHasilFilter === 'LOLOS') matchHasil = r.status === 'LOLOS';
        else if (verifiedHasilFilter === 'TIDAK_LOLOS') matchHasil = r.status === 'TIDAK LOLOS';
        else if (verifiedHasilFilter === 'ERROR') matchHasil = r.has_error && !r.override;
        else if (verifiedHasilFilter === 'SELESAI') matchHasil = r.has_error && r.override;
      }
      return matchSearch && matchKab && matchDesa && matchHasil && matchBatch;
    });
  };

  const getMismatches = () => {
    if (!recordsData || !recordsData.verified_records) return [];
    return recordsData.verified_records.filter(r => r.has_error && !r.override);
  };

  // Saring data berdasarkan filter kategori error di tab Rekonsiliasi
  const getFilteredMismatches = () => {
    const mismatches = getMismatches();
    if (errorFilter === 'ALL') return mismatches;
    return mismatches.filter(r => r.mismatch_type === errorFilter);
  };

  // Hitung jumlah sisa data mismatch (yang belum di-override) untuk badge menu
  const getActiveMismatchCount = () => {
    if (!recordsData || !recordsData.verified_records) return 0;
    return recordsData.verified_records.filter(r => r.has_error && !r.override).length;
  };

  // Progress Bar segments calculation
  const greenCount = overviewStats?.segments?.green || 0;
  const yellowCount = overviewStats?.segments?.yellow || 0;
  const orangeCount = overviewStats?.segments?.orange || 0;
  const redCount = overviewStats?.segments?.red || 0;
  
  const totalSegmentsSum = greenCount + yellowCount + orangeCount + redCount;
  const getPercentStyle = (val) => {
    if (totalSegmentsSum === 0) return '0%';
    return `${(val / totalSegmentsSum * 100).toFixed(1)}%`;
  };

  // Ambil list unik Pengusul untuk dropdown filter pohon
  const getUniquePengusuls = () => {
    if (!pengusulTree) return [];
    return pengusulTree.map(p => p.name);
  };

  const getFilteredPengusulTree = () => {
    if (selectedPengusulFilter === 'ALL') return pengusulTree;
    return pengusulTree.filter(p => p.name === selectedPengusulFilter);
  };

  // Set NIK yang sudah terverifikasi (untuk kolom Status di tabel INVERS)
  const verifiedNikSet = new Set(
    (recordsData?.verified_records || []).map(v => v.no_ktp?.trim())
  );

  // Unique values untuk filter dropdown INVERS
  const uniqueInversKab = [...new Set((recordsData?.invers_records || []).map(r => r.kabupaten_kota).filter(Boolean))].sort();
  const uniqueInversDesa = [...new Set((recordsData?.invers_records || []).map(r => r.desa_kelurahan).filter(Boolean))].sort();

  // Unique values untuk filter dropdown Verified
  const uniqueVerifiedKab = [...new Set((recordsData?.verified_records || []).map(r => r.kabupaten_kota || r.expected_invers?.kabupaten_kota).filter(Boolean))].sort();
  const uniqueVerifiedDesa = [...new Set((recordsData?.verified_records || []).map(r => r.desa_kelurahan || r.expected_invers?.desa_kelurahan).filter(Boolean))].sort();
  const uniqueVerifiedBatch = [...new Set((recordsData?.verified_records || []).map(r => r.batch_name).filter(Boolean))].sort();

  const filteredInvers = getFilteredInvers();
  const totalInversPages = Math.ceil(filteredInvers.length / ITEMS_PER_PAGE) || 1;
  const paginatedInvers = filteredInvers.slice((inversPage - 1) * ITEMS_PER_PAGE, inversPage * ITEMS_PER_PAGE);

  const filteredVerified = getFilteredVerified();
  const totalVerifiedPages = Math.ceil(filteredVerified.length / ITEMS_PER_PAGE) || 1;
  const paginatedVerified = filteredVerified.slice((verifiedPage - 1) * ITEMS_PER_PAGE, verifiedPage * ITEMS_PER_PAGE);

  return (
    <div className="app-container">
      {/* Notifikasi Toast */}
      {toast && (
        <div className={`toast ${toast.type}`}>
          {toast.message}
        </div>
      )}

      {/* Word Export Modal Dialog */}
      {showWordModal && (
        <div className="modal-overlay">
          <form className="modal-content" onSubmit={handleWordExportSubmit}>
            <div className="modal-header">
              <h3><svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "8px", verticalAlign: "middle" }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>Buat Dokumen Berita Acara & Surat Penyampaian</h3>
              <button 
                type="button" 
                className="modal-close"
                onClick={() => {
                  setShowWordModal(false);
                  setSelectedBatchIdForWord(null);
                }}
              >
                &times;
              </button>
            </div>
            
            <div className="modal-body">
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '12px', padding: '8px 12px', backgroundColor: '#f8f9fa', borderRadius: '6px', border: '1px solid var(--border)' }}>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', margin: 0 }}>
                  Isi form berikut untuk mengganti kolom placeholder <code style={{color: 'var(--primary)'}}>[...]</code> yang ada di file template Word.
                </p>
                <div style={{ display: 'flex', gap: '6px', flexShrink: 0, marginLeft: '12px' }}>
                  <button 
                    type="button" 
                    className="btn btn-secondary btn-sm" 
                    onClick={handleSaveTemplate}
                    style={{ fontSize: '0.75rem', padding: '4px 10px', display: 'flex', alignItems: 'center', gap: '4px' }}
                    title="Simpan isi form sebagai template untuk batch ini"
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"></path><polyline points="17 21 17 13 7 13 7 21"></polyline><polyline points="7 3 7 8 15 8"></polyline></svg>
                    Simpan
                  </button>
                  <button 
                    type="button" 
                    className="btn btn-secondary btn-sm" 
                    onClick={handleLoadTemplate}
                    style={{ fontSize: '0.75rem', padding: '4px 10px', display: 'flex', alignItems: 'center', gap: '4px' }}
                    title="Muat template yang tersimpan untuk batch ini"
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"></path></svg>
                    Muat
                  </button>
                  {getSavedTemplate() && (
                    <button 
                      type="button" 
                      className="btn btn-danger btn-sm" 
                      onClick={handleDeleteTemplate}
                      style={{ fontSize: '0.75rem', padding: '4px 10px', display: 'flex', alignItems: 'center', gap: '4px' }}
                      title="Hapus template yang tersimpan"
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
                      Hapus
                    </button>
                  )}
                </div>
              </div>
              
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
                onClick={() => {
                  setShowWordModal(false);
                  setSelectedBatchIdForWord(null);
                }}
              >
                Batal
              </button>
              <button type="button" className="btn btn-secondary" onClick={handlePdfExportSubmit}>
                <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg>Buat PDF
              </button>
              <button type="submit" className="btn btn-primary">
                <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>Buat ZIP Dokumen
              </button>
            </div>
          </form>
        </div>
      )}

      {/* Pasangkan Manual ke INVERS Modal */}
      {showLinkModal && linkingRecord && (
        <div className="modal-overlay">
          <div className="modal-content" style={{ width: '700px', maxWidth: '90%' }}>
            <div className="modal-header">
              <h3><svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "8px", verticalAlign: "middle" }}><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg>Pasangkan dengan Data INVERS Rujukan</h3>
              <button 
                type="button" 
                className="modal-close"
                onClick={() => {
                  setShowLinkModal(false);
                  setLinkingRecord(null);
                }}
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
                    {recordsData?.invers_records
                      ?.filter(ir => {
                        const term = linkSearchTerm.toLowerCase();
                        return ir.nama.toLowerCase().includes(term) ||
                          ir.no_ktp.includes(term) ||
                          ir.no_kk.includes(term) ||
                          (ir.desa_kelurahan || '').toLowerCase().includes(term);
                      })
                      .slice(0, 15)
                      .map(ir => (
                        <tr key={ir.id}>
                          <td style={{ fontWeight: '600' }}>{ir.nama}</td>
                          <td className="mono-digit">{ir.no_ktp}</td>
                          <td className="mono-digit">{ir.no_kk}</td>
                          <td>{ir.desa_kelurahan}</td>
                          <td>
                            <button 
                              className="btn btn-primary btn-sm"
                              onClick={() => handleLinkToInvers(linkingRecord, ir)}
                            >
                              Pilih & Pasangkan
                            </button>
                          </td>
                        </tr>
                      ))
                    }
                    {(!recordsData?.invers_records || recordsData.invers_records.filter(ir => {
                      const term = linkSearchTerm.toLowerCase();
                      return ir.nama.toLowerCase().includes(term) ||
                        ir.no_ktp.includes(term) ||
                        ir.no_kk.includes(term) ||
                        (ir.desa_kelurahan || '').toLowerCase().includes(term);
                    }).length === 0) && (
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
                onClick={() => {
                  setShowLinkModal(false);
                  setLinkingRecord(null);
                }}
              >
                Batal
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Sidebar Navigasi */}
      <aside className="sidebar">
        <div className="logo-container">
          <img src={logopkp} alt="Logo PKP" className="logo-img" />
          <span className="logo-text">SIVERI BSPS</span>
        </div>

        <div className="form-group" style={{ marginBottom: '24px' }}>
          <label className="form-label" style={{ color: '#a7f3d0', fontSize: '0.8rem', fontWeight: '600' }}>Tahap INVERS Aktif</label>
          <select 
            className="form-input" 
            value={selectedStageId} 
            onChange={handleStageSelect}
            style={{ padding: '8px 12px', fontSize: '0.85rem', background: 'rgba(255, 255, 255, 0.15)', color: '#ffffff', border: '1px solid rgba(255, 255, 255, 0.25)' }}
          >
            {getSortedStages().map(s => (
              <option key={s.id} value={s.id}>{s.name} ({s.record_count} CPB)</option>
            ))}
            {stages.length === 0 && <option value="">Belum ada tahap</option>}
          </select>
        </div>

        <nav>
          <div className="menu-label">Menu Utama</div>
          <ul className="menu-list">
            <li 
              className={`menu-item ${activeTab === 'dashboard' ? 'active' : ''}`}
              onClick={() => setActiveTab('dashboard')}
            >
              <IconDashboard /> Dashboard
            </li>
            <li 
              className={`menu-item ${activeTab === 'invers' ? 'active' : ''}`}
              onClick={() => setActiveTab('invers')}
            >
              <IconInvers /> Data INVERS
            </li>
            <li 
              className={`menu-item ${activeTab === 'verified' ? 'active' : ''}`}
              onClick={() => setActiveTab('verified')}
            >
              <IconVerified /> Data Terverifikasi
            </li>
            <li 
              className={`menu-item has-submenu ${activeTab === 'sk-dirjen' ? 'active' : ''}`}
            >
              <div style={{ display: 'flex', alignItems: 'center', width: '100%' }}>
                <div onClick={() => {
                  setActiveTab('sk-dirjen');
                  if (!skDirjenActiveSubTab) setSkDirjenActiveSubTab('daftar-pb');
                }} style={{ flex: 1, display: 'flex', alignItems: 'center', cursor: 'pointer' }}>
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "8px", verticalAlign: "middle" }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>
                  SK Dirjen
                </div>
                <span 
                  onClick={(e) => { e.stopPropagation(); setSkDirjenSubmenuOpen(!skDirjenSubmenuOpen); }}
                  style={{ cursor: 'pointer', padding: '4px', transition: 'transform 0.2s', transform: skDirjenSubmenuOpen ? 'rotate(90deg)' : 'rotate(0deg)', display: 'flex', alignItems: 'center' }}
                >
                  <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="9 18 15 12 9 6"></polyline></svg>
                </span>
              </div>
              {skDirjenSubmenuOpen && (
                <ul className="submenu">
                  <li className={`submenu-item ${skDirjenActiveSubTab === 'daftar-pb' ? 'active' : ''}`}
                    onClick={() => { setSkDirjenActiveSubTab('daftar-pb'); setActiveTab('sk-dirjen'); fetchSkDirjenBatches(); }}>Daftar PB</li>
                  <li className={`submenu-item ${skDirjenActiveSubTab === 'rekap-per-tahap' ? 'active' : ''}`}
                    onClick={() => { setSkDirjenActiveSubTab('rekap-per-tahap'); setActiveTab('sk-dirjen'); fetchSkDirjenRekapPerTahap(); }}>Rekap PB Per Tahap</li>
                  <li className={`submenu-item ${skDirjenActiveSubTab === 'rekap-per-kabupaten' ? 'active' : ''}`}
                    onClick={() => { setSkDirjenActiveSubTab('rekap-per-kabupaten'); setActiveTab('sk-dirjen'); fetchSkDirjenBatches(); setSkDirjenSelectedKabForRekap('all'); fetchSkDirjenRekapPerKab('all'); }}>Rekap PB Per Kabupaten</li>
                </ul>
              )}
            </li>
            <li 
              className={`menu-item ${activeTab === 'overview' ? 'active' : ''}`}
              onClick={() => setActiveTab('overview')}
            >
              <IconOverview /> Pusat Rekap Data
            </li>
            <li 
              className={`menu-item ${activeTab === 'reconciliation' ? 'active' : ''}`}
              onClick={() => setActiveTab('reconciliation')}
            >
              <IconReconcile /> Rekonsiliasi
              {getActiveMismatchCount() > 0 && (
                <span className="menu-badge">{getActiveMismatchCount()}</span>
              )}
            </li>
            <li 
              className={`menu-item ${activeTab === 'rekap' ? 'active' : ''}`}
              onClick={() => { setActiveTab('rekap'); fetchRekapKeseluruhan(); }}
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M16 6l2.29 2.29-4.88 4.88-4-4L2 16.59 3.41 18l6-6 4 4 6.3-6.29L22 12V6h-6z"/></svg>
              Rekap Keseluruhan
            </li>
            <li 
              className={`menu-item ${activeTab === 'rekap-batch-ba' ? 'active' : ''}`}
              onClick={() => { setActiveTab('rekap-batch-ba'); fetchRekapBatchBA(); }}
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/></svg>
              Rekap Batch BA
            </li>
            <li 
              className={`menu-item ${activeTab === 'global-search' ? 'active' : ''}`}
              onClick={() => { setActiveTab('global-search'); fetchGlobalSearch(); }}
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" width="16" height="16"><path d="M15.5 14h-.79l-.28-.27A6.471 6.471 0 0 0 16 9.5 6.5 6.5 0 1 0 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>
              Pencarian Global
            </li>
            <li 
              className={`menu-item ${activeTab === 'rekap-unggahan' ? 'active' : ''}`}
              onClick={() => { setActiveTab('rekap-unggahan'); fetchRekapUnggahan(); }}
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>
              Rekap Unggahan
            </li>
            <li 
              className={`menu-item ${activeTab === 'upload' ? 'active' : ''}`}
              onClick={() => setActiveTab('upload')}
            >
              <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24"><path d="M19.35 10.04C18.67 6.59 15.64 4 12 4 9.11 4 6.6 5.64 5.35 8.04 2.34 8.36 0 10.91 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96zM14 13v4h-4v-4H7l5-5 5 5h-3z"/></svg>
              Unggah Data
            </li>
            <li 
              className={`menu-item ${activeTab === 'settings' ? 'active' : ''}`}
              onClick={() => setActiveTab('settings')}
            >
              <IconSettings /> Pengaturan
            </li>
            <li 
              className={`menu-item ${activeTab === 'help' ? 'active' : ''}`}
              onClick={() => setActiveTab('help')}
            >
              <IconHelp /> Bantuan
            </li>
          </ul>
        </nav>

        <div style={{ marginTop: 'auto', padding: '8px 0' }}>
          <button
            onClick={() => setDarkMode(!darkMode)}
            style={{
              display: 'flex', alignItems: 'center', gap: '8px', width: '100%',
              padding: '10px 12px', border: 'none', borderRadius: '8px', cursor: 'pointer',
              background: 'rgba(255, 255, 255, 0.15)',
              color: '#ffffff',
              fontSize: '0.85rem', fontWeight: '500', transition: 'var(--transition)'
            }}
            title={darkMode ? 'Mode Terang' : 'Mode Gelap'}
          >
            {darkMode ? (
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="12" cy="12" r="5"></circle><line x1="12" y1="1" x2="12" y2="3"></line><line x1="12" y1="21" x2="12" y2="23"></line><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"></line><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"></line><line x1="1" y1="12" x2="3" y2="12"></line><line x1="21" y1="12" x2="23" y2="12"></line><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"></line><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"></line></svg>
            ) : (
              <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"></path></svg>
            )}
            {darkMode ? 'Mode Terang' : 'Mode Gelap'}
          </button>
          <div style={{ marginTop: '10px', fontSize: '0.72rem', color: '#a7f3d0', textAlign: 'center', fontWeight: '500' }}>
            Kementerian PUPR - SIVERI BSPS v1.1.0
          </div>
        </div>
      </aside>

      {/* Panel Utama */}
      <main className="main-content">
        
        {/* Header Bar */}
        <header className="header-bar">
          <div className="header-title">
            <h1 style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
              {selectedStageName}
              {selectedStageId && (
                <button
                  className="btn btn-secondary btn-sm"
                  style={{ padding: '2px 8px', fontSize: '0.72rem', fontWeight: 600, display: 'inline-flex', alignItems: 'center', gap: '4px', verticalAlign: 'middle' }}
                  onClick={() => openRenameStageModal(selectedStageId, selectedStageName)}
                  title="Ubah Nama Tahap INVERS"
                >
                  <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>
                  Ubah Nama Tahap
                </button>
              )}
            </h1>
            <p>Sistem Database Verifikasi Perumahan Swadaya</p>
          </div>

          <div className="search-wrapper">
            <span className="search-icon"><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg></span>
            <input 
              type="text" 
              placeholder="Cari berdasarkan Nama, NIK, KK..." 
              className="search-input"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
          </div>
        </header>

        {/* Dashboard View */}
        {activeTab === 'dashboard' && (
          <>
            {/* Visual Health Progress Bar Section */}
            {totalSegmentsSum > 0 && (
              <div className="progress-section">
                <div className="progress-header">
                  <span className="progress-header-title"><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><line x1="18" y1="20" x2="18" y2="10"></line><line x1="12" y1="20" x2="12" y2="4"></line><line x1="6" y1="20" x2="6" y2="14"></line></svg>Metrik Kualitas & Keberhasilan Verifikasi Tahap</span>
                  <span style={{ fontSize: '0.85rem', fontWeight: 600 }}>Total Data Terproses: {totalSegmentsSum} CPB</span>
                </div>
                
                <div className="progress-bar-multi">
                  {greenCount > 0 && (
                    <div 
                      className="progress-segment green" 
                      style={{ width: getPercentStyle(greenCount) }}
                      title={`Lolos & Cocok: ${greenCount} CPB`}
                    >
                      {getPercentStyle(greenCount)}
                    </div>
                  )}
                  {yellowCount > 0 && (
                    <div 
                      className="progress-segment yellow" 
                      style={{ width: getPercentStyle(yellowCount) }}
                      title={`Butuh Rekonsiliasi NIK/KK: ${yellowCount} CPB`}
                    >
                      {getPercentStyle(yellowCount)}
                    </div>
                  )}
                  {orangeCount > 0 && (
                    <div 
                      className="progress-segment orange" 
                      style={{ width: getPercentStyle(orangeCount) }}
                      title={`Tidak Terdeteksi di INVERS: ${orangeCount} CPB`}
                    >
                      {getPercentStyle(orangeCount)}
                    </div>
                  )}
                  {redCount > 0 && (
                    <div 
                      className="progress-segment red" 
                      style={{ width: getPercentStyle(redCount) }}
                      title={`Belum Terverifikasi Lapangan: ${redCount} CPB`}
                    >
                      {getPercentStyle(redCount)}
                    </div>
                  )}
                </div>

                <div className="progress-legend">
                  <div className="legend-item">
                    <div className="legend-color green"></div>
                    <div className="legend-info">
                      <span className="legend-name">Selesai & Lolos (Hijau)</span>
                      <span className="legend-count">{greenCount} CPB ({getPercentStyle(greenCount)})</span>
                    </div>
                  </div>
                  <div className="legend-item">
                    <div className="legend-color yellow"></div>
                    <div className="legend-info">
                      <span className="legend-name">Perlu Rekonsiliasi (Kuning)</span>
                      <span className="legend-count">{yellowCount} CPB ({getPercentStyle(yellowCount)})</span>
                    </div>
                  </div>
                  <div className="legend-item">
                    <div className="legend-color orange"></div>
                    <div className="legend-info">
                      <span className="legend-name">Tidak Terdeteksi di INVERS (Oranye)</span>
                      <span className="legend-count">{orangeCount} CPB ({getPercentStyle(orangeCount)})</span>
                    </div>
                  </div>
                  <div className="legend-item">
                    <div className="legend-color red"></div>
                    <div className="legend-info">
                      <span className="legend-name">Belum Diverifikasi (Merah)</span>
                      <span className="legend-count">{redCount} CPB ({getPercentStyle(redCount)})</span>
                    </div>
                  </div>
                </div>
              </div>
            )}

            {/* Grid Kartu Metrik */}
            <div className="metrics-grid">
              <div className="metric-card accent-green">
                <span className="metric-title">Total CPB (Awal INVERS)</span>
                <span className="metric-value">
                  {recordsData?.invers_records?.length || 0}
                </span>
                <span className="metric-subtext">Revisi aktif ke-{stageSummary?.active_revision?.revision_num || 1}</span>
              </div>
              <div className="metric-card">
                <span className="metric-title">Terverifikasi Lolos</span>
                <span className="metric-value">{stageSummary?.totals?.lolos || 0}</span>
                <span className="metric-subtext">Memenuhi syarat (Lamp IIA)</span>
              </div>
              <div className="metric-card">
                <span className="metric-title">Tidak Lolos</span>
                <span className="metric-value">{stageSummary?.totals?.tidak_lolos || 0}</span>
                <span className="metric-subtext">Diberhentikan & diganti (Lamp IIIA)</span>
              </div>
              <div className="metric-card" style={{ borderLeft: getActiveMismatchCount() > 0 ? '4px solid var(--danger)' : '4px solid var(--success)' }}>
                <span className="metric-title">Ketidakcocokan / Error</span>
                <span className="metric-value">{getActiveMismatchCount()}</span>
                <span className="metric-subtext" style={{ color: getActiveMismatchCount() > 0 ? 'var(--danger)' : 'var(--success)', fontWeight: '600' }}>
                  {getActiveMismatchCount() > 0 ? `${getActiveMismatchCount()} Data Butuh Tindakan` : `Semua kasus diselesaikan (${recordsData?.mismatch_count || 0} total)`}
                </span>
              </div>
            </div>

            {/* Sesi Berita Acara */}
            <div className="card-section">
              <div className="card-header-title">
                <span>Berita Acara (BA) & Berkas Verifikasi Terunggah</span>
                <div style={{ display: 'flex', gap: '10px' }}>
                  <button 
                    className="btn btn-primary btn-sm"
                    onClick={() => handleExport()}
                    disabled={!stageSummary?.batches?.length}
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>Ekspor Semua BA ke Excel
                  </button>
                </div>
              </div>

              {stageSummary?.batches?.length > 0 && (
                <div style={{ display: 'flex', gap: '12px', alignItems: 'center', padding: '12px 0', borderBottom: '1px solid var(--border)' }}>
                  <div style={{ position: 'relative', flex: 1, maxWidth: '320px' }}>
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }}><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
                    <input
                      type="text"
                      className="form-input"
                      placeholder="Cari nama berita acara..."
                      value={baSearchTerm}
                      onChange={(e) => setBaSearchTerm(e.target.value)}
                      style={{ paddingLeft: '32px', fontSize: '0.85rem' }}
                    />
                  </div>
                  <select
                    className="form-input"
                    value={baFilterStatus}
                    onChange={(e) => setBaFilterStatus(e.target.value)}
                    style={{ fontSize: '0.85rem', maxWidth: '180px' }}
                  >
                    <option value="all">Semua Status</option>
                    <option value="published">Sudah Terbit</option>
                    <option value="unpublished">Belum Terbit</option>
                  </select>
                  <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)', whiteSpace: 'nowrap' }}>
                    {(() => {
                      const total = stageSummary.batches.length;
                      const filtered = stageSummary.batches.filter(b => {
                        const matchSearch = b.name.toLowerCase().includes(baSearchTerm.toLowerCase());
                        const matchStatus = baFilterStatus === 'all' || (baFilterStatus === 'published' && b.is_published) || (baFilterStatus === 'unpublished' && !b.is_published);
                        return matchSearch && matchStatus;
                      }).length;
                      return baSearchTerm || baFilterStatus !== 'all' ? `${filtered} dari ${total} batch` : `${total} batch`;
                    })()}
                  </span>
                </div>
              )}

              <div className="table-responsive">
                <table className="data-table">
                  <thead>
                    <tr>
                      <th>Nama Berita Acara / Batch</th>
                      <th>Tanggal Unggah</th>
                      <th>Jumlah Lolos</th>
                      <th>Tidak Lolos</th>
                      <th>CPB Pengganti</th>
                      <th>Jumlah Verifikasi</th>
                      <th>Terbit</th>
                      <th>Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    {stageSummary?.batches?.filter(b => {
                      const matchSearch = b.name.toLowerCase().includes(baSearchTerm.toLowerCase());
                      const matchStatus = baFilterStatus === 'all' || (baFilterStatus === 'published' && b.is_published) || (baFilterStatus === 'unpublished' && !b.is_published);
                      return matchSearch && matchStatus;
                    }).map((b, idx) => {
                      const isExpanded = expandedBatchId === b.id;
                      const breakdown = batchBreakdownCache[b.id] || [];
                      const isDragging = draggedBatchIndex === idx;
                      const isDropTarget = dragOverBatchIndex === idx;
                      return (
                        <React.Fragment key={b.id}>
                          <tr 
                            draggable
                            onDragStart={(e) => handleBatchDragStart(e, idx)}
                            onDragOver={(e) => handleBatchDragOver(e, idx)}
                            onDrop={(e) => handleBatchDrop(e, idx)}
                            style={{ 
                              cursor: 'grab',
                              opacity: isDragging ? 0.4 : 1,
                              borderTop: isDropTarget ? '2px solid var(--primary)' : undefined,
                              transition: 'all 0.15s ease'
                            }}
                            className={isDropTarget ? 'batch-row-drop-target' : ''}
                          >
                            <td style={{ fontWeight: '600', color: 'var(--primary)' }}>
                              <span 
                                title="Tarik dan lepas untuk mengubah urutan Berita Acara"
                                style={{ cursor: 'grab', marginRight: '8px', color: 'var(--text-muted)', userSelect: 'none', fontSize: '1rem', verticalAlign: 'middle', display: 'inline-block' }}
                              >
                                ⠿
                              </span>
                              <button
                                onClick={() => toggleBatchBreakdown(b.id)}
                                style={{ background: 'none', border: 'none', cursor: 'pointer', padding: '2px 6px', marginRight: '6px', fontSize: '0.8rem', color: 'var(--text-muted)', verticalAlign: 'middle' }}
                                title={isExpanded ? "Sembunyikan Rincian Kabupaten" : "Tampilkan Rincian Kabupaten"}
                              >
                                {isExpanded ? '▾' : '▸'}
                              </button>
                              {b.name}
                            </td>
                            <td>{new Date(b.uploaded_at).toLocaleString('id-ID')}</td>
                            <td>{b.lolos_count} CPB</td>
                            <td>{b.tidak_lolos_count} CPB</td>
                            <td>{b.replacement_count} CPB</td>
                            <td style={{ fontWeight: '600' }}>{b.lolos_count + b.tidak_lolos_count} CPB</td>
                            <td style={{ textAlign: 'center' }}>
                              <label className="publish-checkbox" title={b.is_published ? "Sudah terbit — klik untuk batalkan" : "Belum terbit — klik untuk tandai terbit"}>
                                <input
                                  type="checkbox"
                                  checked={!!b.is_published}
                                  onChange={() => handleTogglePublished(b.id)}
                                />
                                <span className="publish-checkmark"></span>
                                <span className="publish-label">{b.is_published ? 'Ya' : 'Tidak'}</span>
                              </label>
                            </td>
                            <td>
                              <div style={{ display: 'flex', gap: '8px' }}>
                                <button 
                                  className="btn btn-secondary btn-sm" 
                                  onClick={() => {
                                    setSelectedBatchIdForWord(b.id);
                                    setShowWordModal(true);
                                  }}
                                >
                                  <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>Cetak BA
                                </button>
                                <button 
                                  className="btn btn-secondary btn-sm" 
                                  onClick={() => handleExport(b.id)}
                                >
                                  <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>Ekspor Excel
                                </button>
                                <button 
                                  className="btn btn-secondary btn-sm" 
                                  onClick={() => openPreview(b.id)}
                                  title="Preview Isi Excel"
                                >
                                  <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"></path><circle cx="12" cy="12" r="3"></circle></svg>Preview
                                </button>
                                <button 
                                  className="btn btn-secondary btn-sm" 
                                  onClick={() => openRenameBatchModal(b.id, b.name)}
                                  title="Ubah Nama Berita Acara / Batch"
                                >
                                  <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>Ubah Nama
                                </button>
                                <button 
                                  className="btn btn-danger btn-sm"
                                  onClick={() => handleDeleteBatch(b.id)}
                                >
                                  <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>Hapus Batch
                                </button>
                              </div>
                            </td>
                          </tr>
                          {isExpanded && (
                            <>
                              <tr className="batch-breakdown-header-row">
                                <td colSpan="2" style={{ paddingLeft: '32px', fontWeight: '600', fontSize: '0.78rem', color: 'var(--text-muted)' }}>
                                  KABUPATEN / KOTA
                                </td>
                                <td style={{ fontWeight: '600', fontSize: '0.78rem', color: 'var(--text-muted)' }}>LOLOS</td>
                                <td style={{ fontWeight: '600', fontSize: '0.78rem', color: 'var(--text-muted)' }}>TIDAK LOLOS</td>
                                <td style={{ fontWeight: '600', fontSize: '0.78rem', color: 'var(--text-muted)' }}>CPB PENGGANTI</td>
                                <td style={{ fontWeight: '600', fontSize: '0.78rem', color: 'var(--text-muted)' }}>JUMLAH VERIFIKASI</td>
                                <td colSpan="2"></td>
                              </tr>
                              {breakdown.length > 0 ? breakdown.map((row, idx) => (
                                <tr key={idx} className="batch-breakdown-data-row">
                                  <td colSpan="2" style={{ paddingLeft: '36px', fontWeight: '500' }}>
                                    <span style={{ color: 'var(--primary)', marginRight: '6px', fontSize: '0.85rem' }}>↳</span>{row.kabupaten}
                                  </td>
                                  <td>{row.lolos} CPB</td>
                                  <td>{row.tidak_lolos} CPB</td>
                                  <td>{row.replacement} CPB</td>
                                  <td style={{ fontWeight: '600' }}>{row.lolos + row.tidak_lolos} CPB</td>
                                  <td colSpan="2"></td>
                                </tr>
                              )) : (
                                <tr className="batch-breakdown-data-row">
                                  <td colSpan="8" style={{ textAlign: 'center', padding: '12px', color: 'var(--text-muted)', fontStyle: 'italic' }}>
                                    Memuat data kabupaten...
                                  </td>
                                </tr>
                              )}
                            </>
                          )}
                        </React.Fragment>
                      );
                    })}
                    {!stageSummary?.batches?.length && (
                      <tr>
                        <td colSpan="8" style={{ textAlign: 'center', padding: '24px', color: 'var(--text-muted)' }}>
                          Belum ada batch verifikasi yang diunggah. Buka menu Unggah Data untuk menambahkan.
                        </td>
                      </tr>
                    )}
                    {stageSummary?.batches?.length > 0 && stageSummary.batches.filter(b => {
                      const matchSearch = b.name.toLowerCase().includes(baSearchTerm.toLowerCase());
                      const matchStatus = baFilterStatus === 'all' || (baFilterStatus === 'published' && b.is_published) || (baFilterStatus === 'unpublished' && !b.is_published);
                      return matchSearch && matchStatus;
                    }).length === 0 && (
                      <tr>
                        <td colSpan="8" style={{ textAlign: 'center', padding: '24px', color: 'var(--text-muted)' }}>
                          Tidak ada batch yang cocok dengan filter.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </>
        )}

        {/* Data INVERS List View */}
        {activeTab === 'invers' && (
          <div className="card-section">
            <div className="card-header-title">
              <span>File ke-1: Data Awal INVERS (Rujukan Utama)</span>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                <button 
                  className="btn btn-secondary btn-sm"
                  onClick={handleExportFilteredInvers}
                  style={{ display: 'flex', alignItems: 'center', gap: '6px', fontWeight: '600' }}
                >
                  <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                  Ekspor Data Terfilter
                </button>
                <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                  Total data awal: {getFilteredInvers().length} CPB
                </span>
              </div>
            </div>

            {/* Filter Bar INVERS */}
            <div className="filter-bar" style={{ marginBottom: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 600 }}>Filter:</span>
                <select
                  className="filter-select"
                  value={inversKabFilter}
                  onChange={e => { setInversKabFilter(e.target.value); setInversPage(1); }}
                >
                  <option value="ALL">Semua Kabupaten</option>
                  {uniqueInversKab.map((k, idx) => <option key={idx} value={k}>{k}</option>)}
                </select>
                <select
                  className="filter-select"
                  value={inversDesaFilter}
                  onChange={e => { setInversDesaFilter(e.target.value); setInversPage(1); }}
                >
                  <option value="ALL">Semua Desa/Kelurahan</option>
                  {uniqueInversDesa.map((d, idx) => <option key={idx} value={d}>{d}</option>)}
                </select>
                <select
                  className="filter-select"
                  value={inversStatusFilter}
                  onChange={e => { setInversStatusFilter(e.target.value); setInversPage(1); }}
                >
                  <option value="ALL">Semua Status</option>
                  <option value="VERIFIED">Terverifikasi</option>
                  <option value="NOT_VERIFIED">Belum Terverifikasi</option>
                </select>
              </div>
            </div>

            <div className="table-responsive">
              <table className="data-table">
                <thead>
                  <tr>
                    <th>No. Urut</th>
                    <th>Nama</th>
                    <th>NIK / No. KTP</th>
                    <th>No. KK</th>
                    <th>Jenis Kelamin</th>
                    <th>Desa/Kelurahan</th>
                    <th>Kecamatan</th>
                    <th>Kabupaten/Kota</th>
                    <th>Status Verifikasi</th>
                  </tr>
                </thead>
                <tbody>
                  {paginatedInvers.map(r => (
                    <tr key={r.id}>
                      <td>{r.no_urut}</td>
                      <td style={{ fontWeight: '600' }}>{r.nama}</td>
                      <td className="mono-digit">{r.no_ktp}</td>
                      <td className="mono-digit">{r.no_kk}</td>
                      <td>{r.jenis_kelamin}</td>
                      <td>{r.desa_kelurahan}</td>
                      <td>{r.kecamatan}</td>
                      <td>{r.kabupaten_kota}</td>
                      <td>
                        {verifiedNikSet.has(r.no_ktp?.trim()) ? (
                          <span style={{
                            padding: '4px 8px', borderRadius: '4px', fontWeight: '700', fontSize: '0.75rem',
                            backgroundColor: 'var(--success-light)', color: 'var(--success)',
                            border: '1px solid var(--success)', whiteSpace: 'nowrap'
                          }}>Terverifikasi</span>
                        ) : (
                          <span style={{
                            padding: '4px 8px', borderRadius: '4px', fontWeight: '700', fontSize: '0.75rem',
                            backgroundColor: '#fff3cd', color: '#856404',
                            border: '1px solid #ffc107', whiteSpace: 'nowrap'
                          }}>Belum</span>
                        )}
                      </td>
                    </tr>
                  ))}
                  {paginatedInvers.length === 0 && (
                    <tr>
                      <td colSpan="9" style={{ textAlign: 'center', padding: '24px', color: 'var(--text-muted)' }}>
                        Tidak ada data INVERS yang cocok dengan pencarian Anda.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls for Invers */}
            {totalInversPages > 1 && (
              <div className="pagination-container">
                <span className="pagination-info">
                  Menampilkan data {((inversPage - 1) * ITEMS_PER_PAGE) + 1} - {Math.min(inversPage * ITEMS_PER_PAGE, filteredInvers.length)} dari {filteredInvers.length} CPB (Halaman {inversPage} dari {totalInversPages})
                </span>
                <div className="pagination-buttons">
                  <button 
                    type="button" 
                    className="btn btn-secondary btn-sm"
                    onClick={() => setInversPage(prev => Math.max(prev - 1, 1))}
                    disabled={inversPage === 1}
                  >
                    Sebelumnya
                  </button>
                  <button 
                    type="button" 
                    className="btn btn-secondary btn-sm"
                    onClick={() => setInversPage(prev => Math.min(prev + 1, totalInversPages))}
                    disabled={inversPage === totalInversPages}
                  >
                    Berikutnya
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Data Terverifikasi List View */}
        {activeTab === 'verified' && (
          <div className="card-section">
            <div className="card-header-title">
              <span>File ke-2: Data Hasil Verifikasi Lapangan</span>
              <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                {selectedRecordIds.size > 0 && (
                  <button 
                    className="btn btn-danger btn-sm"
                    onClick={handleBulkDelete}
                    style={{ display: 'flex', alignItems: 'center', gap: '6px' }}
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
                    Hapus {selectedRecordIds.size} Data Terpilih
                  </button>
                )}
                <button 
                  className="btn btn-secondary btn-sm"
                  onClick={handleExportFilteredVerified}
                  style={{ display: 'flex', alignItems: 'center', gap: '6px', fontWeight: '600' }}
                >
                  <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                  Ekspor Data Terfilter
                </button>
                <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                  Total data terproses: {getFilteredVerified().length} CPB
                </span>
              </div>
            </div>

            {/* Filter Bar Verified */}
            <div className="filter-bar" style={{ marginBottom: '12px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                <span style={{ fontSize: '0.85rem', fontWeight: 600 }}>Filter:</span>
                <select
                  className="filter-select"
                  value={verifiedKabFilter}
                  onChange={e => { setVerifiedKabFilter(e.target.value); setVerifiedPage(1); }}
                >
                  <option value="ALL">Semua Kabupaten</option>
                  {uniqueVerifiedKab.map((k, idx) => <option key={idx} value={k}>{k}</option>)}
                </select>
                <select
                  className="filter-select"
                  value={verifiedDesaFilter}
                  onChange={e => { setVerifiedDesaFilter(e.target.value); setVerifiedPage(1); }}
                >
                  <option value="ALL">Semua Desa/Kelurahan</option>
                  {uniqueVerifiedDesa.map((d, idx) => <option key={idx} value={d}>{d}</option>)}
                </select>
                <select
                  className="filter-select"
                  value={verifiedHasilFilter}
                  onChange={e => { setVerifiedHasilFilter(e.target.value); setVerifiedPage(1); }}
                >
                  <option value="ALL">Semua Hasil Analisis</option>
                  <option value="LOLOS">Lolos</option>
                  <option value="TIDAK_LOLOS">Tidak Lolos</option>
                  <option value="ERROR">Error (Butuh Rekonsiliasi)</option>
                  <option value="SELESAI">Rekonsiliasi Selesai</option>
                </select>
                <select
                  className="filter-select"
                  value={verifiedBatchFilter}
                  onChange={e => { setVerifiedBatchFilter(e.target.value); setVerifiedPage(1); }}
                >
                  <option value="ALL">Semua Berita Acara</option>
                  {uniqueVerifiedBatch.map((b, idx) => <option key={idx} value={b}>{b}</option>)}
                </select>
              </div>
            </div>

            <div className="table-responsive">
              <table className="data-table">
                <thead>
                  <tr>
                    <th style={{ width: '40px', textAlign: 'center' }}>
                      <input 
                        type="checkbox" 
                        checked={selectedRecordIds.size === getFilteredVerified().length && getFilteredVerified().length > 0}
                        onChange={toggleAllRecordSelection}
                        style={{ cursor: 'pointer' }}
                      />
                    </th>
                    <th>Status</th>
                    <th>Nama</th>
                    <th>NIK / No. KTP</th>
                    <th>No. KK</th>
                    <th>Berita Acara</th>
                    <th>Kabupaten/Kota</th>
                    <th>Desa/Kelurahan</th>
                    <th>Hasil Analisis</th>
                    <th>CPB Pengganti (Jika Ada)</th>
                  </tr>
                </thead>
                <tbody>
                  {paginatedVerified.map(r => (
                    <tr key={r.id} className={r.has_error ? "row-error" : ""}>
                      <td style={{ textAlign: 'center' }}>
                        <input 
                          type="checkbox" 
                          checked={selectedRecordIds.has(r.id)}
                          onChange={() => toggleRecordSelection(r.id)}
                          style={{ cursor: 'pointer' }}
                        />
                      </td>
                      <td>
                        <span style={{
                          padding: '4px 8px',
                          borderRadius: '4px',
                          fontWeight: '700',
                          fontSize: '0.75rem',
                          backgroundColor: r.status === 'LOLOS' ? 'var(--success-light)' : 'var(--danger-light)',
                          color: r.status === 'LOLOS' ? 'var(--success)' : 'var(--danger)',
                          border: `1px solid ${r.status === 'LOLOS' ? 'var(--success)' : 'var(--danger)'}`,
                          whiteSpace: 'nowrap'
                        }}>
                          {r.status}
                        </span>
                      </td>
                      <td style={{ fontWeight: '600' }}>{r.nama}</td>
                      <td className="mono-digit">{r.no_ktp}</td>
                      <td className="mono-digit">{r.no_kk}</td>
                      <td style={{ fontSize: '0.8rem', fontWeight: '500' }}>{r.batch_name}</td>
                      <td>{r.kabupaten_kota || r.expected_invers?.kabupaten_kota || '-'}</td>
                      <td>{r.desa_kelurahan || r.expected_invers?.desa_kelurahan || '-'}</td>
                      <td>
                        {r.has_error ? (
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '4px' }}>
                            {r.errors.map((e, idx) => (
                              <span key={idx} style={{ fontSize: "0.75rem", color: "var(--danger)", fontWeight: "600", display: "flex", alignItems: "center", gap: "4px" }}><svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: "var(--danger)" }}><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>{e}</span>
                            ))}
                            {r.is_mismatch && (
                              <button 
                                className="btn btn-secondary btn-sm"
                                style={{ marginTop: '4px', alignSelf: 'flex-start' }}
                                onClick={() => {
                                  setActiveTab('reconciliation');
                                  setSearchTerm(r.no_ktp);
                                }}
                              >
                                <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>Selesaikan di Rekon
                              </button>
                            )}
                          </div>
                        ) : (
                          <span style={{ color: 'var(--success)', fontWeight: '600', display: 'flex', alignItems: 'center', gap: '4px' }}>
                              <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>
                              {r.override ? `Rekonsiliasi Selesai (${r.override.override_type === 'ACCEPT_VERIFIED' ? 'Disetujui Lapangan' : 'Koreksi Manual'})` : 'Cocok Terverifikasi'}
                            </span>
                        )}
                      </td>
                      <td>
                        {r.nama_pengganti ? (
                          <div style={{ fontSize: '0.8rem' }}>
                            <strong>{r.nama_pengganti}</strong>
                            <div style={{ color: 'var(--text-muted)' }}>NIK: {r.no_ktp_pengganti}</div>
                          </div>
                        ) : (
                          "-"
                        )}
                      </td>
                    </tr>
                  ))}
                  {paginatedVerified.length === 0 && (
                    <tr>
                      <td colSpan="10" style={{ textAlign: 'center', padding: '24px', color: 'var(--text-muted)' }}>
                        Tidak ada data verifikasi yang cocok dengan pencarian Anda.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Pagination Controls for Verified */}
            {totalVerifiedPages > 1 && (
              <div className="pagination-container">
                <span className="pagination-info">
                  Menampilkan data {((verifiedPage - 1) * ITEMS_PER_PAGE) + 1} - {Math.min(verifiedPage * ITEMS_PER_PAGE, filteredVerified.length)} dari {filteredVerified.length} CPB (Halaman {verifiedPage} dari {totalVerifiedPages})
                </span>
                <div className="pagination-buttons">
                  <button 
                    type="button" 
                    className="btn btn-secondary btn-sm"
                    onClick={() => setVerifiedPage(prev => Math.max(prev - 1, 1))}
                    disabled={verifiedPage === 1}
                  >
                    Sebelumnya
                  </button>
                  <button 
                    type="button" 
                    className="btn btn-secondary btn-sm"
                    onClick={() => setVerifiedPage(prev => Math.min(prev + 1, totalVerifiedPages))}
                    disabled={verifiedPage === totalVerifiedPages}
                  >
                    Berikutnya
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Overview Center (Pusat Rekap Data) View */}
        {activeTab === 'overview' && (
          <div className="card-section">
            <div className="card-header-title">
              <span><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><line x1="18" y1="20" x2="18" y2="10"></line><line x1="12" y1="20" x2="12" y2="4"></line><line x1="6" y1="20" x2="6" y2="14"></line></svg>Pusat Rekap Data (Overview Center)</span>
              <div className="tabs-header" style={{ marginBottom: 0 }}>
                <button 
                  className={`tab-btn ${overviewSubTab === 'summary_table' ? 'active' : ''}`}
                  onClick={() => setOverviewSubTab('summary_table')}
                >
                  Tabel Agregasi Wilayah
                </button>
                <button 
                  className={`tab-btn ${overviewSubTab === 'tree_pengusul' ? 'active' : ''}`}
                  onClick={() => setOverviewSubTab('tree_pengusul')}
                >
                  Hirarki Usulan Pengusul
                </button>
                <button 
                  className={`tab-btn ${overviewSubTab === 'kabupaten_pengusul' ? 'active' : ''}`}
                  onClick={() => setOverviewSubTab('kabupaten_pengusul')}
                >
                  Rekap Per Kabupaten
                </button>
              </div>
            </div>

            {overviewSubTab === 'summary_table' ? (
              <div style={{ display: 'flex', flexDirection: 'column', gap: '28px', marginTop: '16px' }}>
                {/* 1. Agregasi Kabupaten */}
                <div>
                  <h3 style={{ fontSize: '1rem', fontWeight: 700, marginBottom: '10px', color: 'var(--primary)' }}>1. Rekapitulasi Berdasarkan Kabupaten / Kota</h3>
                  <div className="table-responsive">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>Kabupaten / Kota</th>
                          <th>Total CPB Awal</th>
                          <th>Terverifikasi Lolos</th>
                          <th>Tidak Lolos</th>
                          <th>Belum Terverifikasi</th>
                        </tr>
                      </thead>
                      <tbody>
                        {overviewTables?.kabupaten?.map((k, idx) => (
                          <tr key={idx}>
                            <td style={{ fontWeight: '600' }}>{k.name}</td>
                            <td>{k.total_cpb} unit</td>
                            <td style={{ color: 'var(--success)', fontWeight: '600' }}>{k.lolos} unit</td>
                            <td style={{ color: '#856404', fontWeight: '600' }}>{k.tidak_lolos} unit</td>
                            <td style={{ color: 'var(--danger)', fontWeight: '600' }}>{k.belum_verifikasi} unit</td>
                          </tr>
                        ))}
                        {!overviewTables?.kabupaten?.length && (
                          <tr>
                            <td colSpan="5" style={{ textAlign: 'center', padding: '16px' }}>Belum ada data wilayah teragregasi.</td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>

                {/* 2. Agregasi Kecamatan */}
                <div>
                  <h3 style={{ fontSize: '1rem', fontWeight: 700, marginBottom: '10px', color: 'var(--primary)' }}>2. Rekapitulasi Berdasarkan Kecamatan</h3>
                  <div className="table-responsive">
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>Kecamatan</th>
                          <th>Total CPB Awal</th>
                          <th>Terverifikasi Lolos</th>
                          <th>Tidak Lolos</th>
                          <th>Belum Terverifikasi</th>
                        </tr>
                      </thead>
                      <tbody>
                        {overviewTables?.kecamatan?.map((kc, idx) => (
                          <tr key={idx}>
                            <td style={{ fontWeight: '600' }}>{kc.name}</td>
                            <td>{kc.total_cpb} unit</td>
                            <td style={{ color: 'var(--success)', fontWeight: '600' }}>{kc.lolos} unit</td>
                            <td style={{ color: '#856404', fontWeight: '600' }}>{kc.tidak_lolos} unit</td>
                            <td style={{ color: 'var(--danger)', fontWeight: '600' }}>{kc.belum_verifikasi} unit</td>
                          </tr>
                        ))}
                        {!overviewTables?.kecamatan?.length && (
                          <tr>
                            <td colSpan="5" style={{ textAlign: 'center', padding: '16px' }}>Belum ada data kecamatan teragregasi.</td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              </div>
            ) : overviewSubTab === 'tree_pengusul' ? (
              <div style={{ marginTop: '16px' }}>
                <div className="filter-bar" style={{ marginBottom: '0' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                    <span style={{ fontSize: '0.9rem', fontWeight: 600 }}>Filter Pengusul:</span>
                    <select 
                      className="filter-select"
                      value={selectedPengusulFilter}
                      onChange={e => setSelectedPengusulFilter(e.target.value)}
                    >
                      <option value="ALL">Tampilkan Semua Pengusul</option>
                      {getUniquePengusuls().map((p, idx) => (
                        <option key={idx} value={p}>{p}</option>
                      ))}
                    </select>
                  </div>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <button 
                      className="btn btn-secondary btn-sm"
                      onClick={handleExportPengusulTree}
                      style={{ display: 'flex', alignItems: 'center', gap: '6px', fontWeight: '600' }}
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                      Ekspor Excel
                    </button>
                    <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                      Klik pada baris untuk membuka hirarki
                    </span>
                  </div>
                </div>

                <div className="tree-container" style={{ padding: '0', maxHeight: '600px' }}>
                  {/* Table Header */}
                  <div style={{ 
                    display: 'grid', 
                    gridTemplateColumns: 'minmax(200px, 2fr) repeat(5, 90px)',
                    alignItems: 'center',
                    padding: '12px 14px',
                    backgroundColor: 'var(--primary)',
                    color: 'white',
                    fontWeight: '700',
                    fontSize: '0.8rem',
                    position: 'sticky',
                    top: 0,
                    zIndex: 10,
                    borderBottom: '2px solid var(--primary-hover)'
                  }}>
                    <div style={{ paddingLeft: '28px' }}>Nama / Wilayah</div>
                    <div style={{ textAlign: 'center' }}>CPB</div>
                    <div style={{ textAlign: 'center' }}>Lolos</div>
                    <div style={{ textAlign: 'center' }}>Tidak Lolos</div>
                    <div style={{ textAlign: 'center' }}>Belum</div>
                    <div style={{ textAlign: 'center' }}>Aksi</div>
                  </div>
                  
                  {/* Tree Rows */}
                  {getFilteredPengusulTree().map((node, idx) => (
                    <TreeNode key={idx} node={node} />
                  ))}
                  {getFilteredPengusulTree().length === 0 && (
                    <div style={{ textAlign: 'center', padding: '24px', color: 'var(--text-muted)' }}>
                      Belum ada data pohon hirarki pengusul. Silakan unggah berkas INVERS terlebih dahulu.
                    </div>
                  )}
                </div>
              </div>
            ) : (
              <div style={{ marginTop: '16px' }}>
                <div className="filter-bar" style={{ marginBottom: '0' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                    <span style={{ fontSize: '0.9rem', fontWeight: 600 }}>Rekap Per Kabupaten</span>
                    <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                      Klik pada baris kabupaten untuk melihat daftar pengusul
                    </span>
                  </div>
                </div>

                <div className="tree-container" style={{ padding: '0', maxHeight: '600px' }}>
                  {/* Table Header */}
                  <div style={{ 
                    display: 'grid', 
                    gridTemplateColumns: 'minmax(200px, 2fr) repeat(5, 90px)',
                    alignItems: 'center',
                    padding: '12px 14px',
                    backgroundColor: 'var(--primary)',
                    color: 'white',
                    fontWeight: '700',
                    fontSize: '0.8rem',
                    position: 'sticky',
                    top: 0,
                    zIndex: 10,
                    borderBottom: '2px solid var(--primary-hover)'
                  }}>
                    <div style={{ paddingLeft: '28px' }}>Kabupaten / Pengusul</div>
                    <div style={{ textAlign: 'center' }}>CPB</div>
                    <div style={{ textAlign: 'center' }}>Lolos</div>
                    <div style={{ textAlign: 'center' }}>Tidak Lolos</div>
                    <div style={{ textAlign: 'center' }}>Belum</div>
                    <div style={{ textAlign: 'center' }}>Aksi</div>
                  </div>

                  {/* Kabupaten Rows */}
                  {kabPengusulTree.map((kab, idx) => (
                    <KabupatenPengusulRow key={idx} kab={kab} />
                  ))}
                  {kabPengusulTree.length === 0 && (
                    <div style={{ textAlign: 'center', padding: '24px', color: 'var(--text-muted)' }}>
                      Belum ada data rekap per kabupaten. Silakan unggah berkas INVERS terlebih dahulu.
                    </div>
                  )}
                </div>
              </div>
            )}
          </div>
        )}

        {/* Reconciliation Center View */}
        {activeTab === 'reconciliation' && (
          <div className="card-section">
            <div className="card-header-title">
              <span><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><polyline points="23 4 23 10 17 10"></polyline><polyline points="1 20 1 14 7 14"></polyline><path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path></svg>Pusat Rekonsiliasi & Perbaikan Mismatch</span>
              <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                Periksa perbedaan entri NIK, KK, atau Nama antara berkas hasil verifikasi dengan data INVERS.
              </span>
            </div>

            {/* Filter Jenis Kesalahan & Pencarian */}
            <div className="filter-bar">
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                <input 
                  type="checkbox" 
                  checked={selectedMismatchNiks.size === getFilteredMismatches().length && getFilteredMismatches().length > 0}
                  onChange={toggleAllMismatchSelection}
                  style={{ cursor: 'pointer' }}
                  title="Pilih semua"
                />
                <span style={{ fontSize: '0.9rem', fontWeight: 600 }}>Filter Kategori Error:</span>
                <select 
                  className="filter-select"
                  value={errorFilter}
                  onChange={e => setErrorFilter(e.target.value)}
                >
                  <option value="ALL">Semua Kesalahan</option>
                  <option value="DUPLICATE">Duplikat (Sudah Terverifikasi Sebelumnya)</option>
                  <option value="NAMA_MISMATCH">Ketidakcocokan Nama</option>
                  <option value="KK_MISMATCH">Ketidakcocokan No. KK</option>
                  <option value="NIK_MISMATCH">Ketidakcocokan NIK (Beda NIK)</option>
                  <option value="MISSING_IN_INVERS">Tidak Ditemukan di INVERS (NIK Baru)</option>
                  <option value="NIK_INVALID">Format NIK Tidak Valid (Bukan 16 Digit)</option>
                  <option value="KK_INVALID">Format KK Tidak Valid (Bukan 16 Digit)</option>
                  <option value="NIK_KK_IDENTICAL">Format NIK dan KK Bernilai Sama</option>
                </select>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                {selectedMismatchNiks.size > 0 && (
                  <>
                    <button 
                      className="btn btn-primary btn-sm"
                      onClick={() => handleBulkReconcile('ACCEPT_VERIFIED')}
                      style={{ display: 'flex', alignItems: 'center', gap: '4px' }}
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="20 6 9 17 4 12"></polyline></svg>
                      Setujui {selectedMismatchNiks.size} Data
                    </button>
                    <button 
                      className="btn btn-secondary btn-sm"
                      onClick={() => handleBulkReconcile('MANUAL_EDIT')}
                      style={{ display: 'flex', alignItems: 'center', gap: '4px' }}
                    >
                      Tolak {selectedMismatchNiks.size} Data
                    </button>
                    <button 
                      className="btn btn-danger btn-sm"
                      onClick={handleBulkDeleteFromRecon}
                      style={{ display: 'flex', alignItems: 'center', gap: '4px' }}
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round"><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>
                      Hapus Data Lapangan
                    </button>
                  </>
                )}
                <span style={{ fontSize: '0.85rem', fontWeight: 600, color: 'var(--danger)' }}>
                  Ditemukan {getFilteredMismatches().length} kasus kesalahan
                  {selectedMismatchNiks.size > 0 && ` (${selectedMismatchNiks.size} dipilih)`}
                </span>
              </div>
            </div>

            {editingRecord && (
              <div className="modal-overlay">
                <div className="modal-content" style={{ width: '600px' }}>
                  <div className="modal-header">
                    <h3><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"></path></svg>Form Perbaikan Data Manual</h3>
                    <button 
                      type="button" 
                      className="modal-close"
                      onClick={() => setEditingRecord(null)}
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
                    <button className="btn btn-secondary btn-sm" onClick={() => setEditingRecord(null)}>Batal</button>
                    <button className="btn btn-primary btn-sm" onClick={() => handleReconcileOverride(editingRecord, 'MANUAL_EDIT')}>Terapkan Perbaikan</button>
                  </div>
                </div>
              </div>
            )}

            <div className="recon-grid">
              {getFilteredMismatches().map(r => (
                <div className="recon-card" key={r.id} style={{ borderLeft: r.override ? '4px solid var(--success)' : '4px solid var(--danger)' }}>
                  <div className="recon-card-header">
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                      {!r.override && (
                        <input 
                          type="checkbox" 
                          checked={selectedMismatchNiks.has(r.no_ktp)}
                          onChange={() => toggleMismatchSelection(r.no_ktp)}
                          style={{ cursor: 'pointer', flexShrink: 0 }}
                          title="Pilih untuk bulk reconcile"
                        />
                      )}
                      <div>
                        <strong>Kasus Error: {
                        r.mismatch_type === 'DUPLICATE' ? 'Duplikat Data' :
                        r.mismatch_type === 'NAMA_MISMATCH' ? 'Mismat Nama CPB' :
                        r.mismatch_type === 'KK_MISMATCH' ? 'Mismat No. KK' :
                        r.mismatch_type === 'NIK_MISMATCH' ? 'Mismat NIK CPB' :
                        r.mismatch_type === 'MISSING_IN_INVERS' ? 'NIK Tidak Ada di INVERS' :
                        r.mismatch_type === 'NIK_INVALID' ? 'Format NIK Salah' :
                        r.mismatch_type === 'KK_INVALID' ? 'Format KK Salah' :
                        r.mismatch_type === 'NIK_KK_IDENTICAL' ? 'NIK & KK Identik' : 'Perbedaan Karakter'
                      }</strong> (Batch: {r.batch_name})
                      </div>
                    </div>
                    {r.override ? (
                      <span style={{ color: 'var(--success)', fontWeight: '700', fontSize: '0.8rem' }}>
                        <span style={{ display: "flex", alignItems: "center", gap: "4px" }}><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ color: "var(--success)" }}><polyline points="20 6 9 17 4 12"></polyline></svg>Resolved ({r.override.override_type === "ACCEPT_VERIFIED" ? "Terima Lapangan" : "Edit Manual"})</span>
                      </span>
                    ) : (
                      <span style={{ color: 'var(--danger)', fontWeight: '700', fontSize: '0.8rem' }}>
                        <span style={{ display: "flex", alignItems: "center", gap: "4px" }}><svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: "var(--danger)" }}><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>Belum Selesai (Butuh Tindakan)</span>
                      </span>
                    )}
                  </div>

                  <div className="recon-card-body">
                    {/* Expected dari INVERS */}
                    <div className="recon-card-column expected">
                      <div style={{ fontWeight: '700', color: 'var(--primary)', marginBottom: '8px', fontSize: '0.8rem' }}>
                        DATA RUJUKAN (File Awal INVERS)
                      </div>
                      
                      {r.expected_invers ? (
                        <>
                          <div className="recon-val-group">
                            <span className="recon-val-label">Nama</span>
                            <span className="recon-val-text">{r.expected_invers.nama}</span>
                          </div>
                          <div className="recon-val-group">
                            <span className="recon-val-label">NIK (KTP)</span>
                            <span className="recon-val-text">{r.expected_invers.no_ktp}</span>
                          </div>
                          <div className="recon-val-group">
                            <span className="recon-val-label">No. KK</span>
                            <span className="recon-val-text">{r.expected_invers.no_kk}</span>
                          </div>
                          <div className="recon-val-group">
                            <span className="recon-val-label">Desa / Kelurahan</span>
                            <span className="recon-val-text">{r.expected_invers.desa_kelurahan || '-'}</span>
                          </div>
                        </>
                      ) : (
                        <div style={{ color: 'var(--text-muted)', fontStyle: 'italic', fontSize: '0.85rem', marginTop: '20px' }}>
                          NIK "{r.no_ktp}" tidak terdaftar di file rujukan INVERS aktif.
                        </div>
                      )}
                    </div>

                    {/* Actual Verified record */}
                    <div className="recon-card-column actual">
                      <div style={{ fontWeight: '700', color: 'var(--danger)', marginBottom: '8px', fontSize: '0.8rem' }}>
                        DATA LAPANGAN (File Hasil Verifikasi)
                      </div>
                      
                      <div className="recon-val-group">
                        <span className="recon-val-label">Nama</span>
                        <span className={`recon-val-text ${r.expected_invers && r.expected_invers.nama !== r.nama ? 'mismatch' : ''}`}>
                          {r.nama}
                        </span>
                      </div>
                      <div className="recon-val-group">
                        <span className="recon-val-label">NIK (KTP)</span>
                        <span className={`recon-val-text ${r.expected_invers && r.expected_invers.no_ktp !== r.no_ktp ? 'mismatch' : ''}`}>
                          {r.no_ktp}
                        </span>
                      </div>
                      <div className="recon-val-group">
                        <span className="recon-val-label">No. KK</span>
                        <span className={`recon-val-text ${r.expected_invers && r.expected_invers.no_kk !== r.no_kk ? 'mismatch' : ''}`}>
                          {r.no_kk}
                        </span>
                      </div>
                      <div className="recon-val-group">
                        <span className="recon-val-label">Desa / Kelurahan</span>
                        <span className="recon-val-text">{r.desa_kelurahan || '-'}</span>
                      </div>
                    </div>
                  </div>

                  <div className="recon-actions">
                    <button 
                      className="btn btn-secondary btn-sm" 
                      onClick={() => handleStartManualEdit(r)}
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"></path><path d="M18.5 2.5a2.121 2.121 0 1 1 3 3L12 15l-4 1 1-4 9.5-9.5z"></path></svg>Koreksi Manual
                    </button>
                    <button 
                      className="btn btn-secondary btn-sm" 
                      onClick={() => {
                        setLinkingRecord(r);
                        setLinkSearchTerm(r.nama);
                        setShowLinkModal(true);
                      }}
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg>Pasangkan dengan INVERS
                    </button>
                    <button 
                      className="btn btn-danger btn-sm" 
                      onClick={() => handleDeleteVerifiedRecord(r.id, r.nama)}
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><polyline points="3 6 5 6 21 6"></polyline><path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path></svg>Hapus Data Lapangan
                    </button>
                    {r.expected_invers && (
                      <button 
                        className="btn btn-primary btn-sm"
                        onClick={() => handleReconcileOverride(r, 'ACCEPT_VERIFIED')}
                      >
                        <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><polyline points="20 6 9 17 4 12"></polyline></svg>Setujui Data Lapangan
                      </button>
                    )}
                  </div>
                </div>
              ))}
              {getFilteredMismatches().length === 0 && (
                <div style={{ textAlign: 'center', padding: '48px', color: 'var(--text-muted)' }}>
                  Tidak ditemukan data mismatch untuk filter ini! Semua data bersih & sinkron.
                </div>
              )}
            </div>

            {/* Section: Belum Terverifikasi (Invers tanpa pasangan di verified) */}
            {unmatchedInvers.length > 0 && (
              <div style={{ marginTop: '24px', borderTop: '2px solid var(--border)', paddingTop: '16px' }}>
                <div
                  onClick={() => setSectionCollapsed(p => ({ ...p, belumVerif: !p.belumVerif }))}
                  style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', padding: '8px 0', userSelect: 'none' }}
                >
                  <span style={{ fontSize: '1rem', color: 'var(--text-muted)' }}>{sectionCollapsed.belumVerif ? '▶' : '▼'}</span>
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#e67e22' }}><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>
                  <span style={{ fontSize: '1rem', fontWeight: 700, color: '#e67e22' }}>
                    Belum Terverifikasi ({unmatchedInvers.length} data)
                  </span>
                </div>
                {!sectionCollapsed.belumVerif && (
                  <div>
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '12px' }}>
                      <span style={{ fontSize: '0.8rem', color: 'var(--text-muted)' }}>
                        — Data INVERS yang belum memiliki pasangan di Berita Acara.
                      </span>
                      <div style={{ display: 'flex', gap: '8px' }}>
                        <button className="btn btn-secondary btn-sm" onClick={() => { fetchSuggestedPairs(selectedStageId); fetchInlineSuggestions(unmatchedInvers.slice(0, 20)); }} style={{ whiteSpace: 'nowrap', fontSize: '0.8rem' }}>
                          🔍 Cari Saran
                        </button>
                        <button className="btn btn-primary btn-sm" onClick={handleAutoPairNIK} disabled={autoPairing} style={{ whiteSpace: 'nowrap', fontSize: '0.8rem' }}>
                          {autoPairing ? '⏳ Memproses...' : '⚡ Auto-Pair NIK/KK Cocok'}
                        </button>
                      </div>
                    </div>
                    {suggestedPairs.length > 0 && (
                      <div style={{ backgroundColor: '#f0f9ff', border: '1px solid #bae6fd', borderRadius: '8px', padding: '16px', marginBottom: '16px' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '12px' }}>
                          <span style={{ fontSize: '0.95rem', fontWeight: 700, color: '#0369a1' }}>📋 Kandidat Cocok Nama + Desa ({suggestedPairs.length} data)</span>
                          <div style={{ display: 'flex', gap: '8px' }}>
                            <button className="btn btn-primary btn-sm" onClick={handleBatchPair} disabled={batchPairing || selectedSuggestions.size === 0} style={{ fontSize: '0.8rem' }}>
                              {batchPairing ? '⏳ Memproses...' : `⚡ Pasangkan yang Dipilih (${selectedSuggestions.size})`}
                            </button>
                            <button className="btn btn-secondary btn-sm" onClick={() => { setSuggestedPairs([]); setSelectedSuggestions(new Set()); }} style={{ fontSize: '0.8rem' }}>Lewati</button>
                          </div>
                        </div>
                        <div style={{ maxHeight: '250px', overflowY: 'auto' }}>
                          <table className="data-table" style={{ width: '100%', fontSize: '0.85rem' }}>
                            <thead><tr><th style={{ width: '40px' }}><input type="checkbox" checked={selectedSuggestions.size === suggestedPairs.length} onChange={(e) => { if (e.target.checked) setSelectedSuggestions(new Set(suggestedPairs.map((_, i) => i))); else setSelectedSuggestions(new Set()); }} /></th><th>Nama INVERS</th><th>NIK INVERS</th><th>Nama Verified</th><th>NIK Verified</th><th>Desa</th></tr></thead>
                            <tbody>{suggestedPairs.map((s, i) => (
                              <tr key={s.invers.no_ktp} style={{ backgroundColor: selectedSuggestions.has(i) ? '#eff6ff' : '#fff' }}>
                                <td style={{ textAlign: 'center' }}><input type="checkbox" checked={selectedSuggestions.has(i)} onChange={(e) => { const next = new Set(selectedSuggestions); if (e.target.checked) next.add(i); else next.delete(i); setSelectedSuggestions(next); }} /></td>
                                <td style={{ fontWeight: 600 }}>{s.invers.nama}</td>
                                <td className="mono-digit">{s.invers.no_ktp}</td>
                                <td>{s.verified.nama}</td>
                                <td className="mono-digit">{s.verified.no_ktp}</td>
                                <td>{s.invers.desa_kelurahan}</td>
                              </tr>
                            ))}</tbody>
                          </table>
                        </div>
                      </div>
                    )}
                    {(() => {
                      const PAGE_SIZE = 20;
                      const totalPages = Math.ceil(unmatchedInvers.length / PAGE_SIZE);
                      const currentPage = sectionPage.belumVerif;
                      const pageItems = unmatchedInvers.slice((currentPage - 1) * PAGE_SIZE, currentPage * PAGE_SIZE);
                      return (
                        <div style={{ overflowX: 'auto' }}>
                          <table className="data-table" style={{ width: '100%' }}>
                            <thead><tr><th style={{ width: '40px' }}>NO</th><th>NIK</th><th>NAMA</th><th>KABUPATEN</th><th>KECAMATAN</th><th>DESA/KEL</th><th>SARAN</th><th style={{ width: '140px' }}>AKSI</th></tr></thead>
                            <tbody>{pageItems.map((ir, idx) => {
                              const suggestion = inlineSuggestions[ir.no_ktp];
                              const isSkipped = skippedNiks.has(ir.no_ktp);
                              const rowNum = (currentPage - 1) * PAGE_SIZE + idx + 1;
                              return (
                                <tr key={ir.no_ktp} style={{ backgroundColor: isSkipped ? '#f5f5f5' : suggestion ? '#f0fdf4' : '#fff8f0', opacity: isSkipped ? 0.5 : 1 }}>
                                  <td style={{ textAlign: 'center' }}>{rowNum}</td>
                                  <td style={{ fontFamily: 'monospace' }}>{ir.no_ktp}</td><td>{ir.nama}</td><td>{ir.kabupaten_kota}</td><td>{ir.kecamatan}</td><td>{ir.desa_kelurahan}</td>
                                  <td style={{ fontSize: '0.8rem' }}>{suggestion ? <span style={{ color: 'var(--success)' }}>✓ {suggestion.nama} ({suggestion.desa_kelurahan})</span> : isSkipped ? <span style={{ color: 'var(--text-muted)' }}>Dilewati</span> : <span style={{ color: 'var(--text-muted)' }}>-</span>}</td>
                                  <td>{suggestion && !isSkipped ? (
                                    <div style={{ display: 'flex', gap: '4px' }}>
                                      <button className="btn btn-primary btn-sm" onClick={() => handleInlinePair(ir, suggestion)} style={{ fontSize: '0.7rem', padding: '2px 8px', backgroundColor: '#16a34a' }}>✓ Ya</button>
                                      <button className="btn btn-sm" onClick={() => setSkippedNiks(prev => new Set([...prev, ir.no_ktp]))} style={{ fontSize: '0.7rem', padding: '2px 8px', backgroundColor: '#e5e7eb', color: '#374151' }}>✗</button>
                                      <button className="btn btn-primary btn-sm" onClick={() => handleOpenPairModal(ir)} style={{ fontSize: '0.7rem', padding: '2px 8px' }}>Manual</button>
                                    </div>
                                  ) : <button className="btn btn-primary btn-sm" onClick={() => handleOpenPairModal(ir)} disabled={isSkipped} style={{ fontSize: '0.75rem', padding: '3px 8px' }}>Pasangkan</button>}</td>
                                </tr>
                              );
                            })}</tbody>
                          </table>
                          {totalPages > 1 && (
                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px', padding: '10px 0' }}>
                              <button className="btn btn-secondary btn-sm" disabled={currentPage <= 1} onClick={() => setSectionPage(p => ({ ...p, belumVerif: p.belumVerif - 1 }))}>◀ Prev</button>
                              <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>Halaman {currentPage} / {totalPages}</span>
                              <button className="btn btn-secondary btn-sm" disabled={currentPage >= totalPages} onClick={() => setSectionPage(p => ({ ...p, belumVerif: p.belumVerif + 1 }))}>Next ▶</button>
                            </div>
                          )}
                        </div>
                      );
                    })()}
                  </div>
                )}
              </div>
            )}

            {/* Section: Daftar Pasangan yang Sudah Dibuat */}
            {manualPairs.length > 0 && (
              <div style={{ marginTop: '20px', borderTop: '2px solid var(--border)', paddingTop: '16px' }}>
                <div
                  onClick={() => setSectionCollapsed(p => ({ ...p, pasangan: !p.pasangan }))}
                  style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', padding: '8px 0', userSelect: 'none' }}
                >
                  <span style={{ fontSize: '1rem', color: 'var(--text-muted)' }}>{sectionCollapsed.pasangan ? '▶' : '▼'}</span>
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: 'var(--success)' }}><polyline points="20 6 9 17 4 12"></polyline></svg>
                  <span style={{ fontSize: '1rem', fontWeight: 700, color: 'var(--success)' }}>
                    Pasangan Tersimpan ({manualPairs.length} data)
                  </span>
                </div>
                {!sectionCollapsed.pasangan && (() => {
                  const PAGE_SIZE = 20;
                  const totalPages = Math.ceil(manualPairs.length / PAGE_SIZE);
                  const currentPage = sectionPage.pasangan;
                  const pageItems = manualPairs.slice((currentPage - 1) * PAGE_SIZE, currentPage * PAGE_SIZE);
                  return (
                    <div style={{ overflowX: 'auto' }}>
                      <table className="data-table" style={{ width: '100%' }}>
                        <thead><tr><th style={{ width: '40px' }}>NO</th><th>NIK INVERS</th><th>NAMA INVERS</th><th>KAB INVERS</th><th>NIK VERIFIED</th><th>NAMA VERIFIED</th><th>STATUS</th><th style={{ width: '80px' }}>AKSI</th></tr></thead>
                        <tbody>{pageItems.map((mp, idx) => {
                          const rowNum = (currentPage - 1) * PAGE_SIZE + idx + 1;
                          return (
                            <tr key={mp.id}>
                              <td style={{ textAlign: 'center' }}>{rowNum}</td>
                              <td className="mono-digit">{mp.invers_nik}</td>
                              <td>{mp.invers_nama}</td>
                              <td>{mp.invers_kabupaten}</td>
                              <td className="mono-digit">{mp.verified_nik}</td>
                              <td>{mp.verified_nama}</td>
                              <td><span style={{ padding: '2px 8px', borderRadius: '12px', fontSize: '0.75rem', fontWeight: '600', backgroundColor: mp.verified_status === 'LOLOS' ? '#d4edda' : '#f8d7da', color: mp.verified_status === 'LOLOS' ? '#155724' : '#721c24' }}>{mp.verified_status}</span></td>
                              <td><button className="btn btn-danger btn-sm" onClick={() => handleUnpairInvers(mp.id)} style={{ fontSize: '0.75rem', padding: '3px 8px' }}>Hapus</button></td>
                            </tr>
                          );
                        })}</tbody>
                      </table>
                      {totalPages > 1 && (
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px', padding: '10px 0' }}>
                          <button className="btn btn-secondary btn-sm" disabled={currentPage <= 1} onClick={() => setSectionPage(p => ({ ...p, pasangan: p.pasangan - 1 }))}>◀ Prev</button>
                          <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>Halaman {currentPage} / {totalPages}</span>
                          <button className="btn btn-secondary btn-sm" disabled={currentPage >= totalPages} onClick={() => setSectionPage(p => ({ ...p, pasangan: p.pasangan + 1 }))}>Next ▶</button>
                        </div>
                      )}
                    </div>
                  );
                })()}
              </div>
            )}

            {/* Section: Terverifikasi Belum Dipasangkan */}
            {unmatchedVerified.length > 0 && (
              <div style={{ marginTop: '20px', borderTop: '2px solid var(--border)', paddingTop: '16px' }}>
                <div
                  onClick={() => setSectionCollapsed(p => ({ ...p, verifiedBelum: !p.verifiedBelum }))}
                  style={{ display: 'flex', alignItems: 'center', gap: '8px', cursor: 'pointer', padding: '8px 0', userSelect: 'none' }}
                >
                  <span style={{ fontSize: '1rem', color: 'var(--text-muted)' }}>{sectionCollapsed.verifiedBelum ? '▶' : '▼'}</span>
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: '#0369a1' }}><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                  <span style={{ fontSize: '1rem', fontWeight: 700, color: '#0369a1' }}>
                    Terverifikasi Belum Dipasangkan ({unmatchedVerified.length} data)
                  </span>
                </div>
                {!sectionCollapsed.verifiedBelum && (() => {
                  const PAGE_SIZE = 20;
                  const totalPages = Math.ceil(unmatchedVerified.length / PAGE_SIZE);
                  const currentPage = sectionPage.verifiedBelum;
                  const pageItems = unmatchedVerified.slice((currentPage - 1) * PAGE_SIZE, currentPage * PAGE_SIZE);
                  return (
                    <div style={{ overflowX: 'auto' }}>
                      <table className="data-table" style={{ width: '100%' }}>
                        <thead><tr><th style={{ width: '40px' }}>NO</th><th>NIK</th><th>NAMA</th><th>DESA/KEL</th><th>STATUS</th></tr></thead>
                        <tbody>{pageItems.map((vr, idx) => {
                          const rowNum = (currentPage - 1) * PAGE_SIZE + idx + 1;
                          return (
                            <tr key={vr.id} style={{ backgroundColor: '#f0f9ff' }}>
                              <td style={{ textAlign: 'center' }}>{rowNum}</td>
                              <td style={{ fontFamily: 'monospace' }}>{vr.no_ktp}</td>
                              <td>{vr.nama}</td>
                              <td>{vr.desa_kelurahan}</td>
                              <td><span style={{ padding: '2px 8px', borderRadius: '12px', fontSize: '0.75rem', fontWeight: '600', backgroundColor: vr.status === 'LOLOS' ? '#d4edda' : '#f8d7da', color: vr.status === 'LOLOS' ? '#155724' : '#721c24' }}>{vr.status}</span></td>
                            </tr>
                          );
                        })}</tbody>
                      </table>
                      {totalPages > 1 && (
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '12px', padding: '10px 0' }}>
                          <button className="btn btn-secondary btn-sm" disabled={currentPage <= 1} onClick={() => setSectionPage(p => ({ ...p, verifiedBelum: p.verifiedBelum - 1 }))}>◀ Prev</button>
                          <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>Halaman {currentPage} / {totalPages}</span>
                          <button className="btn btn-secondary btn-sm" disabled={currentPage >= totalPages} onClick={() => setSectionPage(p => ({ ...p, verifiedBelum: p.verifiedBelum + 1 }))}>Next ▶</button>
                        </div>
                      )}
                    </div>
                  );
                })()}
              </div>
            )}
          </div>
        )}

        {/* Upload Data View */}
        {activeTab === 'upload' && (
          <div className="card-section">
            <div className="card-header-title">
              <span><svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>Unggah Berkas Excel (.xlsx)</span>
            </div>

            <div className="upload-grid">
              <div 
                className="upload-box" 
                onClick={() => setUploadType('invers')}
                style={{ border: uploadType === 'invers' ? '2px solid var(--primary)' : '2px dashed var(--border)' }}
              >
                <span style={{ fontSize: "2.5rem" }}><svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" style={{ color: "var(--primary)" }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline></svg></span>
                <h3>1. Unggah Data INVERS</h3>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                  Gunakan untuk mengunggah daftar target penerima awal (revisi baru akan dibuat secara otomatis).
                </p>
              </div>

              <div 
                className="upload-box"
                onClick={() => setUploadType('verified')}
                style={{ border: uploadType === 'verified' ? '2px solid var(--primary)' : '2px dashed var(--border)' }}
              >
                <span style={{ fontSize: "2.5rem" }}><svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: "var(--success)" }}><polyline points="20 6 9 17 4 12"></polyline></svg></span>
                <h3>2. Unggah Data Hasil Verifikasi</h3>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                  Gunakan untuk mengunggah data Lamp. IIA (Lolos) & Lamp. IIIA (Tidak Lolos/Pengganti) dari lapangan.
                </p>
              </div>

              <div 
                className="upload-box"
                onClick={() => setUploadType('sk_dirjen')}
                style={{ border: uploadType === 'sk_dirjen' ? '2px solid var(--primary)' : '2px dashed var(--border)' }}
              >
                <span style={{ fontSize: "2.5rem" }}><svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ color: "#7c3aed" }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg></span>
                <h3>3. Unggah Data SK Dirjen</h3>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                  Upload daftar Penerima Bantuan dari SK Dirjen. Akan dicocokkan dengan semua Data Terverifikasi.
                </p>
              </div>
            </div>

            {uploadType && uploadType !== 'sk_dirjen' && (
              <form onSubmit={handleUploadSubmit} className="card-section" style={{ marginTop: '24px', backgroundColor: '#fdfdfd' }}>
                <h3 style={{ marginBottom: '16px', color: 'var(--primary)', fontSize: '1rem' }}>
                  {uploadType === 'invers' ? "Parameter Unggah berkas INVERS" : "Parameter Unggah berkas Verifikasi Lapangan"}
                </h3>

                {uploadType === 'invers' ? (
                  <div className="form-group">
                    <label className="form-label">Nama Tahap / Sesi INVERS</label>
                    <input 
                      type="text" 
                      placeholder="Contoh: INVER TAHAP 1" 
                      className="form-input" 
                      value={stageNameInput}
                      onChange={e => setStageNameInput(e.target.value)}
                      required
                    />
                    <small style={{ color: 'var(--text-muted)' }}>
                      Saran format: INVER TAHAP [Angka], e.g., INVER TAHAP 1. Jika nama tahap sudah ada, file diunggah sebagai revisi aktif baru.
                    </small>
                  </div>
                ) : (
                  <div className="form-group">
                    <label className="form-label">Nama Berita Acara / Batch Verifikasi</label>
                    <input 
                      type="text" 
                      placeholder="Contoh: Berita Acara Pertama" 
                      className="form-input" 
                      value={batchNameInput}
                      onChange={e => setBatchNameInput(e.target.value)}
                      required
                    />
                    <small style={{ color: 'var(--text-muted)' }}>
                      Pencocokan NIK & Nama otomatis akan dilakukan dengan rujukan tahap: <strong>{selectedStageName}</strong>.
                    </small>
                  </div>
                )}

                <div className="form-group" style={{ margin: '16px 0' }}>
                  <label className="form-label">Pilih Berkas Excel (.xlsx)</label>
                  <div
                    className={`dropzone ${isDragOver ? 'drag-over' : ''} ${selectedFile ? 'file-selected' : ''}`}
                    onDragOver={(e) => { e.preventDefault(); setIsDragOver(true); }}
                    onDragLeave={() => setIsDragOver(false)}
                    onDrop={(e) => {
                      e.preventDefault();
                      setIsDragOver(false);
                      const file = e.dataTransfer.files[0];
                      if (file && file.name.endsWith('.xlsx')) {
                        setSelectedFile(file);
                      } else {
                        showToast('Hanya file .xlsx yang diterima', 'error');
                      }
                    }}
                    onClick={() => document.getElementById('upload-file-input').click()}
                  >
                    <input
                      id="upload-file-input"
                      type="file"
                      accept=".xlsx"
                      onChange={handleFileChange}
                      style={{ display: 'none' }}
                      required
                    />
                    {selectedFile ? (
                      <>
                        <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="var(--success)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>
                        <span className="dropzone-filename">{selectedFile.name}</span>
                        <span className="dropzone-size">{(selectedFile.size / 1024).toFixed(1)} KB</span>
                        <span className="dropzone-hint">Klik untuk ganti file</span>
                      </>
                    ) : (
                      <>
                        <svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="var(--text-muted)" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="17 8 12 3 7 8"></polyline><line x1="12" y1="3" x2="12" y2="15"></line></svg>
                        <span className="dropzone-text">Seret file Excel ke sini atau <strong>klik untuk memilih</strong></span>
                        <span className="dropzone-hint">Format: .xlsx (maksimal 1 file)</span>
                      </>
                    )}
                  </div>
                  <small style={{ color: 'var(--text-muted)' }}>
                    Sistem akan memvalidasi kolom wajib secara otomatis sebelum berkas dikirim ke server.
                  </small>
                </div>

                <div style={{ display: 'flex', gap: '12px', justifyContent: 'flex-end' }}>
                  <button type="button" className="btn btn-secondary" onClick={() => setUploadType('')}>Batal</button>
                  <button type="submit" className="btn btn-primary">Mulai Ekstraksi Berkas</button>
                </div>
              </form>
            )}

            {uploadType === 'sk_dirjen' && (
              <div className="card-section" style={{ marginTop: '24px', backgroundColor: '#fdfdfd' }}>
                <h3 style={{ marginBottom: '16px', color: '#7c3aed', fontSize: '1rem' }}>
                  Parameter Unggah Data SK Dirjen
                </h3>
                <div className="form-group">
                  <label className="form-label">Nama Tahap SK Dirjen</label>
                  <input 
                    type="text" 
                    placeholder="Contoh: SK DIRJEN TAHAP 1" 
                    className="form-input" 
                    value={skDirjenStageName}
                    onChange={e => setSkDirjenStageName(e.target.value)}
                    required
                  />
                  <small style={{ color: 'var(--text-muted)' }}>
                    Jika nama tahap sudah ada, data lama akan digantikan dengan data baru.
                  </small>
                </div>
                <div className="form-group" style={{ margin: '16px 0' }}>
                  <label className="form-label">Pilih Berkas Excel SK Dirjen (.xlsx)</label>
                  <input 
                    type="file" 
                    accept=".xlsx" 
                    onChange={handleSkDirjenUpload}
                    style={{ padding: '8px' }}
                    disabled={skDirjenUploading}
                  />
                  <small style={{ color: 'var(--text-muted)' }}>
                    {skDirjenUploading ? "Sedang mengupload dan mencocokkan data..." : "File akan langsung diupload dan dicocokkan saat dipilih."}
                  </small>
                </div>
              </div>
            )}
          </div>
        )}

        {/* SK Dirjen View */}
        {activeTab === 'sk-dirjen' && (
          <div className="card-section">
            <div className="card-header-title">
              <span>
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path><polyline points="14 2 14 8 20 8"></polyline><line x1="16" y1="13" x2="8" y2="13"></line><line x1="16" y1="17" x2="8" y2="17"></line></svg>
                SK Dirjen — {skDirjenActiveSubTab === 'daftar-pb' ? 'Daftar PB' : skDirjenActiveSubTab === 'rekap-per-tahap' ? 'Rekap PB Per Tahap' : 'Rekap PB Per Kabupaten'}
              </span>
            </div>

            {/* Sub-tabs content */}
            {skDirjenActiveSubTab === 'daftar-pb' && (
              <div>
                {/* Batch selector */}
                <div className="sk-dirjen-filter-bar">
                  <span className="sk-dirjen-filter-label">Batch:</span>
                  <select 
                    className="sk-dirjen-filter-select"
                    value={skDirjenSelectedBatch || ''}
                    onChange={(e) => {
                      const val = e.target.value || 'all';
                      setSkDirjenSelectedBatch(val);
                      setSkDirjenSearchTerm('');
                      setSkDirjenDebouncedSearch('');
                      setSkDirjenFilterKab('');
                      setSkDirjenFilterKec('');
                      setSkDirjenFilterDesa('');
                      setSkDirjenFilterTahap('');
                      setSkDirjenFilterAsalBatch('');
                      setSkDirjenFilterStatus('');
                      if (val === 'all') {
                        fetchSkDirjenAllRecords({});
                      } else {
                        fetchSkDirjenRecords(parseInt(val), {});
                      }
                    }}
                  >
                    <option value="all">Semua Batch SK Dirjen</option>
                    {skDirjenBatches.map(b => (
                      <option key={b.id} value={b.id}>{b.stage_name} ({b.total_records} data)</option>
                    ))}
                    </select>
                    {skDirjenSelectedBatch && skDirjenSelectedBatch !== 'all' && (
                      <button 
                        className="btn btn-danger" 
                        style={{ marginLeft: 8, padding: '4px 12px', fontSize: '0.75rem' }}
                        onClick={() => {
                          const b = skDirjenBatches.find(b => b.id === skDirjenSelectedBatch);
                          handleSkDirjenDeleteBatch(skDirjenSelectedBatch, b ? b.stage_name : '');
                        }}
                      >Hapus Tahap Ini</button>
                    )}
                </div>

                {/* Filters */}
                {skDirjenSelectedBatch && (
                  <div className="sk-dirjen-filter-bar">
                    <span className="sk-dirjen-filter-label">Filter:</span>
                    <select className="sk-dirjen-filter-select" value={skDirjenFilterKab} onChange={(e) => {
                      setSkDirjenFilterKab(e.target.value);
                      setSkDirjenFilterKec('');
                      setSkDirjenFilterDesa('');
                      fetchSkDirjenWithFilters({ kabupaten: e.target.value, kecamatan: '', desa: '', tahap: skDirjenFilterTahap, asal_batch: skDirjenFilterAsalBatch, status: skDirjenFilterStatus });
                    }}>
                      <option value="">Semua Kabupaten</option>
                      {[...new Set(skDirjenRecords.map(r => r.kabupaten_kota).filter(Boolean))].sort().map(k => (
                        <option key={k} value={k}>{k}</option>
                      ))}
                    </select>
                    <select className="sk-dirjen-filter-select" value={skDirjenFilterKec} onChange={(e) => {
                      setSkDirjenFilterKec(e.target.value);
                      setSkDirjenFilterDesa('');
                      fetchSkDirjenWithFilters({ kabupaten: skDirjenFilterKab, kecamatan: e.target.value, desa: '', tahap: skDirjenFilterTahap, asal_batch: skDirjenFilterAsalBatch, status: skDirjenFilterStatus });
                    }}>
                      <option value="">Semua Kecamatan</option>
                      {[...new Set(skDirjenRecords.filter(r => !skDirjenFilterKab || r.kabupaten_kota === skDirjenFilterKab).map(r => r.kecamatan).filter(Boolean))].sort().map(k => (
                        <option key={k} value={k}>{k}</option>
                      ))}
                    </select>
                    <select className="sk-dirjen-filter-select" value={skDirjenFilterDesa} onChange={(e) => {
                      setSkDirjenFilterDesa(e.target.value);
                      fetchSkDirjenWithFilters({ kabupaten: skDirjenFilterKab, kecamatan: skDirjenFilterKec, desa: e.target.value, tahap: skDirjenFilterTahap, asal_batch: skDirjenFilterAsalBatch, status: skDirjenFilterStatus });
                    }}>
                      <option value="">Semua Desa/Kelurahan</option>
                      {[...new Set(skDirjenRecords.filter(r => (!skDirjenFilterKab || r.kabupaten_kota === skDirjenFilterKab) && (!skDirjenFilterKec || r.kecamatan === skDirjenFilterKec)).map(r => r.desa_kelurahan).filter(Boolean))].sort().map(k => (
                        <option key={k} value={k}>{k}</option>
                      ))}
                    </select>
                    <select className="sk-dirjen-filter-select" value={skDirjenFilterTahap} onChange={(e) => {
                      setSkDirjenFilterTahap(e.target.value);
                      fetchSkDirjenWithFilters({ kabupaten: skDirjenFilterKab, kecamatan: skDirjenFilterKec, desa: skDirjenFilterDesa, tahap: e.target.value, asal_batch: skDirjenFilterAsalBatch, status: skDirjenFilterStatus });
                    }}>
                      <option value="">Semua Asal Tahap</option>
                      {[...new Set(skDirjenRecords.map(r => r.verified_stage_name).filter(Boolean))].sort().map(k => (
                        <option key={k} value={k}>{k}</option>
                      ))}
                    </select>
                    <select className="sk-dirjen-filter-select" value={skDirjenFilterAsalBatch} onChange={(e) => {
                      setSkDirjenFilterAsalBatch(e.target.value);
                      fetchSkDirjenWithFilters({ kabupaten: skDirjenFilterKab, kecamatan: skDirjenFilterKec, desa: skDirjenFilterDesa, tahap: skDirjenFilterTahap, asal_batch: e.target.value, status: skDirjenFilterStatus });
                    }}>
                      <option value="">Semua Asal Batch</option>
                      {[...new Set(skDirjenRecords.map(r => r.verified_batch_name).filter(Boolean))].sort().map(k => (
                        <option key={k} value={k}>{k}</option>
                      ))}
                    </select>
                    <select className="sk-dirjen-filter-select" value={skDirjenFilterStatus} onChange={(e) => {
                      setSkDirjenFilterStatus(e.target.value);
                      fetchSkDirjenWithFilters({ kabupaten: skDirjenFilterKab, kecamatan: skDirjenFilterKec, desa: skDirjenFilterDesa, tahap: skDirjenFilterTahap, asal_batch: skDirjenFilterAsalBatch, status: e.target.value });
                    }}>
                      <option value="">Semua Status</option>
                      <option value="PERFECT">Cocok</option>
                      <option value="NEEDS_APPROVAL">Perlu Persetujuan</option>
                      <option value="NO_MATCH">Tidak Ditemukan</option>
                      <option value="APPROVED">Disetujui</option>
                      <option value="MANUAL_PAIR">Dipasangkan</option>
                    </select>
                    <div style={{ marginLeft: 'auto' }}>
                      {skDirjenSelectedBatch !== 'all' && (
                        <button className="btn btn-secondary btn-sm" onClick={() => window.open(`${BACKEND_URL}/api/sk-dirjen/export/${skDirjenSelectedBatch}`, '_blank')}>
                          <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                          Ekspor Excel
                        </button>
                      )}
                    </div>
                  </div>
                )}

                {/* Stats summary */}
                {skDirjenSelectedBatch && (
                  <div style={{ display: 'flex', gap: '12px', marginBottom: '16px', flexWrap: 'wrap' }}>
                    <div style={{ padding: '8px 16px', borderRadius: '8px', background: '#dcfce7', border: '1px solid #bbf7d0', fontSize: '0.85rem' }}>
                      <strong style={{ color: '#166534' }}>Cocok:</strong> <span style={{ color: '#166534', fontWeight: 600 }}>{skDirjenRecords.filter(r => r.match_type === 'PERFECT').length}</span>
                    </div>
                    <div style={{ padding: '8px 16px', borderRadius: '8px', background: '#fef9c3', border: '1px solid #fde68a', fontSize: '0.85rem' }}>
                      <strong style={{ color: '#854d0e' }}>Perlu Persetujuan:</strong> <span style={{ color: '#854d0e', fontWeight: 600 }}>{skDirjenRecords.filter(r => r.match_type === 'NEEDS_APPROVAL').length}</span>
                    </div>
                    <div style={{ padding: '8px 16px', borderRadius: '8px', background: '#fee2e2', border: '1px solid #fecaca', fontSize: '0.85rem' }}>
                      <strong style={{ color: '#991b1b' }}>Tidak Ditemukan:</strong> <span style={{ color: '#991b1b', fontWeight: 600 }}>{skDirjenRecords.filter(r => r.match_type === 'NO_MATCH').length}</span>
                    </div>
                    <div style={{ padding: '8px 16px', borderRadius: '8px', background: '#e0e7ff', border: '1px solid #c7d2fe', fontSize: '0.85rem' }}>
                      <strong style={{ color: '#3730a3' }}>Disetujui:</strong> <span style={{ color: '#3730a3', fontWeight: 600 }}>{skDirjenRecords.filter(r => r.match_type === 'NEEDS_APPROVAL' && r.override_status === 'APPROVED').length}</span>
                    </div>
                    <div style={{ padding: '8px 16px', borderRadius: '8px', background: '#ede9fe', border: '1px solid #ddd6fe', fontSize: '0.85rem' }}>
                      <strong style={{ color: '#6d28d9' }}>Dipasangkan:</strong> <span style={{ color: '#6d28d9', fontWeight: 600 }}>{skDirjenRecords.filter(r => r.match_type === 'MANUAL_PAIR').length}</span>
                    </div>
                  </div>
                )}

                {/* Search bar */}
                {skDirjenSelectedBatch && (
                  <div style={{ marginBottom: '16px' }}>
                    <div className="search-wrapper" style={{ maxWidth: '450px' }}>
                      <span className="search-icon">
                        <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
                      </span>
                      <input 
                        type="text" 
                        placeholder="Cari Nama, NIK, KK, Desa, Kecamatan, Kabupaten..."
                        className="search-input"
                        value={skDirjenSearchTerm}
                        onChange={(e) => setSkDirjenSearchTerm(e.target.value)}
                      />
                    </div>
                  </div>
                )}

                {/* Records table */}
                {skDirjenSelectedBatch && skDirjenRecords.length > 0 && (
                  <div style={{ overflowX: 'auto' }}>
                    <table className="data-table">
                      <thead>
                        <tr>
                          <th>No</th>
                          {skDirjenSelectedBatch === 'all' && <th>Tahap</th>}
                          <th>NIK</th>
                          <th>No KK</th>
                          <th>Nama</th>
                          <th>Desa/Kel</th>
                          <th>Kecamatan</th>
                          <th>Kab/Kota</th>
                          <th>Status</th>
                          <th>Asal Verifikasi</th>
                          <th>Aksi</th>
                        </tr>
                      </thead>
                      <tbody>
                        {skDirjenRecords.map((r) => {
                          const statusBadge = r.match_type === 'PERFECT' 
                            ? <span style={{ color: '#16a34a', fontWeight: 600, fontSize: '0.8rem' }}>Cocok</span>
                            : r.match_type === 'MANUAL_PAIR'
                              ? <span style={{ color: '#7c3aed', fontWeight: 600, fontSize: '0.8rem' }}>Dipasangkan</span>
                              : r.match_type === 'NEEDS_APPROVAL' 
                                ? <span style={{ color: '#ca8a04', fontWeight: 600, fontSize: '0.8rem' }}>{r.override_status === 'APPROVED' ? 'Disetujui' : 'Perlu Persetujuan'}</span>
                                : <span style={{ color: '#dc2626', fontWeight: 600, fontSize: '0.8rem' }}>Tidak Ditemukan</span>;
                          
                          const asal = r.match_type !== 'NO_MATCH' && r.verified_batch_name 
                            ? `BA ${r.verified_batch_name} Tahap ${r.verified_stage_name}` 
                            : '';

                          return (
                            <tr key={r.id} style={{ background: r.match_type === 'NEEDS_APPROVAL' && r.override_status !== 'APPROVED' ? '#fffbeb' : r.match_type === 'NO_MATCH' ? '#fef2f2' : undefined }}>
                              <td>{r.no_urut}</td>
                              {skDirjenSelectedBatch === 'all' && <td style={{ fontSize: '0.8rem', color: '#6b7280' }}>{r.batch_stage_name}</td>}
                              <td style={{ fontSize: '0.85rem' }}>{r.no_ktp}</td>
                              <td style={{ fontSize: '0.85rem' }}>{r.no_kk}</td>
                              <td>{r.nama}</td>
                              <td style={{ fontSize: '0.85rem' }}>{r.desa_kelurahan}</td>
                              <td style={{ fontSize: '0.85rem' }}>{r.kecamatan}</td>
                              <td style={{ fontSize: '0.85rem' }}>{r.kabupaten_kota}</td>
                              <td>{statusBadge}</td>
                              <td style={{ fontSize: '0.8rem', color: '#6b7280' }}>{asal}</td>
                              <td>
                                {r.match_type === 'NEEDS_APPROVAL' && r.override_status !== 'APPROVED' && (
                                  <button className="btn btn-primary btn-sm" onClick={() => setSkDirjenApprovalRecord(r)}>
                                    Tinjau
                                  </button>
                                )}
                                {r.match_type === 'NO_MATCH' && (
                                  <button className="btn btn-primary btn-sm" onClick={() => {
                                    setSkDirjenPairingRecord(r);
                                    setSkDirjenPairSearchTerm(r.nama);
                                    handleSkDirjenSearchVerified(r.nama);
                                  }}>
                                    Pasangkan
                                  </button>
                                )}
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                )}

                {skDirjenSelectedBatch && skDirjenRecords.length === 0 && (
                  <div style={{ textAlign: 'center', color: '#94a3b8', padding: '40px' }}>
                    {skDirjenDebouncedSearch ? 'Tidak ada data yang cocok dengan pencarian' : 'Tidak ada data ditemukan'}
                  </div>
                )}

                {!skDirjenSelectedBatch && (
                  <div style={{ textAlign: 'center', color: '#94a3b8', padding: '40px' }}>
                    {skDirjenBatches.length === 0 
                      ? 'Belum ada batch SK Dirjen. Silakan upload data SK Dirjen terlebih dahulu di menu Unggah Data.'
                      : 'Pilih batch SK Dirjen untuk melihat data'}
                  </div>
                )}
              </div>
            )}

            {/* Rekap Per Tahap */}
            {skDirjenActiveSubTab === 'rekap-per-tahap' && (
              <div>
                <div style={{ marginBottom: '16px', display: 'flex', justifyContent: 'flex-end' }}>
                  <button className="btn btn-secondary btn-sm" onClick={() => window.open(`${BACKEND_URL}/api/sk-dirjen/rekap-per-tahap/export`, '_blank')}>
                    <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                    Ekspor Excel
                  </button>
                </div>
                {skDirjenRekapPerTahap && skDirjenRekapPerTahap.rekap.length > 0 ? (
                  <div style={{ overflowX: 'auto' }}>
                    <table className="rekap-sk-dirjen-table">
                      <thead>
                        <tr>
                          <th>Tahap SK Dirjen</th>
                          {skDirjenRekapPerTahap.invers_stages.map(s => (
                            <th key={s}>{s}</th>
                          ))}
                          <th>Total</th>
                        </tr>
                      </thead>
                      <tbody>
                        {skDirjenRekapPerTahap.rekap.map((row, i) => (
                          <tr key={i}>
                            <td className="col-text">{row.stage_name}</td>
                            {skDirjenRekapPerTahap.invers_stages.map(s => (
                              <td key={s}>
                                {row.tahap[s] ? <span className="rekap-digit">{row.tahap[s]}</span> : <span style={{ color: '#94a3b8' }}>-</span>}
                              </td>
                            ))}
                            <td style={{ fontWeight: 600 }} className="rekap-digit">{row.total}</td>
                          </tr>
                        ))}
                      </tbody>
                      <tfoot>
                        <tr>
                          <td className="col-text-total">TOTAL</td>
                          {skDirjenRekapPerTahap.invers_stages.map(s => {
                            const total = skDirjenRekapPerTahap.rekap.reduce((sum, row) => sum + (row.tahap[s] || 0), 0);
                            return (
                              <td key={s}>
                                {total ? <span className="rekap-digit">{total}</span> : <span style={{ color: '#94a3b8' }}>-</span>}
                              </td>
                            );
                          })}
                          <td className="rekap-digit">{skDirjenRekapPerTahap.rekap.reduce((sum, row) => sum + row.total, 0)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                ) : (
                  <div style={{ textAlign: 'center', color: '#94a3b8', padding: '40px' }}>Belum ada data rekap SK Dirjen</div>
                )}
              </div>
            )}
            {skDirjenActiveSubTab === 'rekap-per-kabupaten' && (
              <div>
                <div className="sk-dirjen-filter-bar">
                  <span className="sk-dirjen-filter-label">Filter:</span>
                  <select 
                    className="sk-dirjen-filter-select"
                    value={skDirjenSelectedKabForRekap || 'all'}
                    onChange={(e) => {
                      const val = e.target.value || null;
                      setSkDirjenSelectedKabForRekap(val ? (isNaN(val) ? val : parseInt(val)) : null);
                      if (val) {
                        fetchSkDirjenRekapPerKab(isNaN(val) ? val : parseInt(val));
                      }
                    }
                  }
                >
                  <option value="all">Semua Batch (Gabungan)</option>
                  {skDirjenBatches.map(b => (
                    <option key={b.id} value={b.id}>{b.stage_name} ({b.total_records} data)</option>
                  ))}
                </select>
                {skDirjenSelectedKabForRekap && skDirjenSelectedKabForRekap !== 'all' && (
                  <div style={{ marginLeft: 'auto' }}>
                    <button className="btn btn-secondary btn-sm" onClick={() => window.open(`${BACKEND_URL}/api/sk-dirjen/rekap-per-kabupaten/${skDirjenSelectedKabForRekap}/export`, '_blank')}>
                      <svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "4px", verticalAlign: "middle" }}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                      Ekspor Excel
                    </button>
                  </div>
                )}
              </div>

              {skDirjenRekapPerKab && skDirjenRekapPerKab.kabupatens.length > 0 ? (
                <div style={{ overflowX: 'auto' }}>
                  <table className="rekap-sk-dirjen-table">
                    <thead>
                      <tr>
                        <th rowSpan="2">No.</th>
                        <th rowSpan="2">KABUPATEN/KOTA</th>
                        {skDirjenRekapPerKab.batch_id === null && (
                          <th rowSpan="2" className="header-lolos">CPB LOLOS</th>
                        )}
                        <th colSpan={skDirjenRekapPerKab.invers_stages.length} className="header-sk-dirjen">
                          SK DIRJEN {skDirjenRekapPerKab.stage_name.toUpperCase()}
                        </th>
                        <th rowSpan="2">Total</th>
                        {skDirjenRekapPerKab.batch_id === null && (
                          <th rowSpan="2" className="header-selisih">SELISIH SK</th>
                        )}
                        {skDirjenRekapPerKab.batch_id === null && (
                          <th rowSpan="2" className="header-sumber">SUMBER SELISIH</th>
                        )}
                      </tr>
                      <tr>
                        {skDirjenRekapPerKab.invers_stages.map(s => (
                          <th key={s}>{s}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {skDirjenRekapPerKab.kabupatens.map((row, i) => (
                        <tr key={i}>
                          <td className="col-no">{i + 1}</td>
                          <td className="col-text">{row.kabupaten}</td>
                          {skDirjenRekapPerKab.batch_id === null && (
                            <td style={{ fontWeight: 600 }} className="rekap-digit">{row.cpb_lolos || 0}</td>
                          )}
                          {skDirjenRekapPerKab.invers_stages.map(s => (
                            <td key={s}>
                              {row[s] ? <span className="rekap-digit">{row[s]}</span> : <span style={{ color: '#94a3b8' }}>-</span>}
                            </td>
                          ))}
                          <td style={{ fontWeight: 600 }} className="rekap-digit">{row.total}</td>
                          {skDirjenRekapPerKab.batch_id === null && (
                            <td style={{ fontWeight: 600, color: row.selisih_sk > 0 ? '#10b981' : row.selisih_sk < 0 ? '#ef4444' : '#6b7280' }} className="rekap-digit">
                              {row.selisih_sk > 0 ? '+' : ''}{row.selisih_sk}
                            </td>
                          )}
                          {skDirjenRekapPerKab.batch_id === null && (
                            <td className="col-sumber">
                              {row.sumber_selisih?.length > 0 ? (
                                <button
                                  className="sumber-selisih-trigger"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    setSumberSelisihPopup(sumberSelisihPopup?.kabupaten === row.kabupaten ? null : {
                                      kabupaten: row.kabupaten,
                                      selisih: row.selisih_sk,
                                      data: row.sumber_selisih
                                    });
                                  }}
                                  title="Klik untuk lihat rincian"
                                >
                                  {row.sumber_selisih.length} sumber
                                </button>
                              ) : (
                                <span style={{ color: '#94a3b8' }}>-</span>
                              )}
                            </td>
                          )}
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr>
                        <td></td>
                        <td className="col-text-total">TOTAL</td>
                        {skDirjenRekapPerKab.batch_id === null && (
                          <td style={{ fontWeight: 700 }} className="rekap-digit">{skDirjenRekapPerKab.kabupatens.reduce((sum, row) => sum + (row.cpb_lolos || 0), 0)}</td>
                        )}
                        {skDirjenRekapPerKab.invers_stages.map(s => {
                          const total = skDirjenRekapPerKab.kabupatens.reduce((sum, row) => sum + (row[s] || 0), 0);
                          return (
                            <td key={s}>
                              {total ? <span className="rekap-digit">{total}</span> : <span style={{ color: '#94a3b8' }}>-</span>}
                            </td>
                          );
                        })}
                        <td style={{ fontWeight: 700 }} className="rekap-digit">{skDirjenRekapPerKab.kabupatens.reduce((sum, row) => sum + row.total, 0)}</td>
                        {skDirjenRekapPerKab.batch_id === null && (() => {
                          const totalSelisih = skDirjenRekapPerKab.kabupatens.reduce((sum, row) => sum + (row.selisih_sk || 0), 0);
                          return (
                            <td style={{ fontWeight: 700, color: totalSelisih > 0 ? '#10b981' : totalSelisih < 0 ? '#ef4444' : '#6b7280' }} className="rekap-digit">
                              {totalSelisih > 0 ? '+' : ''}{totalSelisih}
                            </td>
                          );
                        })()}
                        {skDirjenRekapPerKab.batch_id === null && (
                          <td style={{ color: '#94a3b8', fontSize: '0.8rem' }}></td>
                        )}
                      </tr>
                    </tfoot>
                  </table>
                </div>
              ) : skDirjenSelectedKabForRekap ? (
                <div style={{ textAlign: 'center', color: '#94a3b8', padding: '40px' }}>Tidak ada data rekap untuk batch ini</div>
              ) : (
                <div style={{ textAlign: 'center', color: '#94a3b8', padding: '40px' }}>Pilih batch SK Dirjen untuk melihat rekap per kabupaten</div>
              )}
            </div>
            )}

            {/* Floating Window: Sumber Selisih */}
            {sumberSelisihPopup && (
              <div className="sumber-selisih-overlay" onClick={() => setSumberSelisihPopup(null)}>
                <div className="sumber-selisih-popup" onClick={(e) => e.stopPropagation()}>
                  <div className="sumber-selisih-header">
                    <div>
                      <strong>Sumber Selisih</strong>
                      <span style={{ marginLeft: '8px', fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                        {sumberSelisihPopup.kabupaten}
                      </span>
                    </div>
                    <span className="sumber-selisih-badge" style={{
                      background: sumberSelisihPopup.selisih > 0 ? 'var(--success-light)' : 'var(--danger-light)',
                      color: sumberSelisihPopup.selisih > 0 ? 'var(--success)' : 'var(--danger)'
                    }}>
                      {sumberSelisihPopup.selisih > 0 ? '+' : ''}{sumberSelisihPopup.selisih} PB
                    </span>
                    <button className="sumber-selisih-close" onClick={() => setSumberSelisihPopup(null)}>&times;</button>
                  </div>
                  <div className="sumber-selisih-body">
                    <p style={{ fontSize: '0.8rem', color: 'var(--text-muted)', marginBottom: '12px' }}>
                      Data CPB lolos yang belum terdokumentasi di SK Dirjen:
                    </p>
                    {sumberSelisihPopup.data.map((item, idx) => {
                      const isExpanded = expandedSumber?.kab === sumberSelisihPopup.kabupaten && expandedSumber?.idx === idx;
                      return (
                        <div key={idx}>
                          <div 
                            className="sumber-selisih-item sumber-selisih-item-expandable"
                            onClick={() => fetchSumberSelisihDetail(sumberSelisihPopup.kabupaten, item.batch_name, item.stage_name, idx)}
                          >
                            <span className="sumber-selisih-num">{idx + 1}.</span>
                            <div className="sumber-selisih-info">
                              <span className="sumber-selisih-batch">{item.batch_name}</span>
                              <span className="sumber-selisih-stage">{item.stage_name}</span>
                            </div>
                            <span className="sumber-selisih-count">{item.cnt} PB</span>
                            <span className="sumber-selisih-expand-icon">{isExpanded ? '\u25B2' : '\u25BC'}</span>
                          </div>
                          {isExpanded && (
                            <div className="sumber-selisih-detail">
                              {sumberDetailLoading ? (
                                <div className="sumber-selisih-detail-loading">Memuat data...</div>
                              ) : sumberDetail.length > 0 ? (
                                <>
                                  <div className="sumber-selisih-detail-info">
                                    Menampilkan {sumberDetail.length} data
                                  </div>
                                  <table className="sumber-selisih-detail-table">
                                    <thead>
                                      <tr>
                                        <th>No</th>
                                        <th>Nama</th>
                                        <th>NIK</th>
                                        <th>No.KK</th>
                                        <th>Desa/Kelurahan</th>
                                        <th>Kabupaten</th>
                                      </tr>
                                    </thead>
                                    <tbody>
                                      {sumberDetail.map((rec, i) => (
                                        <tr key={i}>
                                          <td>{i + 1}</td>
                                          <td>{rec.nama}</td>
                                          <td>{rec.no_ktp}</td>
                                          <td>{rec.no_kk}</td>
                                          <td>{rec.desa_kelurahan}</td>
                                          <td>{rec.kabupaten_kota}</td>
                                        </tr>
                                      ))}
                                    </tbody>
                                  </table>
                                </>
                              ) : (
                                <div className="sumber-selisih-detail-empty">Tidak ada data</div>
                              )}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Settings View */}
        {activeTab === 'settings' && (
          <SettingsPanel 
            stages={stages}
            selectedStageId={selectedStageId}
            onDeleteStage={handleDeleteStage}
            onClearDatabase={handleClearDatabase}
            showToast={showToast}
          />
        )}

        {/* Help Panel (Bantuan) View */}
        {activeTab === 'help' && (
          <div className="help-container">
            <div className="card-section">
              <h2 style={{ fontSize: '1.2rem', fontWeight: 700, color: 'var(--primary)', marginBottom: '8px' }}>
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><path d="M2 3h6a4 4 0 0 1 4 4v14a3 3 0 0 0-3-3H2z"></path><path d="M22 3h-6a4 4 0 0 0-4 4v14a3 3 0 0 1 3-3h7z"></path></svg>Panduan Penggunaan Aplikasi BSPS DB
              </h2>
              <p style={{ color: 'var(--text-muted)', fontSize: '0.9rem' }}>
                Ikuti langkah-langkah berikut untuk melakukan verifikasi, penyelarasan, dan ekspor dokumen berita acara bantuan perumahan swadaya.
              </p>
            </div>

            <div className="help-step-card">
              <div className="help-step-number">1</div>
              <div className="help-step-content">
                <h3>Unduh & Sesuaikan Template</h3>
                <p>
                  Buka menu <strong>Pengaturan</strong>, lalu unduh berkas template. Pastikan baris header utama berkas Excel Anda memiliki judul kolom seperti: <code>NAMA</code>, <code>NIK/NO KTP</code>, dan <code>NO KK</code>.
                </p>
              </div>
            </div>

            <div className="help-step-card">
              <div className="help-step-number">2</div>
              <div className="help-step-content">
                <h3>Unggah Data INVERS (File ke-1)</h3>
                <p>
                  Buka menu <strong>Unggah Data</strong>, pilih opsi <strong>Upload INVERS</strong>. Masukkan Nama Tahap (misal: <code>INVER TAHAP 1</code>), pilih file Excel target, lalu klik <strong>Mulai Ekstraksi Berkas</strong>. Ini akan menjadi data referensi pencocokan utama.
                </p>
              </div>
            </div>

            <div className="help-step-card">
              <div className="help-step-number">3</div>
              <div className="help-step-content">
                <h3>Unggah Hasil Verifikasi (File ke-2)</h3>
                <p>
                  Kembali ke menu <strong>Unggah Data</strong>, pilih opsi <strong>Upload DATA VERIFIKASI</strong>. Masukkan nama Berita Acara / batch (misal: <code>BA Tahap 1 Luwu Utara</code>), pilih file verifikasi lapangan berformat 2 sheet (Lamp. IIA & Lamp. IIIA), lalu unggah. Sistem akan mencocokkan NIK & nama secara otomatis.
                </p>
              </div>
            </div>

            <div className="help-step-card">
              <div className="help-step-number">4</div>
              <div className="help-step-content">
                <h3>Selesaikan Ketidakcocokan di Rekonsiliasi</h3>
                <p>
                  Jika terdapat data yang statusnya <code>Mismatch</code> (dikarenakan beda ejaan nama, KK salah, atau NIK tidak terdaftar), buka tab <strong>Rekonsiliasi</strong>. Anda dapat memilih untuk langsung menyetujui versi lapangan atau melakukan koreksi manual NIK/KK/Nama.
                </p>
              </div>
            </div>

            <div className="help-step-card">
              <div className="help-step-number">5</div>
              <div className="help-step-content">
                <h3>Ekspor Laporan Akhir & Dokumen Word</h3>
                <p>
                  Setelah semua data tersinkronisasi, buka halaman <strong>Dashboard</strong>. Anda bisa mengunduh Laporan Rekapitulasi Akhir terverifikasi dalam format Excel (menggunakan font Bookman Old Style 12pt & border hitam tipis), atau mencetak draf dokumen Berita Acara & Surat Penyampaian berformat Word (.docx) secara instan dalam satu berkas ZIP.
                </p>
              </div>
            </div>
          </div>
        )}

        {/* Pencarian Global View */}
        {activeTab === 'global-search' && (
          <div className="global-search-page">
            <div className="section-header" style={{ marginBottom: '20px' }}>
              <h2 style={{ fontSize: '1.3rem', fontWeight: '700', color: 'var(--text-primary)' }}>
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
                Pencarian Global
              </h2>
              <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginTop: '4px' }}>
                Pencarian data lintas semua tahap — invers dan terverifikasi.
              </p>
            </div>

            {/* Summary Cards */}
            {globalData && (
              <div className="global-summary-cards">
                <div className="global-summary-card">
                  <div className="global-card-value" style={{ color: '#3B82F6' }}>{globalData.summary.total_alokasi.toLocaleString('id-ID')}</div>
                  <div className="global-card-label">Total Data</div>
                </div>
                <div className="global-summary-card">
                  <div className="global-card-value" style={{ color: '#8B5CF6' }}>{globalData.summary.total_verifikasi.toLocaleString('id-ID')}</div>
                  <div className="global-card-label">Terverifikasi</div>
                </div>
                <div className="global-summary-card">
                  <div className="global-card-value" style={{ color: '#10B981' }}>{globalData.summary.total_lolos.toLocaleString('id-ID')}</div>
                  <div className="global-card-label">Lolos</div>
                </div>
                <div className="global-summary-card">
                  <div className="global-card-value" style={{ color: '#EF4444' }}>{globalData.summary.total_tidak_lolos.toLocaleString('id-ID')}</div>
                  <div className="global-card-label">Tidak Lolos</div>
                </div>
                <div className="global-summary-card">
                  <div className="global-card-value" style={{ color: '#F59E0B' }}>{globalData.summary.total_belum.toLocaleString('id-ID')}</div>
                  <div className="global-card-label">Belum Diverifikasi</div>
                </div>
              </div>
            )}

            {/* Filter Bar */}
            <div className="global-search-filters">
              <div className="global-search-input-wrapper" style={{ flex: '1 1 300px', minWidth: '200px' }}>
                <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }}><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
                <input
                  type="text"
                  className="global-search-input"
                  placeholder="Cari Nama, NIK, atau No KK..."
                  value={globalSearchQuery}
                  onChange={(e) => setGlobalSearchQuery(e.target.value)}
                />
              </div>
              <select className="global-filter-select" value={globalFilterType} onChange={(e) => setGlobalFilterType(e.target.value)}>
                <option value="all">Semua Jenis</option>
                <option value="verified">Terverifikasi</option>
                <option value="invers">Belum Diverifikasi</option>
              </select>
              <select className="global-filter-select" value={globalFilterKab} onChange={(e) => { setGlobalFilterKab(e.target.value); setGlobalFilterKec('ALL'); setGlobalFilterDesa('ALL'); }}>
                <option value="ALL">Semua Kabupaten</option>
                {globalData?.filters?.kabupatens?.map(k => <option key={k} value={k}>{k}</option>)}
              </select>
              <select className="global-filter-select" value={globalFilterKec} onChange={(e) => { setGlobalFilterKec(e.target.value); setGlobalFilterDesa('ALL'); }}>
                <option value="ALL">Semua Kecamatan</option>
                {globalData?.filters?.kecamatans?.map(k => <option key={k} value={k}>{k}</option>)}
              </select>
              <select className="global-filter-select" value={globalFilterDesa} onChange={(e) => setGlobalFilterDesa(e.target.value)}>
                <option value="ALL">Semua Desa</option>
                {globalData?.filters?.desas?.map(d => <option key={d} value={d}>{d}</option>)}
              </select>
              <select className="global-filter-select" value={globalFilterStatus} onChange={(e) => setGlobalFilterStatus(e.target.value)}>
                <option value="ALL">Semua Status</option>
                <option value="LOLOS">Lolos</option>
                <option value="TIDAK_LOLOS">Tidak Lolos</option>
                <option value="BELUM">Belum Diverifikasi</option>
              </select>
              <select className="global-filter-select" value={globalFilterTahap} onChange={(e) => setGlobalFilterTahap(e.target.value)}>
                <option value="ALL">Semua Tahap</option>
                {globalData?.filters?.tahaps?.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
              </select>
              <button
                className="btn btn-primary"
                style={{ whiteSpace: 'nowrap' }}
                onClick={async () => {
                  try {
                    const params = new URLSearchParams();
                    if (globalSearchQuery) params.set('q', globalSearchQuery);
                    if (globalFilterKab !== 'ALL') params.set('kabupaten', globalFilterKab);
                    if (globalFilterKec !== 'ALL') params.set('kecamatan', globalFilterKec);
                    if (globalFilterDesa !== 'ALL') params.set('desa', globalFilterDesa);
                    if (globalFilterStatus !== 'ALL') params.set('status', globalFilterStatus);
                    if (globalFilterTahap !== 'ALL') params.set('tahap', globalFilterTahap);
                    if (globalFilterType !== 'all') params.set('record_type', globalFilterType);
                    const res = await fetch(`${BACKEND_URL}/api/global-search/export?${params}`);
                    if (!res.ok) throw new Error("Gagal export");
                    const blob = await res.blob();
                    const url = window.URL.createObjectURL(blob);
                    const a = document.createElement('a');
                    a.href = url;
                    a.download = `Pencarian_Global_${globalFilterKab !== 'ALL' ? globalFilterKab : 'Semua'}_${globalFilterStatus !== 'ALL' ? globalFilterStatus : 'Semua'}_${new Date().toISOString().slice(0,10)}.xlsx`;
                    a.click();
                    window.URL.revokeObjectURL(url);
                  } catch (err) {
                    showToast(err.message, 'error');
                  }
                }}
              >
                <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '4px', verticalAlign: 'middle' }}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                Export
              </button>
            </div>

            {/* Results Table */}
            <div className="global-search-table-container">
              {globalLoading ? (
                <div style={{ textAlign: 'center', padding: '40px', color: 'var(--text-muted)' }}>
                  <div className="spinner" style={{ width: '32px', height: '32px', margin: '0 auto 12px' }}></div>
                  Memuat data...
                </div>
              ) : globalData && globalData.records.length > 0 ? (
                <table className="global-search-table">
                  <thead>
                    <tr>
                      <th style={{ width: '45px' }}>No</th>
                      <th style={{ width: '100px' }}>Tahap</th>
                      <th>Nama</th>
                      <th>NIK</th>
                      <th>No KK</th>
                      <th>Kabupaten</th>
                      <th>Desa</th>
                      <th style={{ width: '130px' }}>Status</th>
                      <th style={{ width: '130px' }}>Asal</th>
                    </tr>
                  </thead>
                  <tbody>
                    {globalData.records.map((rec, idx) => (
                      <tr key={`${rec.record_type}-${rec.id}`}>
                        <td style={{ textAlign: 'center', color: 'var(--text-muted)' }}>{(globalData.page - 1) * 30 + idx + 1}</td>
                        <td><span className="tahap-badge">{rec.tahap_name}</span></td>
                        <td style={{ fontWeight: '500' }}>{rec.nama}</td>
                        <td className="mono-digit">{rec.no_ktp}</td>
                        <td className="mono-digit">{rec.no_kk}</td>
                        <td>{rec.kabupaten_kota}</td>
                        <td>{rec.desa_kelurahan}</td>
                        <td>
                          <span className={`status-badge ${
                            rec.status === 'LOLOS' ? 'lolos' :
                            rec.status === 'TIDAK LOLOS' ? 'tidak' : 'belum'
                          }`}>
                            {rec.status === 'TIDAK LOLOS' ? 'TIDAK LOLOS' : rec.status}
                          </span>
                        </td>
                        <td>
                          <span className={`asal-badge ${rec.record_type === 'verified' ? 'asal-verified' : 'asal-invers'}`}>
                            {rec.record_type === 'verified' ? 'Terverifikasi' : 'Belum Diverifikasi'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div style={{ textAlign: 'center', padding: '60px 20px', color: 'var(--text-muted)' }}>
                  <svg xmlns="http://www.w3.org/2000/svg" width="48" height="48" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1" strokeLinecap="round" strokeLinejoin="round" style={{ margin: '0 auto 16px', opacity: 0.3 }}><circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line></svg>
                  <p style={{ fontSize: '1rem', fontWeight: '500' }}>Tidak ada data ditemukan</p>
                  <p style={{ fontSize: '0.85rem', marginTop: '4px' }}>Coba ubah filter atau kata kunci pencarian</p>
                </div>
              )}
            </div>

            {/* Pagination */}
            {globalData && globalData.total_pages > 1 && (
              <div className="pagination-bar" style={{ marginTop: '16px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                  Menampilkan {((globalData.page - 1) * 30) + 1}–{Math.min(globalData.page * 30, globalData.total)} dari {globalData.total.toLocaleString('id-ID')} data
                </span>
                <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                  <button className="btn btn-outline" disabled={globalData.page <= 1} onClick={() => { setGlobalPage(globalData.page - 1); }}>
                    ← Sebelumnya
                  </button>
                  <span style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>
                    Hal {globalData.page} / {globalData.total_pages}
                  </span>
                  <button className="btn btn-outline" disabled={globalData.page >= globalData.total_pages} onClick={() => { setGlobalPage(globalData.page + 1); }}>
                    Berikutnya →
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* Rekap Keseluruhan View */}
        {activeTab === 'rekap' && (
          <div className="rekap-keseluruhan-page">
            <div className="section-header" style={{ marginBottom: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <h2 style={{ fontSize: '1.3rem', fontWeight: '700', color: 'var(--text-primary)' }}>
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><line x1="18" y1="20" x2="18" y2="10"></line><line x1="12" y1="20" x2="12" y2="4"></line><line x1="6" y1="20" x2="6" y2="14"></line></svg>Rekap Keseluruhan Semua INVERS
                </h2>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginTop: '4px' }}>
                  Rekapitulasi total alokasi, verifikasi, lolos, tidak lolos, dan belum verifikasi per kabupaten lintas semua tahap. Hanya menghitung Berita Acara yang sudah ditandai "Sudah Terbit".
                </p>
              </div>
              {!rekapLoading && rekapData && rekapData.stages.length > 0 && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                  <button 
                    className="btn btn-secondary btn-sm"
                    onClick={handleExportRekapKeseluruhan}
                    style={{ display: 'flex', alignItems: 'center', gap: '6px', fontWeight: '600' }}
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                    Ekspor Rekap Excel
                  </button>
                  <div className="rekap-progress-legend" style={{ margin: 0, background: 'var(--bg-card)', color: 'var(--text-primary)', padding: '10px 16px', borderRadius: '8px', border: '1px solid var(--border)' }}>
                    <span className="rekap-legend-item"><span className="rekap-legend-dot rekap-seg-lolos"></span> Lolos</span>
                    <span className="rekap-legend-item"><span className="rekap-legend-dot rekap-seg-tidak-lolos"></span> Tidak Lolos</span>
                    <span className="rekap-legend-item"><span className="rekap-legend-dot rekap-seg-belum"></span> Belum Verifikasi</span>
                  </div>
                </div>
              )}
            </div>

            {rekapLoading && (
              <div style={{ textAlign: 'center', padding: '60px 0', color: 'var(--text-muted)' }}>
                <div className="spinner" style={{ margin: '0 auto 16px' }}></div>
                Memuat data rekap keseluruhan...
              </div>
            )}

            {!rekapLoading && rekapData && rekapData.stages.length === 0 && (
              <div style={{ textAlign: 'center', padding: '60px 0', color: 'var(--text-muted)' }}>
                Belum ada data INVERS yang tersedia.
              </div>
            )}

            {!rekapLoading && rekapData && rekapData.stages.length > 0 && (() => {
              const allStages = [...rekapData.stages];
              const murniStages = allStages
                .filter(s => s.stage_type !== 'pengganti')
                .sort((a, b) => (parseInt(a.stage_name.replace(/\D/g, '')) || 0) - (parseInt(b.stage_name.replace(/\D/g, '')) || 0));
              const penggantiStages = allStages
                .filter(s => s.stage_type === 'pengganti')
                .sort((a, b) => (parseInt(a.stage_name.replace(/\D/g, '')) || 0) - (parseInt(b.stage_name.replace(/\D/g, '')) || 0));
              const sortedRekapStages = [...murniStages, ...penggantiStages];

              const renderRekapTable = (groupStages, label) => {
                if (groupStages.length === 0) return null;
                return (
                  <div key={label}>
                    <div className="rekap-section-label">{label}</div>
                    <div className="rekap-scroll-container">
                      <table className="rekap-table-unified">
                        <thead>
                          <tr>
                            <th className="rekap-corner-cell" rowSpan="2" style={{ width: '45px', minWidth: '45px', maxWidth: '45px', left: 0 }}>No</th>
                            <th className="rekap-corner-cell" rowSpan="2" style={{ width: '220px', minWidth: '220px', maxWidth: '220px', left: '45px', textAlign: 'left', borderRight: '2px solid #bfc6d0', whiteSpace: 'normal', wordBreak: 'break-word' }}>Kabupaten / Kota</th>
                            <th colSpan="7" className="rekap-corner-cell rekap-total-group-th" style={{ left: '265px', width: '530px', minWidth: '530px', maxWidth: '530px', borderRight: '2px solid #bfc6d0' }}>
                              REKAP TOTAL
                            </th>
                            {groupStages.map(stage => {
                              const t = stage.totals;
                              const totalBase = t.alokasi || 1;
                              const pctLolos = ((t.lolos / totalBase) * 100).toFixed(1);
                              const pctTidakLolos = ((t.tidak_lolos / totalBase) * 100).toFixed(1);
                              const pctBelum = ((t.belum_verifikasi / totalBase) * 100).toFixed(1);
                              return (
                                <th key={stage.stage_id} colSpan="7" className="rekap-stage-th">
                                  <div className="rekap-header-progress-container" style={{ marginBottom: '5px' }}>
                                    <div className="rekap-progress-bar-mini" title={`${stage.stage_name} — Lolos: ${pctLolos}%, Tidak Lolos: ${pctTidakLolos}%, Belum: ${pctBelum}%`}>
                                      <div className="rekap-progress-segment rekap-seg-lolos" style={{ width: `${pctLolos}%`, height: '100%' }}></div>
                                      <div className="rekap-progress-segment rekap-seg-tidak-lolos" style={{ width: `${pctTidakLolos}%`, height: '100%' }}></div>
                                      <div className="rekap-progress-segment rekap-seg-belum" style={{ width: `${pctBelum}%`, height: '100%' }}></div>
                                    </div>
                                  </div>
                                  <div style={{ fontWeight: 700, fontSize: '0.82rem', letterSpacing: '0.3px' }}>{stage.stage_name.toUpperCase()}</div>
                                </th>
                              );
                            })}
                          </tr>
                          <tr>
                            <th className="rekap-sub-th rekap-sub-total-sticky" style={{ left: '265px', width: '70px', minWidth: '70px', maxWidth: '70px' }}>ALOKASI</th>
                            <th className="rekap-sub-th rekap-sub-total-sticky" style={{ left: '335px', width: '85px', minWidth: '85px', maxWidth: '85px' }}>VERIFIKASI</th>
                            <th className="rekap-sub-th rekap-sub-lolos rekap-sub-total-sticky" style={{ left: '420px', width: '70px', minWidth: '70px', maxWidth: '70px' }}>LOLOS</th>
                            <th className="rekap-sub-th rekap-sub-tidak rekap-sub-total-sticky" style={{ left: '490px', width: '80px', minWidth: '80px', maxWidth: '80px' }}>TIDAK LOLOS</th>
                            <th className="rekap-sub-th rekap-sub-belum rekap-sub-total-sticky" style={{ left: '570px', width: '65px', minWidth: '65px', maxWidth: '65px' }}>BELUM</th>
                            <th className="rekap-sub-th rekap-sub-skdirjen-sudah rekap-sub-total-sticky" style={{ left: '635px', width: '65px', minWidth: '65px', maxWidth: '65px' }}>SUDAH</th>
                            <th className="rekap-sub-th rekap-sub-skdirjen-belum rekap-sub-total-sticky" style={{ left: '700px', width: '65px', minWidth: '65px', maxWidth: '65px', borderRight: '2px solid #bfc6d0' }}>BELUM</th>
                            {groupStages.map(stage => (
                              <React.Fragment key={stage.stage_id}>
                                <th className="rekap-sub-th">ALOKASI</th>
                                <th className="rekap-sub-th">VERIFIKASI</th>
                                <th className="rekap-sub-th rekap-sub-lolos">LOLOS</th>
                                <th className="rekap-sub-th rekap-sub-tidak">TIDAK LOLOS</th>
                                <th className="rekap-sub-th rekap-sub-belum">BELUM</th>
                                <th className="rekap-sub-th rekap-sub-skdirjen-sudah">SUDAH</th>
                                <th className="rekap-sub-th rekap-sub-skdirjen-belum" style={{ borderRight: '2px solid #dee2e6' }}>BELUM</th>
                              </React.Fragment>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {rekapData.all_kabupaten.map((kab, kabIdx) => {
                            const hasAnyData = groupStages.some(stage => {
                              const kd = stage.kabupaten_data[kabIdx];
                              return kd && (kd.alokasi > 0 || kd.verifikasi > 0);
                            });
                            let sumA = 0, sumV = 0, sumL = 0, sumTL = 0, sumB = 0, sumSKS = 0, sumSKB = 0;
                            groupStages.forEach(stage => {
                              const kd = stage.kabupaten_data[kabIdx];
                              if (kd) {
                                sumA += kd.alokasi || 0;
                                sumV += kd.verifikasi || 0;
                                sumL += kd.lolos || 0;
                                sumTL += kd.tidak_lolos || 0;
                                sumB += kd.belum_verifikasi || 0;
                                sumSKS += kd.sk_dirjen_sudah || 0;
                                sumSKB += kd.sk_dirjen_belum || 0;
                              }
                            });
                            return (
                              <tr key={kabIdx} className={!hasAnyData ? 'rekap-row-empty' : ''}>
                                <td className="rekap-frozen-no" style={{ width: '45px', minWidth: '45px', maxWidth: '45px', left: 0 }}>{kabIdx + 1}</td>
                                <td className="rekap-frozen-kab" style={{ width: '220px', minWidth: '220px', maxWidth: '220px', left: '45px', whiteSpace: 'normal', wordBreak: 'break-word' }}>{kab}</td>
                                <td className="rekap-frozen-total" style={{ left: '265px', width: '70px', minWidth: '70px', maxWidth: '70px' }}><button className="rekap-link" onClick={() => navigateToData('invers', { kab })}>{sumA || '-'}</button></td>
                                <td className="rekap-frozen-total" style={{ left: '335px', width: '85px', minWidth: '85px', maxWidth: '85px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { kab })}>{sumV || '-'}</button></td>
                                <td className="rekap-frozen-total rekap-cell-lolos" style={{ left: '420px', width: '70px', minWidth: '70px', maxWidth: '70px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { kab, status: 'LOLOS' })}>{sumL || '-'}</button></td>
                                <td className="rekap-frozen-total rekap-cell-tidak" style={{ left: '490px', width: '80px', minWidth: '80px', maxWidth: '80px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { kab, status: 'TIDAK_LOLOS' })}>{sumTL || '-'}</button></td>
                                <td className="rekap-frozen-total rekap-cell-belum" style={{ left: '570px', width: '65px', minWidth: '65px', maxWidth: '65px' }}><button className="rekap-link" onClick={() => navigateToData('invers', { kab, status: 'BELUM' })}>{sumB || '-'}</button></td>
                                <td className="rekap-frozen-total rekap-cell-skdirjen" style={{ left: '635px', width: '65px', minWidth: '65px', maxWidth: '65px' }}><span className="sk-dirjen-badge sudah">{sumSKS || '-'}</span></td>
                                <td className="rekap-frozen-total rekap-cell-skdirjen" style={{ left: '700px', width: '65px', minWidth: '65px', maxWidth: '65px', borderRight: '2px solid #bfc6d0' }}><span className="sk-dirjen-badge belum">{sumSKB || '-'}</span></td>
                                {groupStages.map(stage => {
                                  const kd = stage.kabupaten_data[kabIdx];
                                  const a = kd?.alokasi || 0;
                                  const v = kd?.verifikasi || 0;
                                  const l = kd?.lolos || 0;
                                  const tl = kd?.tidak_lolos || 0;
                                  const b = kd?.belum_verifikasi || 0;
                                  const sks = kd?.sk_dirjen_sudah || 0;
                                  const skb = kd?.sk_dirjen_belum || 0;
                                  return (
                                    <React.Fragment key={stage.stage_id}>
                                      <td className="rekap-cell"><button className="rekap-link" onClick={() => navigateToData('invers', { kab, tahap: stage.stage_id })}>{a || '-'}</button></td>
                                      <td className="rekap-cell"><button className="rekap-link" onClick={() => navigateToData('verified', { kab, tahap: stage.stage_id })}>{v || '-'}</button></td>
                                      <td className="rekap-cell rekap-cell-lolos"><button className="rekap-link" onClick={() => navigateToData('verified', { kab, tahap: stage.stage_id, status: 'LOLOS' })}>{l || '-'}</button></td>
                                      <td className="rekap-cell rekap-cell-tidak"><button className="rekap-link" onClick={() => navigateToData('verified', { kab, tahap: stage.stage_id, status: 'TIDAK_LOLOS' })}>{tl || '-'}</button></td>
                                      <td className="rekap-cell rekap-cell-belum"><button className="rekap-link" onClick={() => navigateToData('invers', { kab, tahap: stage.stage_id, status: 'BELUM' })}>{b || '-'}</button></td>
                                   <td className="rekap-cell rekap-cell-skdirjen"><span className="sk-dirjen-badge sudah">{sks || '-'}</span></td>
                                   <td className="rekap-cell rekap-cell-skdirjen" style={{ borderRight: '2px solid #dee2e6' }}><span className="sk-dirjen-badge belum">{skb || '-'}</span></td>
                                 </React.Fragment>
                                  );
                                })}
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr className="rekap-footer-row">
                            <td className="rekap-footer-frozen" style={{ left: 0, width: '45px', minWidth: '45px', maxWidth: '45px' }}></td>
                            <td className="rekap-footer-frozen" style={{ left: '45px', width: '220px', minWidth: '220px', maxWidth: '220px', fontWeight: 700, textAlign: 'left', borderRight: '2px solid #bfc6d0' }}>TOTAL</td>
                            {(() => {
                              let gA = 0, gV = 0, gL = 0, gTL = 0, gB = 0, gSKS = 0, gSKB = 0;
                              groupStages.forEach(s => {
                                const t = s.totals;
                                gA += t.alokasi || 0;
                                gV += t.verifikasi || 0;
                                gL += t.lolos || 0;
                                gTL += t.tidak_lolos || 0;
                                gB += t.belum_verifikasi || 0;
                                gSKS += t.sk_dirjen_sudah || 0;
                                gSKB += t.sk_dirjen_belum || 0;
                              });
                               return (
                                <>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky" style={{ left: '265px', width: '70px', minWidth: '70px', maxWidth: '70px' }}><button className="rekap-link" onClick={() => navigateToData('invers')}>{gA}</button></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky" style={{ left: '335px', width: '85px', minWidth: '85px', maxWidth: '85px' }}><button className="rekap-link" onClick={() => navigateToData('verified')}>{gV}</button></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky rekap-cell-lolos" style={{ left: '420px', width: '70px', minWidth: '70px', maxWidth: '70px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { status: 'LOLOS' })}>{gL}</button></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky rekap-cell-tidak" style={{ left: '490px', width: '80px', minWidth: '80px', maxWidth: '80px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { status: 'TIDAK_LOLOS' })}>{gTL}</button></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky rekap-cell-belum" style={{ left: '570px', width: '65px', minWidth: '65px', maxWidth: '65px' }}><button className="rekap-link" onClick={() => navigateToData('invers', { status: 'BELUM' })}>{gB}</button></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky rekap-cell-skdirjen" style={{ left: '635px', width: '65px', minWidth: '65px', maxWidth: '65px' }}><span className="sk-dirjen-badge sudah">{gSKS}</span></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky rekap-cell-skdirjen" style={{ left: '700px', width: '65px', minWidth: '65px', maxWidth: '65px', borderRight: '2px solid #bfc6d0' }}><span className="sk-dirjen-badge belum">{gSKB}</span></td>
                                </>
                               );
                             })()}
                             {groupStages.map(stage => {
                               const t = stage.totals;
                               return (
                                 <React.Fragment key={stage.stage_id}>
                                   <td className="rekap-footer-cell"><button className="rekap-link" onClick={() => navigateToData('invers', { tahap: stage.stage_id })}>{t.alokasi}</button></td>
                                   <td className="rekap-footer-cell"><button className="rekap-link" onClick={() => navigateToData('verified', { tahap: stage.stage_id })}>{t.verifikasi}</button></td>
                                   <td className="rekap-footer-cell rekap-cell-lolos"><button className="rekap-link" onClick={() => navigateToData('verified', { tahap: stage.stage_id, status: 'LOLOS' })}>{t.lolos}</button></td>
                                   <td className="rekap-footer-cell rekap-cell-tidak"><button className="rekap-link" onClick={() => navigateToData('verified', { tahap: stage.stage_id, status: 'TIDAK_LOLOS' })}>{t.tidak_lolos}</button></td>
                                   <td className="rekap-footer-cell rekap-cell-belum"><button className="rekap-link" onClick={() => navigateToData('invers', { tahap: stage.stage_id, status: 'BELUM' })}>{t.belum_verifikasi}</button></td>
                                   <td className="rekap-footer-cell rekap-cell-skdirjen"><span className="sk-dirjen-badge sudah">{t.sk_dirjen_sudah || '-'}</span></td>
                                   <td className="rekap-footer-cell rekap-cell-skdirjen" style={{ borderRight: '2px solid #dee2e6' }}><span className="sk-dirjen-badge belum">{t.sk_dirjen_belum || '-'}</span></td>
                                 </React.Fragment>
                               );
                             })}
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                );
              };

              return (
                <>
                  {renderRekapTable(murniStages, 'Tabel Invers Murni')}
                  {renderRekapTable(penggantiStages, 'Tabel Invers Pengganti')}
                  {renderRekapTable(sortedRekapStages, 'Tabel Keseluruhan')}
                </>
              );
            })()}
          </div>
        )}

        {/* Rekap Unggahan View */}
        {activeTab === 'rekap-unggahan' && (
          <div className="rekap-keseluruhan-page">
            <div className="section-header" style={{ marginBottom: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <h2 style={{ fontSize: '1.3rem', fontWeight: '700', color: 'var(--text-primary)' }}>
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-5 14H7v-2h7v2zm3-4H7v-2h10v2zm0-4H7V7h10v2z"/></svg>Rekap Unggahan
                </h2>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginTop: '4px' }}>
                  Rekapitulasi jumlah CPB per kabupaten pada setiap tahap INVERS berdasarkan data yang sudah diunggah.
                </p>
              </div>
              {!rekapLoading && rekapUnggahanData && rekapUnggahanData.stages.length > 0 && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                  <button 
                    className="btn btn-secondary btn-sm"
                    onClick={handleExportRekapKeseluruhan}
                    style={{ display: 'flex', alignItems: 'center', gap: '6px', fontWeight: '600' }}
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                    Ekspor Rekap Excel
                  </button>
                </div>
              )}
            </div>

            {rekapLoading && (
              <div style={{ textAlign: 'center', padding: '60px 0', color: 'var(--text-muted)' }}>
                <div className="spinner" style={{ margin: '0 auto 16px' }}></div>
                Memuat data rekap unggahan...
              </div>
            )}

            {!rekapLoading && rekapUnggahanData && rekapUnggahanData.stages.length === 0 && (
              <div style={{ textAlign: 'center', padding: '60px 0', color: 'var(--text-muted)' }}>
                Belum ada data INVERS yang tersedia.
              </div>
            )}

            {!rekapLoading && rekapUnggahanData && rekapUnggahanData.stages.length > 0 && (() => {
              const allStages = [...rekapUnggahanData.stages];
              const murniStages = allStages
                .filter(s => s.stage_type !== 'pengganti')
                .sort((a, b) => (parseInt(a.stage_name.replace(/\D/g, '')) || 0) - (parseInt(b.stage_name.replace(/\D/g, '')) || 0));
              const penggantiStages = allStages
                .filter(s => s.stage_type === 'pengganti')
                .sort((a, b) => (parseInt(a.stage_name.replace(/\D/g, '')) || 0) - (parseInt(b.stage_name.replace(/\D/g, '')) || 0));
              const sortedRekapStages = [...murniStages, ...penggantiStages];

              const renderRekapTable = (groupStages, label) => {
                if (groupStages.length === 0) return null;
                return (
                  <div key={label}>
                    <div className="rekap-section-label">{label}</div>
                    <div className="rekap-scroll-container">
                      <table className="rekap-table-unified">
                        <thead>
                          <tr>
                            <th className="rekap-corner-cell" rowSpan="2" style={{ width: '45px', minWidth: '45px', maxWidth: '45px', left: 0 }}>No</th>
                            <th className="rekap-corner-cell" rowSpan="2" style={{ width: '220px', minWidth: '220px', maxWidth: '220px', left: '45px', textAlign: 'left', borderRight: '2px solid #bfc6d0', whiteSpace: 'normal', wordBreak: 'break-word' }}>Kabupaten / Kota</th>
                            <th colSpan="5" className="rekap-corner-cell rekap-total-group-th" style={{ left: '265px', width: '390px', minWidth: '390px', maxWidth: '390px', borderRight: '2px solid #bfc6d0' }}>
                              REKAP TOTAL
                            </th>
                            {groupStages.map(stage => {
                              const t = stage.totals;
                              const totalBase = t.alokasi || 1;
                              const pctLolos = ((t.lolos / totalBase) * 100).toFixed(1);
                              const pctTidakLolos = ((t.tidak_lolos / totalBase) * 100).toFixed(1);
                              const pctBelum = ((t.belum_verifikasi / totalBase) * 100).toFixed(1);
                              return (
                                <th key={stage.stage_id} colSpan="5" className="rekap-stage-th">
                                  <div className="rekap-header-progress-container" style={{ marginBottom: '5px' }}>
                                    <div className="rekap-progress-bar-mini" title={`${stage.stage_name} — Lolos: ${pctLolos}%, Tidak Lolos: ${pctTidakLolos}%, Belum: ${pctBelum}%`}>
                                      <div className="rekap-progress-segment rekap-seg-lolos" style={{ width: `${pctLolos}%`, height: '100%' }}></div>
                                      <div className="rekap-progress-segment rekap-seg-tidak-lolos" style={{ width: `${pctTidakLolos}%`, height: '100%' }}></div>
                                      <div className="rekap-progress-segment rekap-seg-belum" style={{ width: `${pctBelum}%`, height: '100%' }}></div>
                                    </div>
                                  </div>
                                  <div style={{ fontWeight: 700, fontSize: '0.82rem', letterSpacing: '0.3px' }}>{stage.stage_name.toUpperCase()}</div>
                                </th>
                              );
                            })}
                          </tr>
                          <tr>
                            <th className="rekap-sub-th rekap-sub-total-sticky" style={{ left: '265px', width: '70px', minWidth: '70px', maxWidth: '70px' }}>ALOKASI</th>
                            <th className="rekap-sub-th rekap-sub-total-sticky" style={{ left: '335px', width: '85px', minWidth: '85px', maxWidth: '85px' }}>VERIFIKASI</th>
                            <th className="rekap-sub-th rekap-sub-lolos rekap-sub-total-sticky" style={{ left: '420px', width: '70px', minWidth: '70px', maxWidth: '70px' }}>LOLOS</th>
                            <th className="rekap-sub-th rekap-sub-tidak rekap-sub-total-sticky" style={{ left: '490px', width: '90px', minWidth: '90px', maxWidth: '90px' }}>TIDAK LOLOS</th>
                            <th className="rekap-sub-th rekap-sub-belum rekap-sub-total-sticky" style={{ left: '580px', width: '75px', minWidth: '75px', maxWidth: '75px', borderRight: '2px solid #bfc6d0' }}>BELUM</th>
                            {groupStages.map(stage => (
                              <React.Fragment key={stage.stage_id}>
                                <th className="rekap-sub-th">ALOKASI</th>
                                <th className="rekap-sub-th">VERIFIKASI</th>
                                <th className="rekap-sub-th rekap-sub-lolos">LOLOS</th>
                                <th className="rekap-sub-th rekap-sub-tidak">TIDAK LOLOS</th>
                                <th className="rekap-sub-th rekap-sub-belum" style={{ borderRight: '2px solid #dee2e6' }}>BELUM</th>
                              </React.Fragment>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {rekapUnggahanData.all_kabupaten.map((kab, kabIdx) => {
                            const hasAnyData = groupStages.some(stage => {
                              const kd = stage.kabupaten_data[kabIdx];
                              return kd && (kd.alokasi > 0 || kd.verifikasi > 0);
                            });
                            let sumA = 0, sumV = 0, sumL = 0, sumTL = 0, sumB = 0;
                            groupStages.forEach(stage => {
                              const kd = stage.kabupaten_data[kabIdx];
                              if (kd) {
                                sumA += kd.alokasi || 0;
                                sumV += kd.verifikasi || 0;
                                sumL += kd.lolos || 0;
                                sumTL += kd.tidak_lolos || 0;
                                sumB += kd.belum_verifikasi || 0;
                              }
                            });
                            return (
                              <tr key={kabIdx} className={!hasAnyData ? 'rekap-row-empty' : ''}>
                                <td className="rekap-frozen-no" style={{ width: '45px', minWidth: '45px', maxWidth: '45px', left: 0 }}>{kabIdx + 1}</td>
                                <td className="rekap-frozen-kab" style={{ width: '220px', minWidth: '220px', maxWidth: '220px', left: '45px', whiteSpace: 'normal', wordBreak: 'break-word' }}>{kab}</td>
                                <td className="rekap-frozen-total" style={{ left: '265px', width: '70px', minWidth: '70px', maxWidth: '70px' }}><button className="rekap-link" onClick={() => navigateToData('invers', { kab })}>{sumA || '-'}</button></td>
                                <td className="rekap-frozen-total" style={{ left: '335px', width: '85px', minWidth: '85px', maxWidth: '85px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { kab })}>{sumV || '-'}</button></td>
                                <td className="rekap-frozen-total rekap-cell-lolos" style={{ left: '420px', width: '70px', minWidth: '70px', maxWidth: '70px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { kab, status: 'LOLOS' })}>{sumL || '-'}</button></td>
                                <td className="rekap-frozen-total rekap-cell-tidak" style={{ left: '490px', width: '90px', minWidth: '90px', maxWidth: '90px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { kab, status: 'TIDAK_LOLOS' })}>{sumTL || '-'}</button></td>
                                <td className="rekap-frozen-total rekap-cell-belum" style={{ left: '580px', width: '75px', minWidth: '75px', maxWidth: '75px', borderRight: '2px solid #bfc6d0' }}><button className="rekap-link" onClick={() => navigateToData('invers', { kab, status: 'BELUM' })}>{sumB || '-'}</button></td>
                                {groupStages.map(stage => {
                                  const kd = stage.kabupaten_data[kabIdx];
                                  const a = kd?.alokasi || 0;
                                  const v = kd?.verifikasi || 0;
                                  const l = kd?.lolos || 0;
                                  const tl = kd?.tidak_lolos || 0;
                                  const b = kd?.belum_verifikasi || 0;
                                  return (
                                    <React.Fragment key={stage.stage_id}>
                                      <td className="rekap-cell"><button className="rekap-link" onClick={() => navigateToData('invers', { kab, tahap: stage.stage_id })}>{a || '-'}</button></td>
                                      <td className="rekap-cell"><button className="rekap-link" onClick={() => navigateToData('verified', { kab, tahap: stage.stage_id })}>{v || '-'}</button></td>
                                      <td className="rekap-cell rekap-cell-lolos"><button className="rekap-link" onClick={() => navigateToData('verified', { kab, tahap: stage.stage_id, status: 'LOLOS' })}>{l || '-'}</button></td>
                                      <td className="rekap-cell rekap-cell-tidak"><button className="rekap-link" onClick={() => navigateToData('verified', { kab, tahap: stage.stage_id, status: 'TIDAK_LOLOS' })}>{tl || '-'}</button></td>
                                      <td className="rekap-cell rekap-cell-belum" style={{ borderRight: '2px solid #dee2e6' }}><button className="rekap-link" onClick={() => navigateToData('invers', { kab, tahap: stage.stage_id, status: 'BELUM' })}>{b || '-'}</button></td>
                                    </React.Fragment>
                                  );
                                })}
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr className="rekap-footer-row">
                            <td className="rekap-footer-frozen" style={{ left: 0, width: '45px', minWidth: '45px', maxWidth: '45px' }}></td>
                            <td className="rekap-footer-frozen" style={{ left: '45px', width: '220px', minWidth: '220px', maxWidth: '220px', fontWeight: 700, textAlign: 'left', borderRight: '2px solid #bfc6d0' }}>TOTAL</td>
                            {(() => {
                              let gA = 0, gV = 0, gL = 0, gTL = 0, gB = 0;
                              groupStages.forEach(s => {
                                const t = s.totals;
                                gA += t.alokasi || 0;
                                gV += t.verifikasi || 0;
                                gL += t.lolos || 0;
                                gTL += t.tidak_lolos || 0;
                                gB += t.belum_verifikasi || 0;
                              });
                               return (
                                <>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky" style={{ left: '265px', width: '70px', minWidth: '70px', maxWidth: '70px' }}><button className="rekap-link" onClick={() => navigateToData('invers')}>{gA}</button></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky" style={{ left: '335px', width: '85px', minWidth: '85px', maxWidth: '85px' }}><button className="rekap-link" onClick={() => navigateToData('verified')}>{gV}</button></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky rekap-cell-lolos" style={{ left: '420px', width: '70px', minWidth: '70px', maxWidth: '70px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { status: 'LOLOS' })}>{gL}</button></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky rekap-cell-tidak" style={{ left: '490px', width: '90px', minWidth: '90px', maxWidth: '90px' }}><button className="rekap-link" onClick={() => navigateToData('verified', { status: 'TIDAK_LOLOS' })}>{gTL}</button></td>
                                  <td className="rekap-footer-frozen rekap-footer-total-sticky rekap-cell-belum" style={{ left: '580px', width: '75px', minWidth: '75px', maxWidth: '75px', borderRight: '2px solid #bfc6d0' }}><button className="rekap-link" onClick={() => navigateToData('invers', { status: 'BELUM' })}>{gB}</button></td>
                                </>
                               );
                             })()}
                             {groupStages.map(stage => {
                               const t = stage.totals;
                               return (
                                 <React.Fragment key={stage.stage_id}>
                                   <td className="rekap-footer-cell"><button className="rekap-link" onClick={() => navigateToData('invers', { tahap: stage.stage_id })}>{t.alokasi}</button></td>
                                   <td className="rekap-footer-cell"><button className="rekap-link" onClick={() => navigateToData('verified', { tahap: stage.stage_id })}>{t.verifikasi}</button></td>
                                   <td className="rekap-footer-cell rekap-cell-lolos"><button className="rekap-link" onClick={() => navigateToData('verified', { tahap: stage.stage_id, status: 'LOLOS' })}>{t.lolos}</button></td>
                                   <td className="rekap-footer-cell rekap-cell-tidak"><button className="rekap-link" onClick={() => navigateToData('verified', { tahap: stage.stage_id, status: 'TIDAK_LOLOS' })}>{t.tidak_lolos}</button></td>
                                   <td className="rekap-footer-cell rekap-cell-belum" style={{ borderRight: '2px solid #dee2e6' }}><button className="rekap-link" onClick={() => navigateToData('invers', { tahap: stage.stage_id, status: 'BELUM' })}>{t.belum_verifikasi}</button></td>
                                 </React.Fragment>
                               );
                             })}
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                );
              };

              return (
                <>
                  {renderRekapTable(murniStages, 'Tabel Invers Murni')}
                  {renderRekapTable(penggantiStages, 'Tabel Invers Pengganti')}
                  {renderRekapTable(sortedRekapStages, 'Tabel Keseluruhan')}
                </>
              );
            })()}
          </div>
        )}

        {/* Rekap Batch Berita Acara View */}
        {activeTab === 'rekap-batch-ba' && (
          <div className="rekap-keseluruhan-page">
            <div className="section-header" style={{ marginBottom: '20px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <div>
                <h2 style={{ fontSize: '1.3rem', fontWeight: '700', color: 'var(--text-primary)' }}>
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "6px", verticalAlign: "middle" }}><path d="M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 16H8v-2h8v2zm0-4H8v-2h8v2zm-3-5V3.5L18.5 9H13z"/></svg>Rekap Batch Berita Acara
                </h2>
                <p style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginTop: '4px' }}>
                  Rekapitulasi verifikasi per Batch Berita Acara di masing-masing Tahap INVERS. Hanya menampilkan batch Berita Acara yang sudah ditandai "Sudah Terbit".
                </p>
              </div>
              {!rekapBatchLoading && rekapBatchData && rekapBatchData.stages.length > 0 && (
                <div style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                  <button 
                    className="btn btn-secondary btn-sm"
                    onClick={handleExportRekapBatchBA}
                    style={{ display: 'flex', alignItems: 'center', gap: '6px', fontWeight: '600' }}
                  >
                    <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                    Ekspor Rekap Excel
                  </button>
                </div>
              )}
            </div>

            {rekapBatchLoading && (
              <div style={{ textAlign: 'center', padding: '60px 0', color: 'var(--text-muted)' }}>
                <div className="spinner" style={{ margin: '0 auto 16px' }}></div>
                Memuat data rekap batch...
              </div>
            )}

            {!rekapBatchLoading && rekapBatchData && rekapBatchData.stages.length === 0 && (
              <div style={{ textAlign: 'center', padding: '60px 0', color: 'var(--text-muted)' }}>
                Belum ada data batch Berita Acara yang tersedia.
              </div>
            )}

            {!rekapBatchLoading && rekapBatchData && rekapBatchData.stages.length > 0 && (() => {
              const allStages = [...rekapBatchData.stages];
              const murniStages = allStages
                .filter(s => s.stage_type !== 'pengganti')
                .sort((a, b) => (parseInt(a.stage_name.replace(/\D/g, '')) || 0) - (parseInt(b.stage_name.replace(/\D/g, '')) || 0));
              const penggantiStages = allStages
                .filter(s => s.stage_type === 'pengganti')
                .sort((a, b) => (parseInt(a.stage_name.replace(/\D/g, '')) || 0) - (parseInt(b.stage_name.replace(/\D/g, '')) || 0));

              const renderRekapBatchTable = (groupStages, label) => {
                if (groupStages.length === 0) return null;
                return (
                  <div key={label} style={{ marginBottom: '40px' }}>
                    <div className="rekap-section-label" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span>{label}</span>
                    </div>
                    <div className="rekap-scroll-container" style={{ border: '1px solid var(--border)', borderRadius: '8px', overflow: 'auto' }}>
                      <table className="rekap-table-unified">
                        <thead>
                          {/* Row 1: Corner + Stages */}
                          <tr>
                            <th className="rekap-batch-corner" rowSpan="3" style={{ left: 0, width: '45px', minWidth: '45px', maxWidth: '45px' }}>No</th>
                            <th className="rekap-batch-corner rekap-batch-border-stage" rowSpan="3" style={{ left: '45px', width: '220px', minWidth: '220px', maxWidth: '220px', textAlign: 'left' }}>Kabupaten / Kota</th>
                            {groupStages.map(stage => {
                              const batchCount = stage.batches.length;
                              if (batchCount === 0) return null;
                              return (
                                <th 
                                  key={stage.stage_id} 
                                  colSpan={batchCount * 3} 
                                  className="rekap-batch-th-level1 rekap-batch-border-stage"
                                >
                                  {stage.stage_name.toUpperCase()}
                                </th>
                              );
                            })}
                          </tr>
                          {/* Row 2: Batches */}
                          <tr>
                            {groupStages.map(stage => 
                              stage.batches.map((batch, bIdx) => {
                                const isLastBatchInStage = bIdx === stage.batches.length - 1;
                                const borderClass = isLastBatchInStage ? 'rekap-batch-border-stage' : 'rekap-batch-border-ba';
                                return (
                                  <th 
                                    key={batch.batch_id} 
                                    colSpan="3" 
                                    className={`rekap-batch-th-level2 ${borderClass}`}
                                  >
                                    <div style={{ fontWeight: 700, fontSize: '0.8rem' }}>{batch.batch_name.toUpperCase()}</div>
                                    <div style={{ fontWeight: 500, fontSize: '0.66rem', marginTop: '2px', opacity: 0.85 }}>
                                      No: {batch.nomor_ba || '—'}
                                    </div>
                                    <div style={{ fontWeight: 500, fontSize: '0.66rem', opacity: 0.85 }}>
                                      Tgl: {batch.tanggal_ba || '—'}
                                    </div>
                                  </th>
                                );
                              })
                            )}
                          </tr>
                          {/* Row 3: Metrics */}
                          <tr>
                            {groupStages.map(stage => 
                              stage.batches.map((batch, bIdx) => {
                                const isLastBatchInStage = bIdx === stage.batches.length - 1;
                                const borderClass = isLastBatchInStage ? 'rekap-batch-border-stage' : 'rekap-batch-border-ba';
                                return (
                                  <React.Fragment key={batch.batch_id}>
                                    <th className="rekap-batch-th-level3 rekap-batch-sub-verif">VERIFIKASI</th>
                                    <th className="rekap-batch-th-level3 rekap-batch-sub-lolos">LOLOS</th>
                                    <th className={`rekap-batch-th-level3 rekap-batch-sub-tidak ${borderClass}`}>TIDAK LOLOS</th>
                                  </React.Fragment>
                                );
                              })
                            )}
                          </tr>
                        </thead>
                        <tbody>
                          {rekapBatchData.all_kabupaten.map((kab, kabIdx) => {
                            const hasAnyData = groupStages.some(stage => 
                              stage.batches.some(batch => {
                                const kd = batch.kabupaten_data[kabIdx];
                                return kd && kd.verifikasi > 0;
                              })
                            );
                            return (
                              <tr key={kabIdx} className={!hasAnyData ? 'rekap-row-empty' : ''}>
                                <td className="rekap-frozen-no" style={{ width: '45px', minWidth: '45px', maxWidth: '45px', left: 0 }}>{kabIdx + 1}</td>
                                <td className="rekap-frozen-kab rekap-batch-border-stage" style={{ width: '220px', minWidth: '220px', maxWidth: '220px', left: '45px' }}>{kab}</td>
                                {groupStages.map(stage => 
                                  stage.batches.map((batch, bIdx) => {
                                    const isLastBatchInStage = bIdx === stage.batches.length - 1;
                                    const borderClass = isLastBatchInStage ? 'rekap-batch-border-stage' : 'rekap-batch-border-ba';
                                    const kd = batch.kabupaten_data[kabIdx];
                                    const v = kd?.verifikasi || 0;
                                    const l = kd?.lolos || 0;
                                    const tl = kd?.tidak_lolos || 0;
                                    return (
                                      <React.Fragment key={batch.batch_id}>
                                        <td className="rekap-cell rekap-batch-cell-verif">
                                          {v > 0 ? (
                                            <button className="rekap-link" onClick={() => navigateToData('verified', { kab, tahap: stage.stage_id })}>
                                              {v}
                                            </button>
                                          ) : '-'}
                                        </td>
                                        <td className="rekap-cell rekap-batch-cell-lolos">
                                          {l > 0 ? (
                                            <button className="rekap-link" onClick={() => navigateToData('verified', { kab, tahap: stage.stage_id, status: 'LOLOS' })}>
                                              {l}
                                            </button>
                                          ) : '-'}
                                        </td>
                                        <td className={`rekap-cell rekap-batch-cell-tidak ${borderClass}`}>
                                          {tl > 0 ? (
                                            <button className="rekap-link" onClick={() => navigateToData('verified', { kab, tahap: stage.stage_id, status: 'TIDAK_LOLOS' })}>
                                              {tl}
                                            </button>
                                          ) : '-'}
                                        </td>
                                      </React.Fragment>
                                    );
                                  })
                                )}
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr className="rekap-footer-row">
                            <td className="rekap-footer-frozen" style={{ left: 0, width: '45px', minWidth: '45px', maxWidth: '45px' }}></td>
                            <td className="rekap-footer-frozen rekap-batch-border-stage" style={{ left: '45px', width: '220px', minWidth: '220px', maxWidth: '220px', fontWeight: 700, textAlign: 'left' }}>TOTAL</td>
                            {groupStages.map(stage => 
                              stage.batches.map((batch, bIdx) => {
                                const isLastBatchInStage = bIdx === stage.batches.length - 1;
                                const borderClass = isLastBatchInStage ? 'rekap-batch-border-stage' : 'rekap-batch-border-ba';
                                const t = batch.totals;
                                return (
                                  <React.Fragment key={batch.batch_id}>
                                    <td className="rekap-footer-cell rekap-batch-cell-verif">
                                      {t.verifikasi > 0 ? (
                                        <button className="rekap-link" onClick={() => navigateToData('verified', { tahap: stage.stage_id })}>{t.verifikasi}</button>
                                      ) : 0}
                                    </td>
                                    <td className="rekap-footer-cell rekap-batch-cell-lolos">
                                      {t.lolos > 0 ? (
                                        <button className="rekap-link" onClick={() => navigateToData('verified', { tahap: stage.stage_id, status: 'LOLOS' })}>{t.lolos}</button>
                                      ) : 0}
                                    </td>
                                    <td className={`rekap-footer-cell rekap-batch-cell-tidak ${borderClass}`}>
                                      {t.tidak_lolos > 0 ? (
                                        <button className="rekap-link" onClick={() => navigateToData('verified', { tahap: stage.stage_id, status: 'TIDAK_LOLOS' })}>{t.tidak_lolos}</button>
                                      ) : 0}
                                    </td>
                                  </React.Fragment>
                                );
                              })
                            )}
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </div>
                );
              };

              return (
                <>
                  {renderRekapBatchTable(murniStages, 'Tabel Rekap Batch Murni')}
                  {renderRekapBatchTable(penggantiStages, 'Tabel Rekap Batch Pengganti')}
                </>
              );
            })()}
          </div>
        )}

      {/* Rename Modal (Stage & Batch) */}
      {showRenameModal && (
        <div className="modal-overlay" onClick={() => setShowRenameModal(false)}>
          <div className="modal-content" style={{ maxWidth: '440px' }} onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: "8px", verticalAlign: "middle" }}><path d="M12 20h9"></path><path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"></path></svg>
                {renameTarget.type === 'stage' ? 'Ubah Nama Tahap INVERS' : 'Ubah Nama Berita Acara / Batch'}
              </h3>
              <button className="modal-close" onClick={() => setShowRenameModal(false)}>&times;</button>
            </div>
            <div className="modal-body" style={{ padding: '20px' }}>
              <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 600, marginBottom: '8px', color: 'var(--text-primary)' }}>
                {renameTarget.type === 'stage' ? 'Nama Tahap Baru:' : 'Nama Berita Acara / Batch Baru:'}
              </label>
              <input
                type="text"
                className="search-input"
                style={{ width: '100%', padding: '10px 14px', fontSize: '0.9rem', borderRadius: '6px', border: '1px solid var(--border)' }}
                value={newRenameName}
                onChange={(e) => setNewRenameName(e.target.value)}
                onKeyDown={(e) => { if (e.key === 'Enter') handleSaveRename(); }}
                autoFocus
              />
              <p style={{ fontSize: '0.78rem', color: 'var(--text-muted)', marginTop: '10px', lineHeight: 1.4 }}>
                Perubahan nama ini akan diperbarui secara otomatis di seluruh antarmuka, laporan rekapitulasi, cetak dokumen, dan berkas ekspor Excel.
              </p>
            </div>
            <div className="modal-footer" style={{ padding: '12px 20px', display: 'flex', justifyContent: 'flex-end', gap: '10px', borderTop: '1px solid var(--border)' }}>
              <button className="btn btn-secondary" onClick={() => setShowRenameModal(false)} disabled={renameLoading}>
                Batal
              </button>
              <button className="btn btn-primary" onClick={handleSaveRename} disabled={renameLoading}>
                {renameLoading ? 'Menyimpan...' : 'Simpan Perubahan'}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Excel Preview Modal */}
      {showPreviewModal && previewData && (
        <div className="modal-overlay" onClick={closePreview}>
          <div className="modal-content" style={{ maxWidth: '1100px', maxHeight: '90vh' }} onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>
                Preview Excel — {previewData.batch_name}
                <span style={{ fontSize: '0.75em', fontWeight: 'normal', color: '#888', marginLeft: '10px' }}>
                  {previewData.stage_name}
                </span>
              </h3>
              <button className="modal-close" onClick={closePreview}>&times;</button>
            </div>
            <div className="modal-body" style={{ padding: 0, display: 'flex', flexDirection: 'column', height: '75vh' }}>
              {/* Tabs */}
              <div style={{ display: 'flex', borderBottom: '2px solid #e2e8f0', background: '#f8fafc' }}>
                <button
                  onClick={() => setPreviewTab('lolos')}
                  style={{
                    flex: 1, padding: '12px', border: 'none', cursor: 'pointer', fontWeight: 'bold', fontSize: '14px',
                    borderBottom: previewTab === 'lolos' ? '3px solid #16a34a' : '3px solid transparent',
                    background: previewTab === 'lolos' ? '#dcfce7' : 'transparent',
                    color: previewTab === 'lolos' ? '#166534' : '#64748b'
                  }}
                >
                  Lamp. II.A — CPB Unit (Lolos) ({filterPreviewRecords(previewData.lolos_records).length})
                </button>
                <button
                  onClick={() => setPreviewTab('tidak_lolos')}
                  style={{
                    flex: 1, padding: '12px', border: 'none', cursor: 'pointer', fontWeight: 'bold', fontSize: '14px',
                    borderBottom: previewTab === 'tidak_lolos' ? '3px solid #dc2626' : '3px solid transparent',
                    background: previewTab === 'tidak_lolos' ? '#fee2e2' : 'transparent',
                    color: previewTab === 'tidak_lolos' ? '#991b1b' : '#64748b'
                  }}
                >
                  Lamp. III.A — Tidak Lolos ({filterPreviewRecords(previewData.tidak_lolos_records).length})
                </button>
              </div>

              {/* Search Bar */}
              <div style={{ padding: '10px 16px', borderBottom: '1px solid #e2e8f0', background: '#fff' }}>
                <div style={{ position: 'relative' }}>
                  <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none"
                    stroke="#94a3b8" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"
                    style={{ position: 'absolute', left: '10px', top: '50%', transform: 'translateY(-50%)' }}>
                    <circle cx="11" cy="11" r="8"></circle><line x1="21" y1="21" x2="16.65" y2="16.65"></line>
                  </svg>
                  <input
                    type="text"
                    placeholder="Cari nama, NIK, alamat, desa, kabupaten..."
                    value={previewSearchTerm}
                    onChange={(e) => setPreviewSearchTerm(e.target.value)}
                    style={{
                      width: '100%', padding: '8px 12px 8px 36px', border: '1px solid #cbd5e1',
                      borderRadius: '6px', fontSize: '14px', outline: 'none', boxSizing: 'border-box'
                    }}
                  />
                  {previewSearchTerm && (
                    <button
                      onClick={() => setPreviewSearchTerm('')}
                      style={{
                        position: 'absolute', right: '8px', top: '50%', transform: 'translateY(-50%)',
                        background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', padding: '2px'
                      }}
                    >
                      <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none"
                        stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                        <line x1="18" y1="6" x2="6" y2="18"></line><line x1="6" y1="6" x2="18" y2="18"></line>
                      </svg>
                    </button>
                  )}
                </div>
              </div>

              {/* Table Content */}
              <div style={{ flex: 1, overflow: 'auto', padding: '0' }}>
                {previewTab === 'lolos' ? (
                  <table className="data-table" style={{ fontSize: '13px' }}>
                    <thead>
                      <tr>
                        <th>No</th>
                        <th>NIK</th>
                        <th>No KK</th>
                        <th>Nama</th>
                        <th>Desa/Kelurahan</th>
                        <th>Kecamatan</th>
                        <th>Kab/Kota</th>
                        <th>Jenis Kelamin</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filterPreviewRecords(previewData.lolos_records).length === 0 ? (
                        <tr><td colSpan="8" style={{ textAlign: 'center', color: '#94a3b8', padding: '40px' }}>Tidak ada data yang cocok</td></tr>
                      ) : (
                        filterPreviewRecords(previewData.lolos_records).map((r, i) => (
                          <tr key={i}>
                            <td>{i + 1}</td>
                            <td>{r.no_ktp}</td>
                            <td>{r.no_kk}</td>
                            <td>{r.nama}</td>
                            <td>{r.desa_kelurahan}</td>
                            <td>{r.kecamatan}</td>
                            <td>{r.kabupaten_kota}</td>
                            <td>{r.jenis_kelamin}</td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                ) : (
                  <table className="data-table" style={{ fontSize: '13px' }}>
                    <thead>
                      <tr>
                        <th>No</th>
                        <th>NIK</th>
                        <th>No KK</th>
                        <th>Nama</th>
                        <th>Desa/Kelurahan</th>
                        <th>Kecamatan</th>
                        <th>Kab/Kota</th>
                        <th>Alasan Tidak Lolos</th>
                        <th>Keterangan</th>
                        <th>Pengganti</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filterPreviewRecords(previewData.tidak_lolos_records).length === 0 ? (
                        <tr><td colSpan="10" style={{ textAlign: 'center', color: '#94a3b8', padding: '40px' }}>Tidak ada data yang cocok</td></tr>
                      ) : (
                        filterPreviewRecords(previewData.tidak_lolos_records).map((r, i) => (
                          <tr key={i}>
                            <td>{i + 1}</td>
                            <td>{r.no_ktp}</td>
                            <td>{r.no_kk}</td>
                            <td>{r.nama}</td>
                            <td>{r.desa_kelurahan}</td>
                            <td>{r.kecamatan}</td>
                            <td>{r.kabupaten_kota}</td>
                            <td>{r.alasan_tidak_lolos}</td>
                            <td>{r.keterangan}</td>
                            <td>
                              {r.nama_pengganti ? (
                                <span style={{ color: '#16a34a', fontSize: '12px' }}>
                                  {r.nama_pengganti} (NIK: {r.no_ktp_pengganti})
                                </span>
                              ) : <span style={{ color: '#cbd5e1' }}>—</span>}
                            </td>
                          </tr>
                        ))
                      )}
                    </tbody>
                  </table>
                )}
              </div>
            </div>
            <div className="modal-footer" style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <span style={{ fontSize: '13px', color: '#64748b' }}>
                Total: Lolos {previewData.lolos_records.length} | Tidak Lolos {previewData.tidak_lolos_records.length}
              </span>
              <div style={{ display: 'flex', gap: '10px' }}>
                <button className="btn btn-secondary" onClick={closePreview}>Tutup</button>
                <button className="btn btn-primary" onClick={() => { closePreview(); handleExport(selectedBatchIdForPreview); }}>
                  <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '6px', verticalAlign: 'middle' }}><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path><polyline points="7 10 12 15 17 10"></polyline><line x1="12" y1="15" x2="12" y2="3"></line></svg>
                  Download Excel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* SK Dirjen Approval Floating Window */}
      {skDirjenApprovalRecord && (
        <div className="modal-overlay" onClick={() => setSkDirjenApprovalRecord(null)}>
          <div className="modal-content" style={{ maxWidth: '900px', maxHeight: '90vh' }} onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3>Tinjau Ketidakcocokan Data</h3>
              <button className="modal-close" onClick={() => setSkDirjenApprovalRecord(null)}>&times;</button>
            </div>
            <div className="modal-body" style={{ display: 'flex', gap: '16px', padding: '16px', overflow: 'auto' }}>
              {/* SK Dirjen Card */}
              <div style={{ flex: 1, border: '2px solid #7c3aed', borderRadius: '12px', padding: '16px' }}>
                <h4 style={{ color: '#7c3aed', marginBottom: '12px', fontSize: '0.95rem' }}>Data SK Dirjen</h4>
                <div style={{ display: 'grid', gap: '8px', fontSize: '0.85rem' }}>
                  <div><strong>Nama:</strong> {skDirjenApprovalRecord.nama}</div>
                  <div><strong>NIK:</strong> {skDirjenApprovalRecord.no_ktp}</div>
                  <div><strong>KK:</strong> {skDirjenApprovalRecord.no_kk}</div>
                  <div><strong>Alamat:</strong> {skDirjenApprovalRecord.alamat}</div>
                  <div><strong>Desa:</strong> {skDirjenApprovalRecord.desa_kelurahan}</div>
                  <div><strong>Kecamatan:</strong> {skDirjenApprovalRecord.kecamatan}</div>
                  <div><strong>Kabupaten:</strong> {skDirjenApprovalRecord.kabupaten_kota}</div>
                </div>
              </div>

              {/* Arrow */}
              <div style={{ display: 'flex', alignItems: 'center', fontSize: '2rem', color: '#94a3b8' }}>→</div>

              {/* Verified Card */}
              <div style={{ flex: 1, border: '2px solid #16a34a', borderRadius: '12px', padding: '16px' }}>
                <h4 style={{ color: '#16a34a', marginBottom: '12px', fontSize: '0.95rem' }}>Data Terverifikasi</h4>
                <div style={{ display: 'grid', gap: '8px', fontSize: '0.85rem' }}>
                  <div><strong>Nama:</strong> {skDirjenApprovalRecord.verified_nama || '—'}</div>
                  <div><strong>NIK:</strong> {skDirjenApprovalRecord.verified_no_ktp || '—'}</div>
                  <div><strong>KK:</strong> {skDirjenApprovalRecord.verified_no_kk || '—'}</div>
                  <div><strong>Alamat:</strong> {skDirjenApprovalRecord.verified_alamat || '—'}</div>
                  <div><strong>Desa:</strong> {skDirjenApprovalRecord.verified_desa_kelurahan || '—'}</div>
                  <div><strong>Kecamatan:</strong> {skDirjenApprovalRecord.verified_kecamatan || '—'}</div>
                  <div><strong>Kabupaten:</strong> {skDirjenApprovalRecord.verified_kabupaten_kota || '—'}</div>
                  {skDirjenApprovalRecord.verified_batch_name && (
                    <div><strong>Asal:</strong> BA {skDirjenApprovalRecord.verified_batch_name} Tahap {skDirjenApprovalRecord.verified_stage_name}</div>
                  )}
                </div>
              </div>
            </div>
            <div className="modal-footer" style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px' }}>
              <button className="btn btn-secondary" onClick={() => setSkDirjenApprovalRecord(null)}>Batal</button>
              <button className="btn btn-danger" onClick={() => handleSkDirjenReject(skDirjenApprovalRecord.id)}>Tolak</button>
              <button className="btn btn-primary" onClick={() => handleSkDirjenApprove(skDirjenApprovalRecord.id)}>
                Setujui Pembaruan
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Invers Manual Pairing Modal — Pasangkan data INVERS "Belum" dengan data Terverifikasi */}
      {showPairModal && pairingInvers && (
        <div className="modal-overlay" onClick={() => setShowPairModal(false)}>
          <div className="modal-content" style={{ width: '98vw', maxHeight: '85vh' }} onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3 style={{ margin: 0, fontSize: '1rem' }}>
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px', verticalAlign: 'middle' }}><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg>
                Pasangkan dengan Data Terverifikasi
              </h3>
              <button type="button" className="modal-close" onClick={() => setShowPairModal(false)}>&times;</button>
            </div>
            <div className="modal-body" style={{ padding: '16px' }}>
              <div style={{ backgroundColor: '#f8f9fa', padding: '12px', borderRadius: '8px', marginBottom: '16px', border: '1px solid #dee2e6' }}>
                <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)', marginBottom: '4px' }}>Data INVERS yang akan dipasangkan:</div>
                <div style={{ fontWeight: 700 }}>{pairingInvers.nama}</div>
                <div style={{ fontFamily: 'monospace', fontSize: '0.9rem' }}>NIK: {pairingInvers.no_ktp}</div>
                <div style={{ fontSize: '0.85rem', color: 'var(--text-muted)' }}>{pairingInvers.kabupaten_kota}, {pairingInvers.kecamatan}, {pairingInvers.desa_kelurahan}</div>
              </div>

              <div className="form-group">
                <label className="form-label">Cari Data Terverifikasi (NIK, Nama, atau KK)</label>
                <input
                  type="text"
                  className="form-input"
                  placeholder="Ketik minimal 2 karakter untuk mencari..."
                  value={pairSearchTerm}
                  onChange={(e) => handleSearchForPairing(e.target.value)}
                  autoFocus
                />
              </div>

              {pairSearchResults.length > 0 && (
                <div style={{ maxHeight: '300px', overflowY: 'auto', overflowX: 'auto', border: '1px solid #dee2e6', borderRadius: '8px' }}>
                  <table className="data-table" style={{ width: '100%', fontSize: '0.85rem' }}>
                    <thead>
                      <tr>
                        <th>NAMA</th>
                        <th>NIK</th>
                        <th>NO KK</th>
                        <th>DESA</th>
                        <th>TAHAP</th>
                        <th></th>
                      </tr>
                    </thead>
                    <tbody>
                      {pairSearchResults.map(vr => (
                        <tr key={vr.id} style={{ cursor: 'pointer' }} onClick={() => handlePairInvers(vr.id)}>
                          <td>{vr.nama}</td>
                          <td style={{ fontFamily: 'monospace', fontSize: '0.8rem' }}>{vr.no_ktp}</td>
                          <td style={{ fontFamily: 'monospace', fontSize: '0.8rem' }}>{vr.no_kk}</td>
                          <td>{vr.desa_kelurahan}</td>
                          <td>{vr.tahap}</td>
                          <td>
                            <button className="btn btn-primary btn-sm" style={{ fontSize: '0.7rem', padding: '2px 6px' }}>
                              Pilih & Pasangkan
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
              {pairSearchTerm.length >= 2 && pairSearchResults.length === 0 && (
                <div style={{ textAlign: 'center', padding: '20px', color: 'var(--text-muted)', fontStyle: 'italic' }}>
                  Tidak ditemukan data terverifikasi yang cocok.
                </div>
              )}
            </div>
            <div className="modal-footer">
              <button className="btn btn-secondary btn-sm" onClick={() => setShowPairModal(false)}>Tutup</button>
            </div>
          </div>
        </div>
      )}

      {/* SK Dirjen Pairing Modal — Pasangkan NO_MATCH dengan data Terverifikasi global */}
      {skDirjenPairingRecord && (
        <div className="modal-overlay" onClick={() => setSkDirjenPairingRecord(null)}>
          <div className="modal-content" style={{ width: '98vw', maxHeight: '85vh' }} onClick={(e) => e.stopPropagation()}>
            <div className="modal-header">
              <h3 style={{ margin: 0, fontSize: '1rem' }}>
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" style={{ marginRight: '8px', verticalAlign: 'middle' }}><path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"></path><path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"></path></svg>
                Pasangkan dengan Data Terverifikasi
              </h3>
              <button className="modal-close" onClick={() => setSkDirjenPairingRecord(null)}>&times;</button>
            </div>
            <div className="modal-body" style={{ padding: '16px', overflow: 'auto' }}>
              <div style={{ backgroundColor: '#f1f5f9', padding: '12px', borderRadius: '6px', marginBottom: '16px', fontSize: '0.85rem' }}>
                <strong>Data SK Dirjen yang Dipilih:</strong>
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px', marginTop: '6px' }}>
                  <div><strong>Nama:</strong> {skDirjenPairingRecord.nama}</div>
                  <div><strong>NIK:</strong> {skDirjenPairingRecord.no_ktp}</div>
                  <div><strong>No. KK:</strong> {skDirjenPairingRecord.no_kk}</div>
                  <div><strong>Desa:</strong> {skDirjenPairingRecord.desa_kelurahan}</div>
                </div>
              </div>

              <div style={{ marginBottom: '12px' }}>
                <label style={{ display: 'block', fontWeight: 600, fontSize: '0.85rem', marginBottom: '4px' }}>Cari Data Terverifikasi (Semua Tahap)</label>
                <input
                  type="text"
                  className="form-input"
                  placeholder="Ketik nama, NIK, KK, atau Desa untuk mencari..."
                  value={skDirjenPairSearchTerm}
                  onChange={(e) => handleSkDirjenSearchVerified(e.target.value)}
                  autoFocus
                  style={{ fontSize: '0.85rem' }}
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
                      <th>Tahap</th>
                      <th>Aksi</th>
                    </tr>
                  </thead>
                  <tbody>
                    {skDirjenPairSearchResults.map(vr => (
                      <tr key={vr.id}>
                        <td style={{ fontWeight: 600 }}>{vr.nama}</td>
                        <td style={{ fontSize: '0.8rem' }}>{vr.no_ktp}</td>
                        <td style={{ fontSize: '0.8rem' }}>{vr.no_kk}</td>
                        <td>{vr.desa_kelurahan}</td>
                        <td style={{ fontSize: '0.8rem', color: '#6b7280' }}>{vr.tahap}</td>
                        <td>
                          <button className="btn btn-primary btn-sm" onClick={() => handleSkDirjenPair(skDirjenPairingRecord.id, vr.id)}>
                            Pilih & Pasangkan
                          </button>
                        </td>
                      </tr>
                    ))}
                    {skDirjenPairSearchTerm.length >= 2 && skDirjenPairSearchResults.length === 0 && (
                      <tr>
                        <td colSpan="6" style={{ textAlign: 'center', padding: '16px', color: 'var(--text-muted)' }}>
                          Tidak ditemukan data Terverifikasi yang cocok.
                        </td>
                      </tr>
                    )}
                    {skDirjenPairSearchTerm.length < 2 && (
                      <tr>
                        <td colSpan="6" style={{ textAlign: 'center', padding: '16px', color: 'var(--text-muted)' }}>
                          Ketik minimal 2 karakter untuk mulai mencari.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
            <div className="modal-footer" style={{ display: 'flex', justifyContent: 'flex-end', gap: '8px', padding: '12px 16px' }}>
              <button className="btn btn-secondary" onClick={() => setSkDirjenPairingRecord(null)}>Batal</button>
            </div>
          </div>
        </div>
      )}

      </main>
    </div>
  );
}

export default App;

