import { IconBook } from './Icons';

export default function HelpPanel() {
  return (
    <div className="help-container">
      <div className="card-section">
        <h2 style={{ fontSize: '1.2rem', fontWeight: 700, color: 'var(--primary)', marginBottom: '8px' }}>
          <IconBook />Panduan Penggunaan Aplikasi BSPS DB
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
  );
}
