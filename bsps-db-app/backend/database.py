import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bsps_db.sqlite")

def get_db_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 0. Provinces Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS provinces (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # Seed default provinces if empty
    cursor.execute("SELECT COUNT(*) as cnt FROM provinces")
    if cursor.fetchone()['cnt'] == 0:
        cursor.execute("INSERT OR IGNORE INTO provinces (id, name) VALUES (1, 'SULAWESI SELATAN')")
        cursor.execute("INSERT OR IGNORE INTO provinces (id, name) VALUES (2, 'SULAWESI TENGGARA')")
    
    # 1. INVERS Stages
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS invers_stages (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT UNIQUE,
        province_id INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(province_id) REFERENCES provinces(id)
    )
    """)

    # Migration: add province_id to invers_stages if missing and default existing stages to 1 (SULAWESI SELATAN)
    try:
        cursor.execute("SELECT province_id FROM invers_stages LIMIT 1")
    except Exception:
        cursor.execute("ALTER TABLE invers_stages ADD COLUMN province_id INTEGER DEFAULT 1")
    
    cursor.execute("UPDATE invers_stages SET province_id = 1 WHERE province_id IS NULL OR province_id = 0")
    conn.commit()
    
    # 2. INVERS Revisions
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS invers_revisions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stage_id INTEGER,
        revision_num INTEGER,
        filename TEXT,
        uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        is_active INTEGER DEFAULT 1,
        FOREIGN KEY(stage_id) REFERENCES invers_stages(id) ON DELETE CASCADE
    )
    """)
    
    # 3. INVERS Records
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS invers_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        revision_id INTEGER,
        no_urut INTEGER,
        kode_desa TEXT,
        nama TEXT,
        jenis_kelamin TEXT,
        no_ktp TEXT,
        no_kk TEXT,
        alamat TEXT,
        desa_kelurahan TEXT,
        kecamatan TEXT,
        kabupaten_kota TEXT,
        provinsi TEXT,
        deliniasi TEXT,
        catatan_katalog TEXT,
        pengusul TEXT,
        tahap TEXT,
        FOREIGN KEY(revision_id) REFERENCES invers_revisions(id) ON DELETE CASCADE
    )
    """)
    
    # 4. Verified Batches (Berita Acara batches)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS verified_batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stage_id INTEGER,
        name TEXT,
        uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        is_published INTEGER DEFAULT 0,
        FOREIGN KEY(stage_id) REFERENCES invers_stages(id) ON DELETE CASCADE
    )
    """)
    
    # Migration: add is_published to existing verified_batches if missing
    try:
        cursor.execute("SELECT is_published FROM verified_batches LIMIT 1")
    except Exception:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN is_published INTEGER DEFAULT 0")
        
    # Migration: add nomor_ba and tanggal_ba to verified_batches if missing
    try:
        cursor.execute("SELECT nomor_ba FROM verified_batches LIMIT 1")
    except Exception:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN nomor_ba TEXT")
    try:
        cursor.execute("SELECT tanggal_ba FROM verified_batches LIMIT 1")
    except Exception:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN tanggal_ba TEXT")
    try:
        cursor.execute("SELECT sort_order FROM verified_batches LIMIT 1")
    except Exception:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN sort_order INTEGER DEFAULT 0")
    
    # 5. Verified Records
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS verified_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_id INTEGER,
        no_urut INTEGER,
        kode_desa TEXT,
        nama TEXT,
        jenis_kelamin TEXT,
        no_ktp TEXT,
        no_kk TEXT,
        alamat TEXT,
        desa_kelurahan TEXT,
        kecamatan TEXT,
        kabupaten_kota TEXT,
        status TEXT, -- 'LOLOS' or 'TIDAK LOLOS'
        latitude REAL,
        longitude REAL,
        tahap TEXT,
        tanggal TEXT,
        alasan_tidak_lolos TEXT,
        keterangan TEXT,
        is_duplicate_in_previous INTEGER DEFAULT 0,
        FOREIGN KEY(batch_id) REFERENCES verified_batches(id) ON DELETE CASCADE
    )
    """)
    
    # 6. Replacement CPB Events (linked directly to verified record)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS replacement_events (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        disqualified_record_id INTEGER UNIQUE,
        nama_pengganti TEXT,
        jenis_kelamin_pengganti TEXT,
        no_ktp_pengganti TEXT,
        no_kk_pengganti TEXT,
        alamat_pengganti TEXT,
        desa_kelurahan_pengganti TEXT,
        kecamatan_pengganti TEXT,
        kabupaten_pengganti TEXT,
        FOREIGN KEY(disqualified_record_id) REFERENCES verified_records(id) ON DELETE CASCADE
    )
    """)
    
    # 7. Reconciliation Overrides
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS reconciliation_overrides (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stage_id INTEGER,
        original_no_ktp TEXT, -- original verified record NIK
        override_type TEXT, -- 'ACCEPT_VERIFIED' or 'MANUAL_EDIT'
        corrected_nama TEXT,
        corrected_no_ktp TEXT,
        corrected_no_kk TEXT,
        overridden_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(stage_id, original_no_ktp),
        FOREIGN KEY(stage_id) REFERENCES invers_stages(id) ON DELETE CASCADE
    )
    """)
    
    # 8. SK Dirjen Batches
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sk_dirjen_batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stage_name TEXT NOT NULL,
        filename TEXT,
        uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    # 9. SK Dirjen Records
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sk_dirjen_records (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        batch_id INTEGER,
        no_urut INTEGER,
        kode_desa TEXT,
        nama TEXT,
        jenis_kelamin TEXT,
        no_ktp TEXT,
        no_kk TEXT,
        alamat TEXT,
        desa_kelurahan TEXT,
        kecamatan TEXT,
        kabupaten_kota TEXT,
        keterangan TEXT,
        FOREIGN KEY(batch_id) REFERENCES sk_dirjen_batches(id) ON DELETE CASCADE
    )
    """)
    
    # 10. SK Dirjen Matches
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sk_dirjen_matches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sk_record_id INTEGER,
        verified_record_id INTEGER,
        verified_batch_id INTEGER,
        verified_stage_id INTEGER,
        match_type TEXT DEFAULT 'NO_MATCH',
        override_status TEXT DEFAULT 'PENDING',
        UNIQUE(sk_record_id),
        FOREIGN KEY(sk_record_id) REFERENCES sk_dirjen_records(id) ON DELETE CASCADE,
        FOREIGN KEY(verified_record_id) REFERENCES verified_records(id) ON DELETE SET NULL,
        FOREIGN KEY(verified_batch_id) REFERENCES verified_batches(id) ON DELETE SET NULL,
        FOREIGN KEY(verified_stage_id) REFERENCES invers_stages(id) ON DELETE SET NULL
    )
    """)
    
    # 11. Invers Manual Pairs (rekonsiliasi manual invers ↔ verified)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS invers_manual_pairs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stage_id INTEGER,
        invers_nik TEXT,
        invers_nama TEXT,
        invers_kabupaten TEXT,
        verified_record_id INTEGER,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        UNIQUE(stage_id, invers_nik),
        FOREIGN KEY(stage_id) REFERENCES invers_stages(id) ON DELETE CASCADE,
        FOREIGN KEY(verified_record_id) REFERENCES verified_records(id) ON DELETE CASCADE
    )
    """)
    
    # 12. Users Table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'viewer',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Seed default users if empty
    cursor.execute("SELECT COUNT(*) as cnt FROM users")
    if cursor.fetchone()['cnt'] == 0:
        cursor.execute("INSERT INTO users (username, password, role) VALUES ('yayatbalai', 'semangat45', 'admin')")
        cursor.execute("INSERT INTO users (username, password, role) VALUES ('balai_sul_3', 'balaimk5', 'viewer')")

    conn.commit()
    conn.close()

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully at:", DB_PATH)
