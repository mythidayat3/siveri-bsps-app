import sqlite3
import os
import requests
import json
import base64
import re
try:
    import psycopg2
    import psycopg2.extras
except ImportError:
    psycopg2 = None

DB_PATH = os.getenv("DATABASE_PATH") or os.path.join(os.path.dirname(os.path.abspath(__file__)), "bsps_db.sqlite")
SUPABASE_DATABASE_URL = os.getenv("SUPABASE_DATABASE_URL") or os.getenv("DATABASE_URL")
TURSO_DATABASE_URL = os.getenv("TURSO_DATABASE_URL") or os.getenv("TURSO_URL")
TURSO_AUTH_TOKEN = os.getenv("TURSO_AUTH_TOKEN") or os.getenv("TURSO_TOKEN")

def translate_sqlite_to_pg(sql):
    if not sql:
        return sql
    
    # 0. Replace AUTOINCREMENT / DDL keywords for Postgres
    sql_t = re.sub(r'\bINTEGER\s+PRIMARY\s+KEY\s+AUTOINCREMENT\b', 'SERIAL PRIMARY KEY', sql, flags=re.IGNORECASE)
    sql_t = re.sub(r'\bAUTOINCREMENT\b', '', sql_t, flags=re.IGNORECASE)
    
    # 1. Replace IFNULL with COALESCE
    sql_t = re.sub(r'\bIFNULL\b', 'COALESCE', sql_t, flags=re.IGNORECASE)
    
    # 2. Replace ? with %s when not inside quotes
    parts = []
    in_single_quote = False
    in_double_quote = False
    i = 0
    while i < len(sql_t):
        ch = sql_t[i]
        if ch == "'" and not in_double_quote:
            in_single_quote = not in_single_quote
            parts.append(ch)
        elif ch == '"' and not in_single_quote:
            in_double_quote = not in_double_quote
            parts.append(ch)
        elif ch == '?' and not in_single_quote and not in_double_quote:
            parts.append('%s')
        else:
            parts.append(ch)
        i += 1
    sql_t = "".join(parts)
    
    # 3. Handle INSERT OR IGNORE
    if re.search(r'\bINSERT\s+OR\s+IGNORE\s+INTO\b', sql_t, flags=re.IGNORECASE):
        sql_t = re.sub(r'\bINSERT\s+OR\s+IGNORE\s+INTO\b', 'INSERT INTO', sql_t, flags=re.IGNORECASE)
        if not re.search(r'\bON\s+CONFLICT\b', sql_t, flags=re.IGNORECASE):
            sql_t = sql_t.rstrip().rstrip(';') + " ON CONFLICT DO NOTHING;"

    # 4. Handle PRAGMA (ignore)
    if sql_t.strip().upper().startswith("PRAGMA"):
        return "SELECT 1;"

    # 5. Handle ALTER TABLE ... ADD COLUMN -> ADD COLUMN IF NOT EXISTS
    sql_t = re.sub(r'\bADD\s+COLUMN\s+(?!IF\s+NOT\s+EXISTS\b)', 'ADD COLUMN IF NOT EXISTS ', sql_t, flags=re.IGNORECASE)

    return sql_t

class PostgresRow(dict):
    """Row object supporting both dictionary key access (case-insensitive) and integer index access."""
    def __init__(self, cols, vals):
        super().__init__()
        self._cols = cols
        self._vals = vals
        for c, v in zip(cols, vals):
            self[c] = v
            
    def __getitem__(self, item):
        if isinstance(item, int):
            return self._vals[item]
        if isinstance(item, str) and item not in self:
            for k in self.keys():
                if k.lower() == item.lower():
                    return self[k]
        return super().__getitem__(item)
        
    def get(self, k, default=None):
        if k in self:
            return self[k]
        for key in self.keys():
            if key.lower() == str(k).lower():
                return self[key]
        return default
        
    def keys(self):
        return self._cols

class PostgresCursor:
    def __init__(self, raw_cursor):
        self._cur = raw_cursor
        self.description = None
        self.lastrowid = None
        self.rowcount = 0

    def execute(self, sql, params=None):
        pg_sql = translate_sqlite_to_pg(sql)
        is_insert = pg_sql.strip().upper().startswith("INSERT INTO") and "RETURNING" not in pg_sql.upper()
        try_sql = pg_sql.rstrip(";") + " RETURNING id" if is_insert else pg_sql
        
        try:
            if params is not None:
                if isinstance(params, (list, tuple)):
                    clean_params = tuple(None if p is None else p for p in params)
                elif isinstance(params, dict):
                    clean_params = params
                else:
                    clean_params = (params,)
                try:
                    self._cur.execute(try_sql, clean_params)
                except Exception as e:
                    if is_insert and ("column \"id\" does not exist" in str(e).lower() or "syntax error" in str(e).lower()):
                        self._cur.connection.rollback()
                        self._cur.execute(pg_sql, clean_params)
                    else:
                        raise e
            else:
                try:
                    self._cur.execute(try_sql)
                except Exception as e:
                    if is_insert and ("column \"id\" does not exist" in str(e).lower() or "syntax error" in str(e).lower()):
                        self._cur.connection.rollback()
                        self._cur.execute(pg_sql)
                    else:
                        raise e

            if is_insert and self._cur.description:
                try:
                    res = self._cur.fetchone()
                    if res:
                        self.lastrowid = res[0]
                except Exception:
                    pass
        except Exception as e:
            try:
                self._cur.connection.rollback()
            except Exception:
                pass
            raise e

        self.description = self._cur.description
        self.rowcount = self._cur.rowcount
        return self

    def executemany(self, sql, seq_of_params):
        pg_sql = translate_sqlite_to_pg(sql)
        self._cur.executemany(pg_sql, seq_of_params)
        self.description = self._cur.description
        self.rowcount = self._cur.rowcount
        return self

    def fetchone(self):
        row = self._cur.fetchone()
        if row is None:
            return None
        cols = [d[0] for d in self._cur.description]
        vals = list(row)
        return PostgresRow(cols, vals)

    def fetchall(self):
        rows = self._cur.fetchall()
        if not rows:
            return []
        cols = [d[0] for d in self._cur.description]
        return [PostgresRow(cols, list(r)) for r in rows]

    def close(self):
        try:
            self._cur.close()
        except Exception:
            pass

class PostgresConnection:
    def __init__(self, db_url):
        if psycopg2 is None:
            raise ImportError("psycopg2 is not installed. Please run: pip install psycopg2-binary")
        self._conn = psycopg2.connect(db_url)
        self._conn.autocommit = True
        self.row_factory = None

    def cursor(self):
        return PostgresCursor(self._conn.cursor())

    def execute(self, sql, params=None):
        cur = self.cursor()
        return cur.execute(sql, params)

    def commit(self):
        try:
            self._conn.commit()
        except Exception:
            pass

    def rollback(self):
        try:
            self._conn.rollback()
        except Exception:
            pass

    def close(self):
        try:
            self._conn.close()
        except Exception:
            pass

class TursoRow(dict):
    """Row object supporting both dictionary key access and integer index access."""
    def __init__(self, cols, vals):
        super().__init__()
        self._cols = cols
        self._vals = vals
        for c, v in zip(cols, vals):
            self[c] = v
            
    def __getitem__(self, item):
        if isinstance(item, int):
            return self._vals[item]
        return super().__getitem__(item)
        
    def get(self, k, default=None):
        return super().get(k, default)
        
    def keys(self):
        return self._cols

class TursoCursor:
    def __init__(self, conn):
        self.conn = conn
        self.description = None
        self.lastrowid = None
        self.rowcount = 0
        self._rows = []
        self._idx = 0

    def _convert_val(self, val):
        if val is None:
            return {"type": "null"}
        elif isinstance(val, bool):
            return {"type": "integer", "value": "1" if val else "0"}
        elif isinstance(val, int):
            return {"type": "integer", "value": str(val)}
        elif isinstance(val, float):
            return {"type": "float", "value": val}
        elif isinstance(val, bytes):
            return {"type": "blob", "base64": base64.b64encode(val).decode('utf-8')}
        else:
            return {"type": "text", "value": str(val)}

    def execute(self, sql, params=None):
        sql = sql.strip()
        if not sql:
            return self
            
        args = []
        if params is not None:
            if isinstance(params, (list, tuple)):
                args = [self._convert_val(p) for p in params]
            elif isinstance(params, dict):
                args = [self._convert_val(v) for v in params.values()]
            else:
                args = [self._convert_val(params)]

        payload = {
            "requests": [
                {
                    "type": "execute",
                    "stmt": {
                        "sql": sql,
                        "args": args
                    }
                },
                {"type": "close"}
            ]
        }

        try:
            resp = self.conn._session.post(
                self.conn._url,
                json=payload,
                timeout=30
            )
            resp.raise_for_status()
            data = resp.json()
            
            results = data.get("results", [])
            exec_res = None
            for res in results:
                if res.get("response", {}).get("type") == "execute":
                    exec_res = res.get("response", {}).get("result", {})
                    break
                if "error" in res:
                    raise Exception(f"Turso SQL Error: {res.get('error', {}).get('message')}")
            
            if exec_res:
                cols = [c["name"] for c in exec_res.get("cols", [])]
                self.description = [(c, None, None, None, None, None, None) for c in cols]
                self.lastrowid = exec_res.get("last_insert_rowid")
                self.rowcount = exec_res.get("affected_row_count", 0)
                
                rows_data = exec_res.get("rows", [])
                self._rows = []
                for r in rows_data:
                    row_vals = []
                    for col_val in r:
                        t = col_val.get("type")
                        v = col_val.get("value")
                        if t == "null" or v is None:
                            row_vals.append(None)
                        elif t == "integer":
                            row_vals.append(int(v))
                        elif t == "float":
                            row_vals.append(float(v))
                        elif t == "blob":
                            row_vals.append(base64.b64decode(col_val.get("base64", "")))
                        else:
                            row_vals.append(v)
                    self._rows.append(TursoRow(cols, row_vals))
                self._idx = 0
            else:
                self._rows = []
                self._idx = 0
                
            return self
        except Exception as e:
            raise Exception(f"Turso HTTP Execution Error: {str(e)}")

    def executemany(self, sql, seq_of_params):
        for params in seq_of_params:
            self.execute(sql, params)
        return self

    def fetchone(self):
        if self._idx < len(self._rows):
            row = self._rows[self._idx]
            self._idx += 1
            return row
        return None

    def fetchall(self):
        rows = self._rows[self._idx:]
        self._idx = len(self._rows)
        return rows

class TursoConnection:
    def __init__(self, url, token):
        self._raw_url = url
        clean_url = url.replace("libsql://", "https://").rstrip("/")
        if not clean_url.endswith("/v2/pipeline"):
            clean_url = f"{clean_url}/v2/pipeline"
        self._url = clean_url
        self._token = token
        self._session = requests.Session()
        self._session.headers.update({
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        })
        self.row_factory = None

    def cursor(self):
        return TursoCursor(self)

    def execute(self, sql, params=None):
        cur = self.cursor()
        return cur.execute(sql, params)

    def commit(self):
        pass

    def rollback(self):
        pass

    def close(self):
        pass

def get_db_connection():
    supabase_url = os.getenv("SUPABASE_DATABASE_URL") or os.getenv("DATABASE_URL") or SUPABASE_DATABASE_URL
    if supabase_url:
        return PostgresConnection(supabase_url)
    turso_url = os.getenv("TURSO_DATABASE_URL") or os.getenv("TURSO_URL") or TURSO_DATABASE_URL
    turso_token = os.getenv("TURSO_AUTH_TOKEN") or os.getenv("TURSO_TOKEN") or TURSO_AUTH_TOKEN
    if turso_url and turso_token:
        return TursoConnection(turso_url, turso_token)
        
    conn = sqlite3.connect(DB_PATH, timeout=60.0)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute("PRAGMA journal_mode=WAL;")
        conn.execute("PRAGMA busy_timeout=60000;")
        conn.execute("PRAGMA foreign_keys = ON;")
    except Exception:
        pass
    return conn

def init_db():
    if SUPABASE_DATABASE_URL or (TURSO_DATABASE_URL and TURSO_AUTH_TOKEN):
        return
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
        name TEXT,
        province_id INTEGER DEFAULT 1,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        surat_filename TEXT,
        surat_data TEXT,
        surat_uploaded_at TIMESTAMP,
        FOREIGN KEY(province_id) REFERENCES provinces(id),
        UNIQUE(name, province_id)
    )
    """)

    # Migration: add province_id to invers_stages if missing and default existing stages to 1 (SULAWESI SELATAN)
    try:
        cursor.execute("ALTER TABLE invers_stages ADD COLUMN province_id INTEGER DEFAULT 1")
    except Exception:
        pass
    try:
        cursor.execute("UPDATE invers_stages SET province_id = 1 WHERE province_id IS NULL OR province_id = 0")
    except Exception:
        pass

    # Migration: drop legacy global name unique constraint if present on Postgres and add compound unique
    try:
        cursor.execute("ALTER TABLE invers_stages DROP CONSTRAINT IF EXISTS invers_stages_name_key")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE invers_stages ADD CONSTRAINT invers_stages_name_province_unique UNIQUE (name, province_id)")
    except Exception:
        pass

    # Migration: add surat_filename, surat_data, surat_uploaded_at to invers_stages
    try:
        cursor.execute("ALTER TABLE invers_stages ADD COLUMN surat_filename TEXT")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE invers_stages ADD COLUMN surat_data TEXT")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE invers_stages ADD COLUMN surat_uploaded_at TIMESTAMP")
    except Exception:
        pass

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
    
    # Migration: add columns to existing verified_batches if missing
    try:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN is_published INTEGER DEFAULT 0")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN nomor_ba TEXT")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN tanggal_ba TEXT")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN sort_order INTEGER DEFAULT 0")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN batch_type TEXT DEFAULT 'REGULAR'")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN kabupaten TEXT")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE verified_batches ADD COLUMN metadata_json TEXT")
    except Exception:
        pass
    
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
        batch_type TEXT DEFAULT 'REGULAR',
        FOREIGN KEY(stage_id) REFERENCES invers_stages(id) ON DELETE CASCADE
    )
    """)

    try:
        cursor.execute("SELECT batch_type FROM reconciliation_overrides LIMIT 1")
    except Exception:
        cursor.execute("ALTER TABLE reconciliation_overrides ADD COLUMN batch_type TEXT DEFAULT 'REGULAR'")
    
    # 8. SK Dirjen Batches
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sk_dirjen_batches (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        stage_name TEXT NOT NULL,
        filename TEXT,
        province_id INTEGER DEFAULT 1,
        uploaded_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(province_id) REFERENCES provinces(id) ON DELETE CASCADE
    )
    """)

    try:
        cursor.execute("SELECT province_id FROM sk_dirjen_batches LIMIT 1")
    except Exception:
        cursor.execute("ALTER TABLE sk_dirjen_batches ADD COLUMN province_id INTEGER DEFAULT 1")
    
    cursor.execute("UPDATE sk_dirjen_batches SET province_id = 1 WHERE province_id IS NULL OR province_id = 0")
    conn.commit()
    
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
        full_name TEXT,
        role TEXT NOT NULL DEFAULT 'viewer',
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)

    # Ensure full_name column exists for existing table
    try:
        cursor.execute("ALTER TABLE users ADD COLUMN full_name TEXT")
    except Exception:
        pass

    # Seed admin and viewer accounts
    default_users = [
        ('yayatbalai', 'semangat45', 'Yayat Balai', 'admin'),
        ('balai_sul_3', 'balaimk5', 'Balai Sul 3', 'viewer'),
        ('balaip3kp', 'semangat45', 'Balai P3KP', 'admin'),
        ('sittihusniah', 'Pkp2025!', 'Husniah', 'admin'),
        ('nurulaulia', 'Pkp2025!', 'Nurul', 'admin'),
        ('mhidayat', 'Pkp2025!', 'Hidayat', 'admin'),
        ('irmawan', 'Pkp2025!', 'Irmawan', 'admin'),
        ('um', 'Pkp2025!', 'Khumairah', 'admin'),
        ('alfianasn', 'Pkp2025!', 'Alfian', 'admin'),
        ('novia', 'Pkp2025!', 'Novia', 'admin'),
        ('aliyah', 'Pkp2025!', 'Aliyah', 'admin'),
        ('inggrid', 'Pkp2025!', 'Inggrid', 'admin'),
        ('alfiansyah', 'Pkp2025!', 'Alfiansyah', 'admin'),
        ('alfianw', 'Pkp2025!', 'Alfian Wiratama', 'admin'),
        ('wahyuni', 'Pkp2025!', 'Yuni', 'admin'),
        ('imamakbar', 'Pkp2025!', 'Imam', 'admin'),
        ('agussalim', 'Pkp2025!', 'Agus', 'admin'),
        ('ilham', 'Pkp2025!', 'Ilham', 'admin'),
        ('anugrah', 'Pkp2025!', 'Anugrah', 'admin'),
        ('dwiayu', 'Pkp2025!', 'A. Dwi Ayu', 'admin')
    ]

    for u_name, u_pass, u_full, u_role in default_users:
        cursor.execute("""
            INSERT INTO users (username, password, full_name, role)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(username) DO UPDATE SET
                full_name = excluded.full_name,
                role = excluded.role
        """, (u_name, u_pass, u_full, u_role))

    # 13. Activity Logs Table (Audit Trail)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS activity_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        username TEXT NOT NULL,
        full_name TEXT,
        action TEXT NOT NULL,
        entity_type TEXT,
        entity_name TEXT,
        details TEXT,
        ip_address TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_activity_logs_user ON activity_logs(username);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_activity_logs_action ON activity_logs(action);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_activity_logs_created ON activity_logs(created_at DESC);")

    # 13. Village Codes Table (Master Kode Desa / Kelurahan)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS village_codes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kode_desa TEXT,
        provinsi TEXT,
        kabupaten_kota TEXT,
        kecamatan TEXT,
        desa_kelurahan TEXT,
        delineasi TEXT,
        clean_kab TEXT,
        clean_kec TEXT,
        clean_desa TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    # --- DATABASE INDEXES FOR TURSO CLOUD & LOCAL OPTIMIZATION (Up to 99% Rows Read reduction) ---
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_invers_records_rev ON invers_records(revision_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_invers_records_nik ON invers_records(no_ktp);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_invers_records_kk ON invers_records(no_kk);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_invers_records_nama ON invers_records(nama);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_invers_records_rev_geo ON invers_records(revision_id, kabupaten_kota);")
    
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_invers_revisions_stage ON invers_revisions(stage_id, is_active);")
    
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_verified_records_batch ON verified_records(batch_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_verified_records_nik ON verified_records(no_ktp);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_verified_records_kk ON verified_records(no_kk);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_verified_records_nama ON verified_records(nama);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_verified_records_batch_status ON verified_records(batch_id, status, kabupaten_kota);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_verified_records_geo ON verified_records(kabupaten_kota, kecamatan, desa_kelurahan, status);")
    
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_verified_batches_stage ON verified_batches(stage_id, batch_type);")
    
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_reconciliation_stage_ktp ON reconciliation_overrides(stage_id, original_no_ktp);")
    
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sk_records_batch ON sk_dirjen_records(batch_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sk_records_nik ON sk_dirjen_records(no_ktp);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sk_batches_stage ON sk_dirjen_batches(province_id, stage_name);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sk_matches_sk ON sk_dirjen_matches(sk_record_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_sk_matches_verified ON sk_dirjen_matches(verified_record_id);")
    
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_replacement_disq ON replacement_events(disqualified_record_id);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_invers_manual_pairs_stage ON invers_manual_pairs(stage_id);")
    
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_village_codes_clean ON village_codes(clean_kab, clean_kec, clean_desa);")
    cursor.execute("CREATE INDEX IF NOT EXISTS idx_village_codes_lookup ON village_codes(kabupaten_kota, kecamatan, desa_kelurahan);")

    conn.commit()
    conn.close()

def normalize_geo_name(name):
    if not name:
        return ""
    name = str(name).strip().upper()
    for prefix in ["KABUPATEN ", "KAB. ", "KAB ", "KOTA ", "KECAMATAN ", "KEC. ", "KEC ", "DESA ", "KELURAHAN ", "KEL. ", "KEL "]:
        if name.startswith(prefix):
            name = name[len(prefix):].strip()
    return name.replace(" ", "").replace(".", "").replace("-", "")

def lookup_village_code(conn, kabupaten, kecamatan, desa):
    if not desa:
        return ""
    cursor = conn.cursor()
    # 1. Exact match
    cursor.execute(
        "SELECT kode_desa FROM village_codes WHERE UPPER(TRIM(kabupaten_kota)) = UPPER(TRIM(?)) AND UPPER(TRIM(kecamatan)) = UPPER(TRIM(?)) AND UPPER(TRIM(desa_kelurahan)) = UPPER(TRIM(?)) LIMIT 1",
        (kabupaten or "", kecamatan or "", desa or "")
    )
    row = cursor.fetchone()
    if row and row['kode_desa']:
        return str(row['kode_desa']).strip()
    
    # 2. Normalized match
    c_kab = normalize_geo_name(kabupaten)
    c_kec = normalize_geo_name(kecamatan)
    c_desa = normalize_geo_name(desa)
    
    if c_desa and c_kec:
        cursor.execute(
            "SELECT kode_desa FROM village_codes WHERE clean_desa = ? AND (clean_kec = ? OR clean_kab = ?) LIMIT 1",
            (c_desa, c_kec, c_kab)
        )
        row = cursor.fetchone()
        if row and row['kode_desa']:
            return str(row['kode_desa']).strip()
        
    if c_desa:
        cursor.execute(
            "SELECT kode_desa FROM village_codes WHERE clean_desa = ? LIMIT 1",
            (c_desa,)
        )
        row = cursor.fetchone()
        if row and row['kode_desa']:
            return str(row['kode_desa']).strip()
        
    return ""

if __name__ == "__main__":
    init_db()
    print("Database initialized successfully at:", DB_PATH)
