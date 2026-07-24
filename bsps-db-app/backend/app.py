import io
import os
import sqlite3
import openpyxl
import docx
import zipfile
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Body
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, FileResponse
from pydantic import BaseModel
from database import get_db_connection, DB_PATH, init_db
init_db()

BASE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")

app = FastAPI(title="Backend BSPS DB")

# Aktifkan CORS untuk React frontend pada port 3000
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

def find_header_and_data(sheet):
    header_row_idx = None
    headers = []
    rows = list(sheet.iter_rows(values_only=True))
    for r_idx, row in enumerate(rows):
        cleaned_cells = [str(c).strip().replace('\n', ' ').upper() if c is not None else '' for c in row]
        if 'NAMA' in cleaned_cells and any('KTP' in c or 'NIK' in c or 'KARTU' in c for c in cleaned_cells):
            header_row_idx = r_idx
            headers = [str(c).strip().replace('\n', ' ') if c is not None else f'COL_{i}' for i, c in enumerate(row)]
            break
    if header_row_idx is not None:
        data_rows = []
        for row in rows[header_row_idx+1:]:
            if any(c is not None for c in row):
                data_rows.append(row)
        return headers, data_rows, header_row_idx
    return None, None, None

@app.get("/api/stages")
def get_stages():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT s.id, s.name, s.created_at, 
               (SELECT COUNT(*) FROM invers_records ir 
                JOIN invers_revisions irv ON ir.revision_id = irv.id 
                WHERE irv.stage_id = s.id AND irv.is_active = 1) as record_count,
               (SELECT MAX(revision_num) FROM invers_revisions WHERE stage_id = s.id) as max_revision
        FROM invers_stages s
        ORDER BY s.created_at DESC
    """)
    stages = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return stages

@app.post("/api/invers/upload")
async def upload_invers(stage_name: str = Form(...), file: UploadFile = File(...)):
    if not stage_name.strip():
        raise HTTPException(status_code=400, detail="Nama tahap tidak boleh kosong")
    
    file_bytes = await file.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Gagal membaca file Excel: {str(e)}")
    
    sheet = wb.active
    headers, data, header_row_idx = find_header_and_data(sheet)
    if headers is None:
        raise HTTPException(status_code=400, detail="Tidak dapat menemukan baris header yang berisi 'NAMA' dan 'NO KTP/NIK'")
    
    header_map = {}
    for idx, h in enumerate(headers):
        h_upper = h.upper().replace(' ', '')
        if 'NAMA' in h_upper and 'PENGGANTI' not in h_upper:
            header_map['nama'] = idx
        elif 'KTP' in h_upper or 'NIK' in h_upper:
            header_map['no_ktp'] = idx
        elif 'KK' in h_upper or 'KELUARGA' in h_upper:
            header_map['no_kk'] = idx
        elif 'DESA' in h_upper and 'KODE' in h_upper:
            header_map['kode_desa'] = idx
        elif 'DESA' in h_upper or 'KELURAHAN' in h_upper:
            header_map['desa_kelurahan'] = idx
        elif 'KECAMATAN' in h_upper:
            header_map['kecamatan'] = idx
        elif 'KABUPATEN' in h_upper or 'KOTA' in h_upper:
            header_map['kabupaten_kota'] = idx
        elif 'ALAMAT' in h_upper:
            header_map['alamat'] = idx
        elif 'PROVINSI' in h_upper:
            header_map['provinsi'] = idx
        elif 'DELINIASI' in h_upper:
            header_map['deliniasi'] = idx
        elif 'KATALOG' in h_upper:
            header_map['catatan_katalog'] = idx
        elif 'PENGUSUL' in h_upper:
            header_map['pengusul'] = idx
        elif 'TAHAP' in h_upper:
            header_map['tahap'] = idx
        elif 'URUT' in h_upper or h_upper == 'NO':
            header_map['no_urut'] = idx
            
    required_cols = ['nama', 'no_ktp', 'no_kk']
    missing_cols = [c for c in required_cols if c not in header_map]
    if missing_cols:
        raise HTTPException(status_code=400, detail=f"Kolom wajib tidak ditemukan di sheet: {', '.join(missing_cols)}")
        
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id FROM invers_stages WHERE name = ?", (stage_name,))
    stage_row = cursor.fetchone()
    if stage_row:
        stage_id = stage_row['id']
    else:
        cursor.execute("INSERT INTO invers_stages (name) VALUES (?)", (stage_name,))
        stage_id = cursor.lastrowid
        
    cursor.execute("SELECT MAX(revision_num) as max_rev FROM invers_revisions WHERE stage_id = ?", (stage_id,))
    rev_row = cursor.fetchone()
    next_rev = (rev_row['max_rev'] or 0) + 1
    
    cursor.execute("UPDATE invers_revisions SET is_active = 0 WHERE stage_id = ?", (stage_id,))
    
    cursor.execute("""
        INSERT INTO invers_revisions (stage_id, revision_num, filename, is_active)
        VALUES (?, ?, ?, 1)
    """, (stage_id, next_rev, file.filename))
    revision_id = cursor.lastrowid
    
    inserted_count = 0
    duplicate_warnings = []
    seen_in_file = set()
    
    for row in data:
        no_urut = row[header_map.get('no_urut')] if 'no_urut' in header_map and header_map['no_urut'] < len(row) else None
        kode_desa = str(row[header_map.get('kode_desa')]).strip() if 'kode_desa' in header_map and header_map['kode_desa'] < len(row) and row[header_map['kode_desa']] is not None else None
        nama = str(row[header_map.get('nama')]).strip() if row[header_map['nama']] is not None else ""
        jenis_kelamin = str(row[header_map.get('jenis_kelamin')]).strip() if 'jenis_kelamin' in header_map and header_map['jenis_kelamin'] < len(row) and row[header_map['jenis_kelamin']] is not None else None
        no_ktp = str(row[header_map.get('no_ktp')]).strip() if row[header_map['no_ktp']] is not None else ""
        no_kk = str(row[header_map.get('no_kk')]).strip() if row[header_map['no_kk']] is not None else ""
        alamat = str(row[header_map.get('alamat')]).strip() if 'alamat' in header_map and header_map['alamat'] < len(row) and row[header_map['alamat']] is not None else None
        desa_kelurahan = str(row[header_map.get('desa_kelurahan')]).strip() if 'desa_kelurahan' in header_map and header_map['desa_kelurahan'] < len(row) and row[header_map['desa_kelurahan']] is not None else None
        kecamatan = str(row[header_map.get('kecamatan')]).strip() if 'kecamatan' in header_map and header_map['kecamatan'] < len(row) and row[header_map['kecamatan']] is not None else None
        kabupaten_kota = str(row[header_map.get('kabupaten_kota')]).strip() if 'kabupaten_kota' in header_map and header_map['kabupaten_kota'] < len(row) and row[header_map['kabupaten_kota']] is not None else None
        provinsi = str(row[header_map.get('provinsi')]).strip() if 'provinsi' in header_map and header_map['provinsi'] < len(row) and row[header_map['provinsi']] is not None else None
        deliniasi = str(row[header_map.get('deliniasi')]).strip() if 'deliniasi' in header_map and header_map['deliniasi'] < len(row) and row[header_map['deliniasi']] is not None else None
        catatan_katalog = str(row[header_map.get('catatan_katalog')]).strip() if 'catatan_katalog' in header_map and header_map['catatan_katalog'] < len(row) and row[header_map['catatan_katalog']] is not None else None
        pengusul = str(row[header_map.get('pengusul')]).strip() if 'pengusul' in header_map and header_map['pengusul'] < len(row) and row[header_map['pengusul']] is not None else None
        tahap = str(row[header_map.get('tahap')]).strip() if 'tahap' in header_map and header_map['tahap'] < len(row) and row[header_map['tahap']] is not None else None

        if no_ktp.endswith('.0'): no_ktp = no_ktp[:-2]
        if no_kk.endswith('.0'): no_kk = no_kk[:-2]

        if not nama and not no_ktp and not no_kk:
            continue

        key_all = (nama, no_ktp, no_kk)
        if no_ktp in [k[1] for k in seen_in_file] or no_kk in [k[2] for k in seen_in_file]:
            duplicate_warnings.append(f"Duplikat ditemukan di file: {nama} (NIK: {no_ktp}, KK: {no_kk})")
        seen_in_file.add(key_all)

        cursor.execute("""
            INSERT INTO invers_records (
                revision_id, no_urut, kode_desa, nama, jenis_kelamin, no_ktp, no_kk,
                alamat, desa_kelurahan, kecamatan, kabupaten_kota, provinsi,
                deliniasi, catatan_katalog, pengusul, tahap
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (revision_id, no_urut, kode_desa, nama, jenis_kelamin, no_ktp, no_kk,
              alamat, desa_kelurahan, kecamatan, kabupaten_kota, provinsi,
              deliniasi, catatan_katalog, pengusul, tahap))
        inserted_count += 1
        
    conn.commit()
    conn.close()
    
    return {
        "stage_id": stage_id,
        "revision_num": next_rev,
        "filename": file.filename,
        "inserted_records": inserted_count,
        "warnings": duplicate_warnings[:10]
    }

@app.post("/api/verified/upload")
async def upload_verified(
    stage_id: int = Form(...),
    batch_name: str = Form(...),
    file: UploadFile = File(...)
):
    if not batch_name.strip():
        raise HTTPException(status_code=400, detail="Nama Berita Acara / Batch tidak boleh kosong")
        
    file_bytes = await file.read()
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Gagal membaca file Excel: {str(e)}")
        
    sheet_names = wb.sheetnames
    if len(sheet_names) < 2:
        raise HTTPException(status_code=400, detail="Excel harus memiliki minimal 2 sheet (Sheet 1: Lamp.IIA, Sheet 2: Lamp.IIIA)")
        
    ws_lolos = wb.worksheets[0]
    headers_lolos, data_lolos, _ = find_header_and_data(ws_lolos)
    if headers_lolos is None:
        raise HTTPException(status_code=400, detail="Tidak dapat menemukan baris header di sheet 1 (Lamp.IIA)")
        
    ws_tidak = wb.worksheets[1]
    headers_tidak, data_tidak, _ = find_header_and_data(ws_tidak)
    if headers_tidak is None:
        raise HTTPException(status_code=400, detail="Tidak dapat menemukan baris header di sheet 2 (Lamp.IIIA)")
        
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id FROM invers_stages WHERE id = ?", (stage_id,))
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=400, detail="Tahap INVERS yang dipilih tidak terdaftar")
        
    cursor.execute("SELECT COALESCE(MAX(sort_order), 0) + 1 FROM verified_batches WHERE stage_id = ?", (stage_id,))
    next_sort_order = cursor.fetchone()[0]
    cursor.execute("INSERT INTO verified_batches (stage_id, name, sort_order) VALUES (?, ?, ?)", (stage_id, batch_name, next_sort_order))
    batch_id = cursor.lastrowid
    
    h_lolos_map = {}
    for idx, h in enumerate(headers_lolos):
        h_upper = h.upper().replace(' ', '')
        if 'NAMA' in h_upper: h_lolos_map['nama'] = idx
        elif 'KTP' in h_upper or 'NIK' in h_upper: h_lolos_map['no_ktp'] = idx
        elif 'KK' in h_upper or 'KELUARGA' in h_upper: h_lolos_map['no_kk'] = idx
        elif 'URUT' in h_upper or h_upper == 'NO.': h_lolos_map['no_urut'] = idx
        elif 'DESA' in h_upper and 'KODE' in h_upper: h_lolos_map['kode_desa'] = idx
        elif 'DESA' in h_upper or 'KELURAHAN' in h_upper: h_lolos_map['desa_kelurahan'] = idx
        elif 'KECAMATAN' in h_upper: h_lolos_map['kecamatan'] = idx
        elif 'KABUPATEN' in h_upper or 'KOTA' in h_upper: h_lolos_map['kabupaten_kota'] = idx
        elif 'ALAMAT' in h_upper: h_lolos_map['alamat'] = idx
        elif 'LATITUDE' in h_upper: h_lolos_map['latitude'] = idx
        elif 'LONGITUDE' in h_upper: h_lolos_map['longitude'] = idx
        elif 'TAHAP' in h_upper: h_lolos_map['tahap'] = idx
        elif 'TANGGAL' in h_upper: h_lolos_map['tanggal'] = idx
        elif 'KETERANGAN' in h_upper: h_lolos_map['keterangan'] = idx
        elif 'KELAMIN' in h_upper or 'JENIS' in h_upper: h_lolos_map['jenis_kelamin'] = idx
        
    h_tidak_map = {}
    for idx, h in enumerate(headers_tidak):
        h_upper = h.upper().replace(' ', '').replace('\n', '')
        if h_upper == 'NO.': h_tidak_map['no_urut'] = idx
        elif idx == 1: h_tidak_map['nama'] = idx
        elif idx == 2: h_tidak_map['jenis_kelamin'] = idx
        elif idx == 3: h_tidak_map['no_ktp'] = idx
        elif idx == 4: h_tidak_map['no_kk'] = idx
        elif idx == 5: h_tidak_map['alamat'] = idx
        elif idx == 6: h_tidak_map['desa_kelurahan'] = idx
        elif idx == 7: h_tidak_map['kecamatan'] = idx
        elif idx == 8: h_tidak_map['kabupaten_kota'] = idx
        elif idx == 9: h_tidak_map['alasan_tidak_lolos'] = idx
        elif idx == 10: h_tidak_map['bnba'] = idx
        elif idx == 11: h_tidak_map['nama_pengganti'] = idx
        elif idx == 12: h_tidak_map['jenis_kelamin_pengganti'] = idx
        elif idx == 13: h_tidak_map['no_ktp_pengganti'] = idx
        elif idx == 14: h_tidak_map['no_kk_pengganti'] = idx
        elif idx == 15: h_tidak_map['alamat_pengganti'] = idx
        elif idx == 16: h_tidak_map['desa_kelurahan_pengganti'] = idx
        elif idx == 17: h_tidak_map['kecamatan_pengganti'] = idx
        elif idx == 18: h_tidak_map['kabupaten_pengganti'] = idx
        elif idx == 19: h_tidak_map['tahap'] = idx
        elif idx == 20: h_tidak_map['tanggal'] = idx
        elif idx == 21: h_tidak_map['keterangan'] = idx

    if 'nama' not in h_tidak_map:
        for idx, h in enumerate(headers_tidak):
            h_upper = h.upper()
            if 'NAMA' in h_upper and 'PENGGANTI' not in h_upper: h_tidak_map['nama'] = idx
            elif 'KTP' in h_upper and 'PENGGANTI' not in h_upper: h_tidak_map['no_ktp'] = idx
            elif 'KK' in h_upper and 'PENGGANTI' not in h_upper: h_tidak_map['no_kk'] = idx
            elif 'ALASAN' in h_upper: h_tidak_map['alasan_tidak_lolos'] = idx
            elif 'NAMA' in h_upper and 'PENGGANTI' in h_upper: h_tidak_map['nama_pengganti'] = idx
            elif 'KTP' in h_upper and 'PENGGANTI' in h_upper: h_tidak_map['no_ktp_pengganti'] = idx
            elif 'KK' in h_upper and 'PENGGANTI' in h_upper: h_tidak_map['no_kk_pengganti'] = idx
            
    cursor.execute("""
        SELECT no_ktp, no_kk, status FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        WHERE vb.stage_id = ?
    """, (stage_id,))
    # Map: key = NIK atau KK, value = set of statuses
    previously_verified_map = {}
    for row in cursor.fetchall():
        nik = row['no_ktp'].strip()
        kk = row['no_kk'].strip()
        status = row['status']
        if nik not in previously_verified_map:
            previously_verified_map[nik] = set()
        previously_verified_map[nik].add(status)
        if kk not in previously_verified_map:
            previously_verified_map[kk] = set()
        previously_verified_map[kk].add(status)
    
    def is_duplicate(nik, kk, status):
        nik_match = nik in previously_verified_map and status in previously_verified_map[nik]
        kk_match = kk in previously_verified_map and status in previously_verified_map[kk]
        return nik_match or kk_match
    
    def register_verified(nik, kk, status):
        if nik not in previously_verified_map:
            previously_verified_map[nik] = set()
        previously_verified_map[nik].add(status)
        if kk not in previously_verified_map:
            previously_verified_map[kk] = set()
        previously_verified_map[kk].add(status)
    
    stats = {
        "lolos_total": 0,
        "lolos_added": 0,
        "tidak_lolos_total": 0,
        "tidak_lolos_added": 0
    }
    
    # Lolos
    for row in data_lolos:
        nama = str(row[h_lolos_map.get('nama')]).strip() if h_lolos_map.get('nama') is not None and h_lolos_map['nama'] < len(row) and row[h_lolos_map['nama']] is not None else ""
        no_ktp = str(row[h_lolos_map.get('no_ktp')]).strip() if h_lolos_map.get('no_ktp') is not None and h_lolos_map['no_ktp'] < len(row) and row[h_lolos_map['no_ktp']] is not None else ""
        no_kk = str(row[h_lolos_map.get('no_kk')]).strip() if h_lolos_map.get('no_kk') is not None and h_lolos_map['no_kk'] < len(row) and row[h_lolos_map['no_kk']] is not None else ""
        
        if no_ktp.endswith('.0'): no_ktp = no_ktp[:-2]
        if no_kk.endswith('.0'): no_kk = no_kk[:-2]
        
        if not nama and not no_ktp and not no_kk:
            continue
            
        stats["lolos_total"] += 1
        if is_duplicate(no_ktp, no_kk, 'LOLOS'):
            is_dup = 1
        else:
            is_dup = 0
            register_verified(no_ktp, no_kk, 'LOLOS')
            stats["lolos_added"] += 1
            
        no_urut = row[h_lolos_map.get('no_urut')] if 'no_urut' in h_lolos_map and h_lolos_map['no_urut'] < len(row) else None
        kode_desa = str(row[h_lolos_map.get('kode_desa')]).strip() if 'kode_desa' in h_lolos_map and h_lolos_map['kode_desa'] < len(row) and row[h_lolos_map['kode_desa']] is not None else None
        jenis_kelamin = str(row[h_lolos_map.get('jenis_kelamin')]).strip() if 'jenis_kelamin' in h_lolos_map and h_lolos_map['jenis_kelamin'] < len(row) and row[h_lolos_map['jenis_kelamin']] is not None else None
        alamat = str(row[h_lolos_map.get('alamat')]).strip() if 'alamat' in h_lolos_map and h_lolos_map['alamat'] < len(row) and row[h_lolos_map['alamat']] is not None else None
        desa_kelurahan = str(row[h_lolos_map.get('desa_kelurahan')]).strip() if 'desa_kelurahan' in h_lolos_map and h_lolos_map['desa_kelurahan'] < len(row) and row[h_lolos_map['desa_kelurahan']] is not None else None
        kecamatan = str(row[h_lolos_map.get('kecamatan')]).strip() if 'kecamatan' in h_lolos_map and h_lolos_map['kecamatan'] < len(row) and row[h_lolos_map['kecamatan']] is not None else None
        kabupaten_kota = str(row[h_lolos_map.get('kabupaten_kota')]).strip() if 'kabupaten_kota' in h_lolos_map and h_lolos_map['kabupaten_kota'] < len(row) and row[h_lolos_map['kabupaten_kota']] is not None else None
        lat = row[h_lolos_map.get('latitude')] if 'latitude' in h_lolos_map and h_lolos_map['latitude'] < len(row) and row[h_lolos_map['latitude']] != '' else None
        lng = row[h_lolos_map.get('longitude')] if 'longitude' in h_lolos_map and h_lolos_map['longitude'] < len(row) and row[h_lolos_map['longitude']] != '' else None
        tahap = str(row[h_lolos_map.get('tahap')]).strip() if 'tahap' in h_lolos_map and h_lolos_map['tahap'] < len(row) and row[h_lolos_map['tahap']] is not None else None
        tanggal = str(row[h_lolos_map.get('tanggal')]).strip() if 'tanggal' in h_lolos_map and h_lolos_map['tanggal'] < len(row) and row[h_lolos_map['tanggal']] is not None else None
        keterangan = str(row[h_lolos_map.get('keterangan')]).strip() if 'keterangan' in h_lolos_map and h_lolos_map['keterangan'] < len(row) and row[h_lolos_map['keterangan']] is not None else None

        cursor.execute("""
            INSERT INTO verified_records (
                batch_id, no_urut, kode_desa, nama, jenis_kelamin, no_ktp, no_kk,
                alamat, desa_kelurahan, kecamatan, kabupaten_kota, status,
                latitude, longitude, tahap, tanggal, keterangan, is_duplicate_in_previous
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'LOLOS', ?, ?, ?, ?, ?, ?)
        """, (batch_id, no_urut, kode_desa, nama, jenis_kelamin, no_ktp, no_kk,
              alamat, desa_kelurahan, kecamatan, kabupaten_kota, lat, lng, tahap, tanggal, keterangan, is_dup))

    # Tidak Lolos
    for row in data_tidak:
        nama = str(row[h_tidak_map.get('nama')]).strip() if h_tidak_map.get('nama') is not None and h_tidak_map['nama'] < len(row) and row[h_tidak_map['nama']] is not None else ""
        no_ktp = str(row[h_tidak_map.get('no_ktp')]).strip() if h_tidak_map.get('no_ktp') is not None and h_tidak_map['no_ktp'] < len(row) and row[h_tidak_map['no_ktp']] is not None else ""
        no_kk = str(row[h_tidak_map.get('no_kk')]).strip() if h_tidak_map.get('no_kk') is not None and h_tidak_map['no_kk'] < len(row) and row[h_tidak_map['no_kk']] is not None else ""
        
        if no_ktp.endswith('.0'): no_ktp = no_ktp[:-2]
        if no_kk.endswith('.0'): no_kk = no_kk[:-2]
        
        if not nama and not no_ktp and not no_kk:
            continue
            
        stats["tidak_lolos_total"] += 1
        if is_duplicate(no_ktp, no_kk, 'TIDAK LOLOS'):
            is_dup = 1
        else:
            is_dup = 0
            register_verified(no_ktp, no_kk, 'TIDAK LOLOS')
            stats["tidak_lolos_added"] += 1
            
        no_urut = row[h_tidak_map.get('no_urut')] if 'no_urut' in h_tidak_map and h_tidak_map['no_urut'] < len(row) else None
        jenis_kelamin = str(row[h_tidak_map.get('jenis_kelamin')]).strip() if 'jenis_kelamin' in h_tidak_map and h_tidak_map['jenis_kelamin'] < len(row) and row[h_tidak_map['jenis_kelamin']] is not None else None
        alamat = str(row[h_tidak_map.get('alamat')]).strip() if 'alamat' in h_tidak_map and h_tidak_map['alamat'] < len(row) and row[h_tidak_map['alamat']] is not None else None
        desa_kelurahan = str(row[h_tidak_map.get('desa_kelurahan')]).strip() if 'desa_kelurahan' in h_tidak_map and h_tidak_map['desa_kelurahan'] < len(row) and row[h_tidak_map['desa_kelurahan']] is not None else None
        kecamatan = str(row[h_tidak_map.get('kecamatan')]).strip() if 'kecamatan' in h_tidak_map and h_tidak_map['kecamatan'] < len(row) and row[h_tidak_map['kecamatan']] is not None else None
        kabupaten_kota = str(row[h_tidak_map.get('kabupaten_kota')]).strip() if 'kabupaten_kota' in h_tidak_map and h_tidak_map['kabupaten_kota'] < len(row) and row[h_tidak_map['kabupaten_kota']] is not None else None
        alasan = str(row[h_tidak_map.get('alasan_tidak_lolos')]).strip() if 'alasan_tidak_lolos' in h_tidak_map and h_tidak_map['alasan_tidak_lolos'] < len(row) and row[h_tidak_map['alasan_tidak_lolos']] is not None else None
        tahap = str(row[h_tidak_map.get('tahap')]).strip() if 'tahap' in h_tidak_map and h_tidak_map['tahap'] < len(row) and row[h_tidak_map['tahap']] is not None else None
        tanggal = str(row[h_tidak_map.get('tanggal')]).strip() if 'tanggal' in h_tidak_map and h_tidak_map['tanggal'] < len(row) and row[h_tidak_map['tanggal']] is not None else None
        keterangan = str(row[h_tidak_map.get('keterangan')]).strip() if 'keterangan' in h_tidak_map and h_tidak_map['keterangan'] < len(row) and row[h_tidak_map['keterangan']] is not None else None

        cursor.execute("""
            INSERT INTO verified_records (
                batch_id, no_urut, kode_desa, nama, jenis_kelamin, no_ktp, no_kk,
                alamat, desa_kelurahan, kecamatan, kabupaten_kota, status,
                tahap, tanggal, alasan_tidak_lolos, keterangan, is_duplicate_in_previous
            ) VALUES (?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, 'TIDAK LOLOS', ?, ?, ?, ?, ?)
        """, (batch_id, no_urut, nama, jenis_kelamin, no_ktp, no_kk,
              alamat, desa_kelurahan, kecamatan, kabupaten_kota, tahap, tanggal, alasan, keterangan, is_dup))
        record_id = cursor.lastrowid
        
        # Pengganti
        nama_pengganti = str(row[h_tidak_map.get('nama_pengganti')]).strip() if h_tidak_map.get('nama_pengganti') is not None and h_tidak_map['nama_pengganti'] < len(row) and row[h_tidak_map['nama_pengganti']] is not None else ""
        no_ktp_pengganti = str(row[h_tidak_map.get('no_ktp_pengganti')]).strip() if h_tidak_map.get('no_ktp_pengganti') is not None and h_tidak_map['no_ktp_pengganti'] < len(row) and row[h_tidak_map['no_ktp_pengganti']] is not None else ""
        no_kk_pengganti = str(row[h_tidak_map.get('no_kk_pengganti')]).strip() if h_tidak_map.get('no_kk_pengganti') is not None and h_tidak_map['no_kk_pengganti'] < len(row) and row[h_tidak_map['no_kk_pengganti']] is not None else ""

        if no_ktp_pengganti.endswith('.0'): no_ktp_pengganti = no_ktp_pengganti[:-2]
        if no_kk_pengganti.endswith('.0'): no_kk_pengganti = no_kk_pengganti[:-2]
        
        if nama_pengganti or no_ktp_pengganti or no_kk_pengganti:
            jenis_kelamin_pengganti = str(row[h_tidak_map.get('jenis_kelamin_pengganti')]).strip() if 'jenis_kelamin_pengganti' in h_tidak_map and h_tidak_map['jenis_kelamin_pengganti'] < len(row) and row[h_tidak_map['jenis_kelamin_pengganti']] is not None else None
            alamat_pengganti = str(row[h_tidak_map.get('alamat_pengganti')]).strip() if 'alamat_pengganti' in h_tidak_map and h_tidak_map['alamat_pengganti'] < len(row) and row[h_tidak_map['alamat_pengganti']] is not None else None
            desa_kelurahan_pengganti = str(row[h_tidak_map.get('desa_kelurahan_pengganti')]).strip() if 'desa_kelurahan_pengganti' in h_tidak_map and h_tidak_map['desa_kelurahan_pengganti'] < len(row) and row[h_tidak_map['desa_kelurahan_pengganti']] is not None else None
            kecamatan_pengganti = str(row[h_tidak_map.get('kecamatan_pengganti')]).strip() if 'kecamatan_pengganti' in h_tidak_map and h_tidak_map['kecamatan_pengganti'] < len(row) and row[h_tidak_map['kecamatan_pengganti']] is not None else None
            kabupaten_pengganti = str(row[h_tidak_map.get('kabupaten_pengganti')]).strip() if 'kabupaten_pengganti' in h_tidak_map and h_tidak_map['kabupaten_pengganti'] < len(row) and row[h_tidak_map['kabupaten_pengganti']] is not None else None
            
            cursor.execute("""
                INSERT INTO replacement_events (
                    disqualified_record_id, nama_pengganti, jenis_kelamin_pengganti,
                    no_ktp_pengganti, no_kk_pengganti, alamat_pengganti,
                    desa_kelurahan_pengganti, kecamatan_pengganti, kabupaten_pengganti
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (record_id, nama_pengganti, jenis_kelamin_pengganti,
                  no_ktp_pengganti, no_kk_pengganti, alamat_pengganti,
                  desa_kelurahan_pengganti, kecamatan_pengganti, kabupaten_pengganti))
            
    conn.commit()
    conn.close()
    
    return {
        "batch_id": batch_id,
        "batch_name": batch_name,
        "stats": stats
    }

@app.get("/api/stage/{stage_id}/summary")
def get_stage_summary(stage_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, revision_num, filename FROM invers_revisions WHERE stage_id = ? AND is_active = 1", (stage_id,))
    active_rev = cursor.fetchone()
    
    cursor.execute("SELECT id, name, uploaded_at, is_published, nomor_ba, tanggal_ba, sort_order FROM verified_batches WHERE stage_id = ? ORDER BY sort_order ASC, uploaded_at ASC, id ASC", (stage_id,))
    batches = [dict(row) for row in cursor.fetchall()]
    
    lolos_total = 0
    tidak_lolos_total = 0
    replacement_total = 0
    
    for b in batches:
        cursor.execute("SELECT COUNT(*) as cnt FROM verified_records WHERE batch_id = ? AND status = 'LOLOS'", (b['id'],))
        l_cnt = cursor.fetchone()['cnt']
        lolos_total += l_cnt
        b['lolos_count'] = l_cnt
        
        cursor.execute("SELECT COUNT(*) as cnt FROM verified_records WHERE batch_id = ? AND status = 'TIDAK LOLOS'", (b['id'],))
        tl_cnt = cursor.fetchone()['cnt']
        tidak_lolos_total += tl_cnt
        b['tidak_lolos_count'] = tl_cnt
        
        cursor.execute("""
            SELECT COUNT(*) as cnt FROM replacement_events re
            JOIN verified_records vr ON re.disqualified_record_id = vr.id
            WHERE vr.batch_id = ?
        """, (b['id'],))
        r_cnt = cursor.fetchone()['cnt']
        replacement_total += r_cnt
        b['replacement_count'] = r_cnt
        
    conn.close()
    
    return {
        "active_revision": dict(active_rev) if active_rev else None,
        "batches": batches,
        "totals": {
            "lolos": lolos_total,
            "tidak_lolos": tidak_lolos_total,
            "replacements": replacement_total,
            "total_verified": lolos_total + tidak_lolos_total
        }
    }

@app.get("/api/batch/{batch_id}/kabupaten-breakdown")
def get_batch_kabupaten_breakdown(batch_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id FROM verified_batches WHERE id = ?", (batch_id,))
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=404, detail="Batch tidak ditemukan")

    cursor.execute("""
        SELECT kabupaten_kota, COUNT(*) as cnt FROM verified_records
        WHERE batch_id = ? AND status = 'LOLOS' AND kabupaten_kota IS NOT NULL
        GROUP BY kabupaten_kota
    """, (batch_id,))
    lolos_map = {row['kabupaten_kota']: row['cnt'] for row in cursor.fetchall()}

    cursor.execute("""
        SELECT kabupaten_kota, COUNT(*) as cnt FROM verified_records
        WHERE batch_id = ? AND status = 'TIDAK LOLOS' AND kabupaten_kota IS NOT NULL
        GROUP BY kabupaten_kota
    """, (batch_id,))
    tidak_lolos_map = {row['kabupaten_kota']: row['cnt'] for row in cursor.fetchall()}

    cursor.execute("""
        SELECT vr.kabupaten_kota, COUNT(*) as cnt FROM replacement_events re
        JOIN verified_records vr ON re.disqualified_record_id = vr.id
        WHERE vr.batch_id = ? AND vr.kabupaten_kota IS NOT NULL
        GROUP BY vr.kabupaten_kota
    """, (batch_id,))
    replacement_map = {row['kabupaten_kota']: row['cnt'] for row in cursor.fetchall()}

    conn.close()

    all_kabs = set(lolos_map.keys()) | set(tidak_lolos_map.keys()) | set(replacement_map.keys())
    breakdown = []
    for kab in sorted(all_kabs):
        breakdown.append({
            "kabupaten": kab,
            "lolos": lolos_map.get(kab, 0),
            "tidak_lolos": tidak_lolos_map.get(kab, 0),
            "replacement": replacement_map.get(kab, 0)
        })

    return {"breakdown": breakdown}

@app.get("/api/batch/{batch_id}/export-preview")
def get_batch_export_preview(batch_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT stage_id, name FROM verified_batches WHERE id = ?", (batch_id,))
    batch_row = cursor.fetchone()
    if not batch_row:
        conn.close()
        raise HTTPException(status_code=404, detail="Batch tidak ditemukan")
    stage_id = batch_row['stage_id']
    batch_name = batch_row['name']

    cursor.execute("SELECT name FROM invers_stages WHERE id = ?", (stage_id,))
    stage_row = cursor.fetchone()
    stage_name = stage_row['name'] if stage_row else ""

    cursor.execute("""
        SELECT vr.*, re.nama_pengganti, re.no_ktp_pengganti, re.no_kk_pengganti, re.alamat_pengganti,
               re.desa_kelurahan_pengganti as desa_pengganti, re.kecamatan_pengganti as kec_pengganti,
               re.kabupaten_pengganti as kab_pengganti, re.jenis_kelamin_pengganti,
               ro.id as override_id
        FROM verified_records vr
        LEFT JOIN replacement_events re ON re.disqualified_record_id = vr.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = ?
        WHERE vr.batch_id = ?
    """, (stage_id, batch_id))
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()

    def should_include(r):
        if r['is_duplicate_in_previous'] == 0:
            return True
        return r.get('override_id') is not None

    def clean(r):
        return {k: v for k, v in r.items() if k not in ('override_id', 'is_duplicate_in_previous', 'batch_id', 'id')}

    lolos = [clean(r) for r in records if r['status'] == 'LOLOS' and should_include(r)]
    tidak_lolos = [clean(r) for r in records if r['status'] == 'TIDAK LOLOS' and should_include(r)]

    return {
        "batch_name": batch_name,
        "stage_name": stage_name,
        "lolos_records": lolos,
        "tidak_lolos_records": tidak_lolos
    }

@app.post("/api/verified/batch/{batch_id}/delete")
def delete_verified_batch(batch_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Ambil info batch
        cursor.execute("SELECT stage_id FROM verified_batches WHERE id = ?", (batch_id,))
        batch = cursor.fetchone()
        if not batch:
            raise HTTPException(status_code=404, detail="Batch tidak ditemukan")
        
        stage_id = batch['stage_id']
        
        # Ambil semua NIK dari batch ini SEBELUM dihapus
        cursor.execute("SELECT DISTINCT no_ktp FROM verified_records WHERE batch_id = ?", (batch_id,))
        niks_in_batch = [row['no_ktp'] for row in cursor.fetchall()]
        
        # Hapus reconciliation overrides untuk NIK-NIK tersebut
        deleted_overrides = 0
        for nik in niks_in_batch:
            cursor.execute("DELETE FROM reconciliation_overrides WHERE stage_id = ? AND original_no_ktp = ?", (stage_id, nik))
            deleted_overrides += cursor.rowcount
        
        # Hapus batch (ON DELETE CASCADE akan menghapus verified_records dan replacement_events)
        cursor.execute("DELETE FROM verified_batches WHERE id = ?", (batch_id,))
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal menghapus batch: {str(e)}")
    conn.close()
    return {"message": f"Berita Acara {batch_id} berhasil dihapus. {deleted_overrides} hasil rekonsiliasi terkait ikut terhapus."}

@app.post("/api/verified/batch/{batch_id}/toggle-published")
def toggle_batch_published(batch_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id, is_published FROM verified_batches WHERE id = ?", (batch_id,))
        batch = cursor.fetchone()
        if not batch:
            raise HTTPException(status_code=404, detail="Batch tidak ditemukan")
        
        new_status = 0 if batch['is_published'] else 1
        cursor.execute("UPDATE verified_batches SET is_published = ? WHERE id = ?", (new_status, batch_id))
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal mengubah status: {str(e)}")
    conn.close()
    return {"batch_id": batch_id, "is_published": new_status}

@app.post("/api/verified/batch/{batch_id}/save-metadata")
def save_batch_metadata(batch_id: int, body: dict = Body(...)):
    nomor_ba = body.get("nomor_ba", "")
    tanggal_ba = body.get("tanggal_ba", "")
    
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id FROM verified_batches WHERE id = ?", (batch_id,))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail="Batch tidak ditemukan")
            
        cursor.execute("""
            UPDATE verified_batches 
            SET nomor_ba = ?, tanggal_ba = ? 
            WHERE id = ?
        """, (nomor_ba, tanggal_ba, batch_id))
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal menyimpan metadata: {str(e)}")
    conn.close()
    return {"status": "success", "message": "Metadata batch berhasil disimpan"}

@app.post("/api/verified/batch/{batch_id}/rename")
def rename_batch(batch_id: int, body: dict = Body(...)):
    new_name = body.get("name", "").strip()
    if not new_name:
        raise HTTPException(status_code=400, detail="Nama Berita Acara / Batch tidak boleh kosong")
        
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id FROM verified_batches WHERE id = ?", (batch_id,))
        if not cursor.fetchone():
            raise HTTPException(status_code=404, detail="Batch tidak ditemukan")
            
        cursor.execute("UPDATE verified_batches SET name = ? WHERE id = ?", (new_name, batch_id))
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal mengubah nama batch: {str(e)}")
    conn.close()
    return {"batch_id": batch_id, "name": new_name, "message": "Nama Berita Acara / Batch berhasil diperbarui"}

@app.post("/api/verified/batches/reorder")
def reorder_verified_batches(body: dict = Body(...)):
    orders = body.get("orders", [])
    if not isinstance(orders, list):
        raise HTTPException(status_code=400, detail="Format orders tidak valid")
        
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        for idx, item in enumerate(orders):
            b_id = item.get("id")
            s_order = item.get("sort_order", idx + 1)
            if b_id:
                cursor.execute("UPDATE verified_batches SET sort_order = ? WHERE id = ?", (s_order, b_id))
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal mengurutkan batch: {str(e)}")
    conn.close()
    return {"status": "success", "message": "Urutan Berita Acara / Batch berhasil diperbarui"}

@app.post("/api/verified/record/{record_id}/update-status")
def update_verified_record_status(record_id: int, body: dict = Body(...)):
    new_status = body.get("status", "").upper().strip()
    if new_status not in ["LOLOS", "TIDAK LOLOS"]:
        raise HTTPException(status_code=400, detail="Status harus 'LOLOS' atau 'TIDAK LOLOS'")
        
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id, status, batch_id, no_ktp FROM verified_records WHERE id = ?", (record_id,))
        rec = cursor.fetchone()
        if not rec:
            raise HTTPException(status_code=404, detail="Data terverifikasi tidak ditemukan")
            
        cursor.execute("UPDATE verified_records SET status = ? WHERE id = ?", (new_status, record_id))
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal mengolah pembaruan status: {str(e)}")
    conn.close()
    return {"status": "success", "record_id": record_id, "new_status": new_status, "message": f"Status berhasil diubah menjadi {new_status}"}

@app.post("/api/verified/records/bulk-update-status")
def bulk_update_verified_records_status(body: dict = Body(...)):
    record_ids = body.get("record_ids", [])
    new_status = body.get("status", "").upper().strip()
    if not record_ids or not isinstance(record_ids, list):
        raise HTTPException(status_code=400, detail="Daftar CPB tidak boleh kosong")
    if new_status not in ["LOLOS", "TIDAK LOLOS"]:
        raise HTTPException(status_code=400, detail="Status harus 'LOLOS' atau 'TIDAK LOLOS'")
        
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        placeholders = ",".join("?" for _ in record_ids)
        cursor.execute(f"UPDATE verified_records SET status = ? WHERE id IN ({placeholders})", [new_status] + record_ids)
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal mengolah pembaruan status massal: {str(e)}")
    conn.close()
    return {"status": "success", "updated_count": len(record_ids), "new_status": new_status, "message": f"{len(record_ids)} CPB berhasil diubah statusnya menjadi {new_status}"}

@app.get("/api/stage/{stage_id}/records")
def get_stage_records(stage_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT ir.* FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.stage_id = ? AND irv.is_active = 1
    """, (stage_id,))
    invers_rows = [dict(row) for row in cursor.fetchall()]
    
    cursor.execute("""
        SELECT vr.*, vb.name as batch_name, 
               re.nama_pengganti, re.no_ktp_pengganti, re.no_kk_pengganti, re.alamat_pengganti,
               re.desa_kelurahan_pengganti, re.kecamatan_pengganti, re.kabupaten_pengganti
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN replacement_events re ON re.disqualified_record_id = vr.id
        WHERE vb.stage_id = ?
    """, (stage_id,))
    verified_rows = [dict(row) for row in cursor.fetchall()]
    
    cursor.execute("SELECT * FROM reconciliation_overrides WHERE stage_id = ?", (stage_id,))
    overrides = {row['original_no_ktp']: dict(row) for row in cursor.fetchall()}
    
    conn.close()
    
    invers_map = {}
    for ir in invers_rows:
        key = (ir['nama'].strip().upper(), ir['no_ktp'].strip(), ir['no_kk'].strip())
        invers_map[key] = ir
        
    invers_by_nik = {}
    invers_by_name_desa = {}
    invers_by_kk = {}
    invers_by_name = {}
    for ir in invers_rows:
        n_clean = ir['nama'].strip().upper()
        d_clean = (ir['desa_kelurahan'] or '').strip().upper()
        nik_clean = ir['no_ktp'].strip()
        kk_clean = ir['no_kk'].strip()
        
        invers_by_nik[nik_clean] = ir
        invers_by_name_desa[(n_clean, d_clean)] = ir
        invers_by_kk[kk_clean] = ir
        if n_clean not in invers_by_name:
            invers_by_name[n_clean] = []
        invers_by_name[n_clean].append(ir)
        
    analyzed_verified = []
    mismatch_count = 0
    
    for vr in verified_rows:
        nik = vr['no_ktp'].strip()
        kk = vr['no_kk'].strip()
        nama = vr['nama'].strip()
        
        errors = []
        is_mismatch = False
        mismatch_type = ""
        vr['expected_invers'] = None
        
        if len(nik) != 16:
            errors.append(f"Panjang NIK adalah {len(nik)} karakter, seharusnya 16 digit")
            mismatch_type = "NIK_INVALID"
        if len(kk) != 16:
            errors.append(f"Panjang KK adalah {len(kk)} karakter, seharusnya 16 digit")
            mismatch_type = "KK_INVALID"
            
        if nik == kk and len(nik) > 0:
            errors.append("NIK dan KK bernilai sama")
            mismatch_type = "NIK_KK_IDENTICAL"
            
        if vr['is_duplicate_in_previous'] == 1:
            errors.append("Data sudah pernah terverifikasi di Berita Acara sebelumnya (Duplikat)")
            if not mismatch_type:
                mismatch_type = "DUPLICATE"
            
        override = overrides.get(nik)
        matched_invers = None
        
        if override:
            if override['override_type'] == 'ACCEPT_VERIFIED':
                pass
            elif override['override_type'] == 'MANUAL_EDIT':
                nama = override['corrected_nama']
                nik = override['corrected_no_ktp']
                kk = override['corrected_no_kk']
        
        key = (nama.upper(), nik, kk)
        
        if key in invers_map:
            matched_invers = invers_map[key]
        else:
            expected_invers = invers_by_nik.get(nik)
            
            # Fallback jika NIK tidak cocok, cari berdasarkan Nama + Desa, KK, atau Nama Unik
            if not expected_invers:
                n_clean = nama.upper()
                d_clean = (vr.get('desa_kelurahan') or '').strip().upper()
                kk_clean = kk
                
                if (n_clean, d_clean) in invers_by_name_desa:
                    expected_invers = invers_by_name_desa[(n_clean, d_clean)]
                elif kk_clean in invers_by_kk:
                    expected_invers = invers_by_kk[kk_clean]
                elif n_clean in invers_by_name and len(invers_by_name[n_clean]) == 1:
                    expected_invers = invers_by_name[n_clean][0]
                    
            if override:
                # Mismatch resolved via reconciliation override
                is_mismatch = False
            else:
                is_mismatch = True
                mismatch_count += 1
                
                if expected_invers:
                    mismatch_fields = []
                    if expected_invers['nama'].strip().upper() != nama.upper():
                        mismatch_fields.append("Nama")
                        if not mismatch_type: mismatch_type = "NAMA_MISMATCH"
                    if expected_invers['no_ktp'].strip() != nik:
                        mismatch_fields.append("NIK")
                        if not mismatch_type: mismatch_type = "NIK_MISMATCH"
                    if expected_invers['no_kk'].strip() != kk:
                        mismatch_fields.append("KK")
                        if not mismatch_type: mismatch_type = "KK_MISMATCH"
                    errors.append(f"Ketidakcocokan dengan data INVERS pada kolom: {', '.join(mismatch_fields)}. Seharusnya Nama: '{expected_invers['nama']}', NIK: '{expected_invers['no_ktp']}', KK: '{expected_invers['no_kk']}'")
                    vr['expected_invers'] = expected_invers
                else:
                    errors.append("Data tidak ditemukan dalam database INVERS")
                    mismatch_type = "MISSING_IN_INVERS"
                
        vr['errors'] = errors
        vr['has_error'] = len(errors) > 0 or is_mismatch
        vr['is_mismatch'] = is_mismatch
        vr['mismatch_type'] = mismatch_type
        vr['override'] = override
        
        analyzed_verified.append(vr)
        
    return {
        "invers_records": invers_rows,
        "verified_records": analyzed_verified,
        "mismatch_count": mismatch_count
    }

@app.post("/api/reconciliation/override")
def save_reconciliation_override(
    stage_id: int = Form(...),
    original_no_ktp: str = Form(...),
    override_type: str = Form(...),
    corrected_nama: str = Form(None),
    corrected_no_ktp: str = Form(None),
    corrected_no_kk: str = Form(None)
):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("""
            INSERT OR REPLACE INTO reconciliation_overrides (
                stage_id, original_no_ktp, override_type, corrected_nama, corrected_no_ktp, corrected_no_kk
            ) VALUES (?, ?, ?, ?, ?, ?)
        """, (stage_id, original_no_ktp, override_type, corrected_nama, corrected_no_ktp, corrected_no_kk))
        conn.commit()
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=400, detail=f"Gagal menyimpan perbaikan: {str(e)}")
        
    conn.close()
    return {"message": "Perbaikan rekonsiliasi berhasil disimpan"}

# --- ENDPOINT BARU UNTUK DASHBOARD PROGRESS BAR & OVERVIEW CENTER ---

@app.get("/api/stage/{stage_id}/overview-stats")
def get_overview_stats(stage_id: int):
    # Ambil semua data
    data = get_stage_records(stage_id)
    invers = data["invers_records"]
    verified = data["verified_records"]
    
    total_invers = len(invers)
    
    # Sertakan semua data terverifikasi termasuk duplikat (duplikat masuk ke reconciliation)
    active_verified = verified
    
    # Hitung kategori untuk Multi-Color Progress Bar
    green_count = 0  # Lolos & Cocok (atau memiliki override)
    yellow_count = 0 # Ada di INVERS tapi mismatch Nama/KK (belum override)
    orange_count = 0 # Tidak ada di INVERS sama sekali (belum override)
    
    matched_invers_niks = set()
    
    for v in active_verified:
        # Tentukan status pencocokan
        if v.get("override"):
            # Jika ada override (sudah diselesaikan), hitung sebagai Selesai/Lolos (Hijau)
            green_count += 1
            # Tambahkan NIK target (yang terverifikasi/terkoreksi) ke matched list agar tidak dihitung Red
            target_nik = v["override"]["corrected_no_ktp"] if v["override"]["override_type"] == 'MANUAL_EDIT' else v["no_ktp"]
            matched_invers_niks.add(target_nik)
        elif v["is_mismatch"]:
            if v.get("expected_invers"):
                yellow_count += 1
                matched_invers_niks.add(v["no_ktp"])
            else:
                orange_count += 1
        else:
            green_count += 1
            matched_invers_niks.add(v["no_ktp"])
            
    # Red: Belum verifikasi (nik invers yang tidak terjamah sama sekali)
    # Tambahkan manual pairs ke matched set
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT invers_nik FROM invers_manual_pairs WHERE stage_id = ?", (stage_id,))
    for row in cursor.fetchall():
        matched_invers_niks.add(row['invers_nik'].strip())
    conn.close()
    
    red_count = 0
    for ir in invers:
        if ir["no_ktp"] not in matched_invers_niks:
            red_count += 1
            
    return {
        "total_invers": total_invers,
        "segments": {
            "green": green_count,
            "yellow": yellow_count,
            "orange": orange_count,
            "red": red_count
        }
    }

@app.get("/api/stage/{stage_id}/overview-tables")
def get_overview_tables(stage_id: int):
    data = get_stage_records(stage_id)
    invers = data["invers_records"]
    verified = data["verified_records"]
    
    # Kunci verifikasi terpasang (gunakan record terakhir per NIK jika ada duplikat)
    verified_map = {}
    for v in verified:
        verified_map[v["no_ktp"]] = v
    
    # Tambahkan manual pairs - map invers_nik -> verified record
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT mp.invers_nik, vr.*
        FROM invers_manual_pairs mp
        JOIN verified_records vr ON vr.id = mp.verified_record_id
        WHERE mp.stage_id = ?
    """, (stage_id,))
    for row in cursor.fetchall():
        row_dict = dict(row)
        invers_nik = row_dict.pop('invers_nik', None) or row_dict.get('no_ktp', '')
        # Store with invers NIK as key so lookup by invers NIK works
        if invers_nik and invers_nik not in verified_map:
            verified_map[invers_nik] = row_dict
    conn.close()
            
    # 1. Agregasi Kabupaten
    kab_stats = {}
    # 2. Agregasi Kecamatan
    kec_stats = {}
    # 3. Agregasi Pengusul
    pengusul_stats = {}
    
    # Inisialisasi dari data INVERS
    for ir in invers:
        kab = (ir["kabupaten_kota"] or "LAINNYA").upper().strip()
        kec = (ir["kecamatan"] or "LAINNYA").upper().strip()
        peng = (ir["pengusul"] or "LAINNYA").upper().strip()
        
        # Check verified match
        v = verified_map.get(ir["no_ktp"])
        is_lolos = 1 if v and v["status"] == "LOLOS" else 0
        is_tidak_lolos = 1 if v and v["status"] == "TIDAK LOLOS" else 0
        is_belum = 1 if not v else 0
        
        # Kabupaten
        if kab not in kab_stats:
            kab_stats[kab] = {"total_cpb": 0, "lolos": 0, "tidak_lolos": 0, "belum_verifikasi": 0}
        kab_stats[kab]["total_cpb"] += 1
        kab_stats[kab]["lolos"] += is_lolos
        kab_stats[kab]["tidak_lolos"] += is_tidak_lolos
        kab_stats[kab]["belum_verifikasi"] += is_belum
        
        # Kecamatan
        if kec not in kec_stats:
            kec_stats[kec] = {"total_cpb": 0, "lolos": 0, "tidak_lolos": 0, "belum_verifikasi": 0}
        kec_stats[kec]["total_cpb"] += 1
        kec_stats[kec]["lolos"] += is_lolos
        kec_stats[kec]["tidak_lolos"] += is_tidak_lolos
        kec_stats[kec]["belum_verifikasi"] += is_belum
        
        # Pengusul
        if peng not in pengusul_stats:
            pengusul_stats[peng] = {"total_cpb": 0, "lolos": 0, "tidak_lolos": 0, "belum_verifikasi": 0}
        pengusul_stats[peng]["total_cpb"] += 1
        pengusul_stats[peng]["lolos"] += is_lolos
        pengusul_stats[peng]["tidak_lolos"] += is_tidak_lolos
        pengusul_stats[peng]["belum_verifikasi"] += is_belum
        
    return {
        "kabupaten": [{"name": k, **v} for k, v in kab_stats.items()],
        "kecamatan": [{"name": k, **v} for k, v in kec_stats.items()],
        "pengusul": [{"name": k, **v} for k, v in pengusul_stats.items()]
    }

@app.get("/api/stage/{stage_id}/pengusul-tree")
def get_pengusul_tree(stage_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Get INVERS records for hierarchy + CPB counts
    cursor.execute("""
        SELECT ir.nama, ir.no_ktp, ir.pengusul, 
               UPPER(TRIM(COALESCE(ir.kabupaten_kota, ''))) as kabupaten_kota,
               UPPER(TRIM(COALESCE(ir.kecamatan, ''))) as kecamatan,
               UPPER(TRIM(COALESCE(ir.desa_kelurahan, ''))) as desa_kelurahan
        FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.stage_id = ? AND irv.is_active = 1
    """, (stage_id,))
    invers_rows = [dict(r) for r in cursor.fetchall()]
    
    # 2. Get verified records with same filter as rekap
    cursor.execute("""
        SELECT vr.no_ktp, vr.status
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
    """, (stage_id,))
    verified_rows = [dict(r) for r in cursor.fetchall()]
    
    # 3. Get manual pairs (invers ↔ verified)
    cursor.execute("""
        SELECT mp.invers_nik, vr.status
        FROM invers_manual_pairs mp
        JOIN verified_records vr ON vr.id = mp.verified_record_id
        WHERE mp.stage_id = ?
    """, (stage_id,))
    manual_pairs = [dict(r) for r in cursor.fetchall()]
    
    conn.close()
    
    # Build NIK -> status map
    verified_map = {}
    for vr in verified_rows:
        verified_map[vr["no_ktp"].strip()] = vr["status"]
    
    # Add manual pairs to verified_map
    for mp in manual_pairs:
        nik = mp["invers_nik"].strip()
        if nik not in verified_map:
            verified_map[nik] = mp["status"]
    
    # Build the tree hierarchy from INVERS, count via NIK matching
    tree = {}
    for ir in invers_rows:
        peng = (ir["pengusul"] or "LAINNYA").strip()
        kab = ir["kabupaten_kota"] or "TIDAK DIKETAHUI"
        kec = ir["kecamatan"] or "LAINNYA"
        desa = ir["desa_kelurahan"] or "LAINNYA"
        
        if peng not in tree: tree[peng] = {}
        if kab not in tree[peng]: tree[peng][kab] = {}
        if kec not in tree[peng][kab]: tree[peng][kab][kec] = {}
        if desa not in tree[peng][kab][kec]:
            tree[peng][kab][kec][desa] = {"cpb": 0, "lolos": 0, "tidak_lolos": 0}
        
        node = tree[peng][kab][kec][desa]
        node["cpb"] += 1
        
        status = verified_map.get(ir["no_ktp"].strip())
        if status == "LOLOS":
            node["lolos"] += 1
        elif status == "TIDAK LOLOS":
            node["tidak_lolos"] += 1
    
    # Konversi ke format JSON tree
    tree_list = []
    for p_name, kabs in tree.items():
        p_node = {"name": p_name, "cpb": 0, "lolos": 0, "tidak_lolos": 0, "belum_verifikasi": 0, "children": []}
        for kb_name, kecs in kabs.items():
            kb_node = {"name": kb_name, "cpb": 0, "lolos": 0, "tidak_lolos": 0, "belum_verifikasi": 0, "children": []}
            for kc_name, desas in kecs.items():
                kc_node = {"name": kc_name, "cpb": 0, "lolos": 0, "tidak_lolos": 0, "belum_verifikasi": 0, "children": []}
                for ds_name, stats in desas.items():
                    ds_belum = stats["cpb"] - stats["lolos"] - stats["tidak_lolos"]
                    ds_node = {"name": ds_name, "cpb": stats["cpb"], "lolos": stats["lolos"], "tidak_lolos": stats["tidak_lolos"], "belum_verifikasi": ds_belum}
                    kc_node["children"].append(ds_node)
                    kc_node["cpb"] += stats["cpb"]
                    kc_node["lolos"] += stats["lolos"]
                    kc_node["tidak_lolos"] += stats["tidak_lolos"]
                kc_node["belum_verifikasi"] = kc_node["cpb"] - kc_node["lolos"] - kc_node["tidak_lolos"]
                kb_node["children"].append(kc_node)
                kb_node["cpb"] += kc_node["cpb"]
                kb_node["lolos"] += kc_node["lolos"]
                kb_node["tidak_lolos"] += kc_node["tidak_lolos"]
            kb_node["belum_verifikasi"] = kb_node["cpb"] - kb_node["lolos"] - kb_node["tidak_lolos"]
            p_node["children"].append(kb_node)
            p_node["cpb"] += kb_node["cpb"]
            p_node["lolos"] += kb_node["lolos"]
            p_node["tidak_lolos"] += kb_node["tidak_lolos"]
        p_node["belum_verifikasi"] = p_node["cpb"] - p_node["lolos"] - p_node["tidak_lolos"]
        tree_list.append(p_node)
        
    return tree_list

@app.post("/api/database/clear")
def clear_database():
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Bersihkan tabel transaksi utama
        cursor.execute("DELETE FROM replacement_events")
        cursor.execute("DELETE FROM verified_records")
        cursor.execute("DELETE FROM verified_batches")
        cursor.execute("DELETE FROM invers_records")
        cursor.execute("DELETE FROM invers_revisions")
        cursor.execute("DELETE FROM invers_stages")
        cursor.execute("DELETE FROM reconciliation_overrides")
        conn.commit()
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal membersihkan data database: {str(e)}")
    conn.close()
    return {"message": "Database berhasil disetel ulang (bersih)"}

@app.post("/api/database/cleanup-overrides")
def cleanup_orphaned_overrides():
    """Hapus reconciliation overrides yang tidak memiliki record verifikasi terkait"""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Hapus overrides yang NIK-nya tidak ada di verified_records
        cursor.execute("""
            DELETE FROM reconciliation_overrides 
            WHERE id IN (
                SELECT ro.id FROM reconciliation_overrides ro
                LEFT JOIN verified_records vr ON ro.original_no_ktp = vr.no_ktp 
                    AND vr.batch_id IN (SELECT id FROM verified_batches WHERE stage_id = ro.stage_id)
                WHERE vr.id IS NULL
            )
        """)
        deleted_count = cursor.rowcount
        conn.commit()
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal membersihkan data: {str(e)}")
    conn.close()
    return {"message": f"Berhasil membersihkan {deleted_count} rekonsiliasi yang sudah tidak terkait dengan data verifikasi"}

@app.post("/api/stage/{stage_id}/delete")
def delete_stage(stage_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Cek apakah stage ada
        cursor.execute("SELECT id, name FROM invers_stages WHERE id = ?", (stage_id,))
        stage = cursor.fetchone()
        if not stage:
            raise HTTPException(status_code=404, detail="Tahap tidak ditemukan")
        
        stage_name = stage['name']
        
        # Hapus stage (ON DELETE CASCADE akan menghapus semua data terkait):
        # - invers_revisions -> invers_records
        # - verified_batches -> verified_records -> replacement_events
        # - reconciliation_overrides
        # - invers_manual_pairs
        cursor.execute("DELETE FROM invers_stages WHERE id = ?", (stage_id,))
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal menghapus tahap: {str(e)}")
    conn.close()
    return {"message": f"Tahap '{stage_name}' berhasil dihapus beserta semua data terkait"}

@app.post("/api/stage/{stage_id}/rename")
def rename_stage(stage_id: int, body: dict = Body(...)):
    new_name = body.get("name", "").strip()
    if not new_name:
        raise HTTPException(status_code=400, detail="Nama Tahap INVERS tidak boleh kosong")
        
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        cursor.execute("SELECT id, name FROM invers_stages WHERE id = ?", (stage_id,))
        stage = cursor.fetchone()
        if not stage:
            raise HTTPException(status_code=404, detail="Tahap tidak ditemukan")
            
        cursor.execute("SELECT id FROM invers_stages WHERE LOWER(name) = LOWER(?) AND id != ?", (new_name, stage_id))
        if cursor.fetchone():
            raise HTTPException(status_code=400, detail=f"Nama Tahap '{new_name}' sudah digunakan")
            
        cursor.execute("UPDATE invers_stages SET name = ? WHERE id = ?", (new_name, stage_id))
        cursor.execute("""
            UPDATE invers_records 
            SET tahap = ? 
            WHERE revision_id IN (SELECT id FROM invers_revisions WHERE stage_id = ?)
        """, (new_name, stage_id))
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal mengubah nama tahap: {str(e)}")
    conn.close()
    return {"stage_id": stage_id, "name": new_name, "message": "Nama Tahap INVERS berhasil diperbarui"}

@app.get("/api/templates/download/{template_type}")
def download_template(template_type: str):
    if template_type == "invers":
        filepath = os.path.join(BASE_DIR, "INVERS.xlsx")
        filename = "TEMPLATE_INVERS.xlsx"
    elif template_type == "verified":
        filepath = os.path.join(BASE_DIR, "TEMPLATE_TERVERIFIKASI.xlsx")
        filename = "TEMPLATE_VERIFIKASI.xlsx"
    elif template_type == "sk_dirjen":
        filepath = os.path.join(BASE_DIR, "TEMPLATE SK DIRJEN.xlsx")
        filename = "TEMPLATE_SK_DIRJEN.xlsx"
    else:
        raise HTTPException(status_code=400, detail="Tipe template tidak dikenal")
        
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="File template tidak ditemukan di direktori lokal")
        
    return FileResponse(filepath, filename=filename)

@app.post("/api/export/docx")
async def export_docx_files(
    stage_id: int = Form(...),
    batch_id: int = Form(None),
    nomor_ba: str = Form(""),
    nomor_surat: str = Form(""),
    tanggal_ba: str = Form(""),
    lokasi_ba: str = Form(""),
    no_surat_dirjen: str = Form(""),
    tgl_surat_dirjen: str = Form(""),
    hal_surat_dirjen: str = Form("")
):
    base_dir = BASE_DIR
    path_ba_template = os.path.join(base_dir, "FORMAT BERITA ACARA.docx")
    path_sp_template = os.path.join(base_dir, "SURAT PENYAMPAIAN BA.docx")
    
    if not os.path.exists(path_ba_template) or not os.path.exists(path_sp_template):
        raise HTTPException(status_code=404, detail="File template Word (.docx) tidak ditemukan")
        
    # Ambil Data
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM invers_stages WHERE id = ?", (stage_id,))
    stage_name = cursor.fetchone()['name']
    
    batch_name = None
    if batch_id:
        cursor.execute("SELECT name FROM verified_batches WHERE id = ?", (batch_id,))
        batch_row = cursor.fetchone()
        if batch_row:
            batch_name = batch_row['name']
        cursor.execute("""
            UPDATE verified_batches 
            SET nomor_ba = ?, tanggal_ba = ? 
            WHERE id = ?
        """, (nomor_ba, tanggal_ba, batch_id))
        conn.commit()
            
    # Total data INVERS di tahap aktif
    cursor.execute("""
        SELECT COUNT(*) as cnt FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.stage_id = ? AND irv.is_active = 1
    """, (stage_id,))
    total_invers = cursor.fetchone()['cnt']
    
    # Total verifikasi aktif (Lolos + Tidak Lolos di Berita Acara yang aktif)
    if batch_id:
        cursor.execute("""
            SELECT COUNT(*) as cnt FROM verified_records vr
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = ?
            WHERE vr.batch_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
        """, (stage_id, batch_id))
    else:
        cursor.execute("""
            SELECT COUNT(*) as cnt FROM verified_records vr
            JOIN verified_batches vb ON vr.batch_id = vb.id
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
            WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
        """, (stage_id,))
    total_verifikasi_aktif = cursor.fetchone()['cnt']
    
    # Total verifikasi seluruhnya di tahap aktif (untuk menghitung sisa belum verifikasi)
    cursor.execute("""
        SELECT COUNT(*) as cnt FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
    """, (stage_id,))
    total_verif_tahap_seluruhnya = cursor.fetchone()['cnt']
    sisa_belum_verif = max(0, total_invers - total_verif_tahap_seluruhnya)
    
    # Daftar rincian usulan per kabupaten untuk Berita Acara yang aktif
    if batch_id:
        cursor.execute("""
            SELECT vr.kabupaten_kota,
                   SUM(CASE WHEN vr.status = 'LOLOS' THEN 1 ELSE 0 END) as lolos,
                   SUM(CASE WHEN vr.status = 'TIDAK LOLOS' THEN 1 ELSE 0 END) as tidak_lolos,
                   COUNT(*) as total
            FROM verified_records vr
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = ?
            WHERE vr.batch_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
            GROUP BY UPPER(TRIM(vr.kabupaten_kota))
            ORDER BY vr.kabupaten_kota ASC
        """, (stage_id, batch_id))
    else:
        cursor.execute("""
            SELECT vr.kabupaten_kota,
                   SUM(CASE WHEN vr.status = 'LOLOS' THEN 1 ELSE 0 END) as lolos,
                   SUM(CASE WHEN vr.status = 'TIDAK LOLOS' THEN 1 ELSE 0 END) as tidak_lolos,
                   COUNT(*) as total
            FROM verified_records vr
            JOIN verified_batches vb ON vr.batch_id = vb.id
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
            WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
            GROUP BY UPPER(TRIM(vr.kabupaten_kota))
            ORDER BY vr.kabupaten_kota ASC
        """, (stage_id,))
    kab_details = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    # Perihal surat penyampaian
    perihal_surat_penyampaian = f"Penyampaian Hasil Verifikasi Calon Penerima Bantuan (CPB) Kegiatan Bantuan Stimulan Perumahan Swadaya (BSPS) {stage_name} TA 2026 Provinsi Sulawesi Selatan"
    
    replacements = {
        "[NOMOR BA]": nomor_ba,
        "[Nomor Berita Acara]": nomor_ba,
        "[Nomor Surat]": nomor_surat,
        "[NAMA TAHAP]": stage_name.upper(),
        "[Nama Tahap]": stage_name,
        "[NAMA PROVINSI]": "SULAWESI SELATAN",
        "[Nama Provinsi]": "Sulawesi Selatan",
        "[Tanggal Eksport BA]": tanggal_ba,
        "[Lokasi BA]": lokasi_ba,
        "[Nomor Surat Dirjen]": no_surat_dirjen,
        "[Tanggal Surat Dirjen]": tgl_surat_dirjen,
        "[Hal Surat Dirjen]": hal_surat_dirjen,
        "[Perihal Surat]": perihal_surat_penyampaian,
        "[Total Hasil Verifikasi di Tahap Aktif]": str(total_verifikasi_aktif),
        "[Total Data Invers Di Tahap aktif]": str(total_invers),
        "[Tanggal Batas Proses Verifikasi]": tanggal_ba
    }
    
    def process_document(template_path, reps):
        doc = docx.Document(template_path)
        for p in doc.paragraphs:
            for k, v in reps.items():
                if k in p.text:
                    for run in p.runs:
                        if k in run.text:
                            run.text = run.text.replace(k, v)
                    if k in p.text:
                        p.text = p.text.replace(k, v)
                        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for k, v in reps.items():
                            for run in p.runs:
                                if k in run.text:
                                    run.text = run.text.replace(k, v)
                            if k in p.text:
                                p.text = p.text.replace(k, v)
        return doc
        
    doc_ba = process_document(path_ba_template, replacements)
    doc_sp = process_document(path_sp_template, replacements)
    
    # FORMATTING OVERRIDES - 1. FORMAT BERITA ACARA
    for p in doc_ba.paragraphs:
        # baris Nomor BA samakan font dengan judul (Bookman Old Style, 11 pt, Bold)
        if p.text.strip().upper().startswith("NOMOR"):
            for run in p.runs:
                run.font.name = "Bookman Old Style"
                run.font.size = docx.shared.Pt(11)
                run.font.bold = True
        # Paragraf "Sehubungan dengan... dst." menggunakan font size 11
        if p.text.startswith("Sehubungan dengan"):
            for run in p.runs:
                run.font.size = docx.shared.Pt(11)
                
    # FORMATTING OVERRIDES - 2. SURAT PENYAMPAIAN BA
    for p in doc_sp.paragraphs:
        # baris yang menuliskan nomor dan tanggal surat menjadi Arial, font size 11, tambah [Lokasi BA] didepan tanggal
        if p.text.startswith("Nomor") and "[Tanggal Eksport BA]" in p.text:
            p.text = f"Nomor\t\t: {nomor_surat}\t\t\t\t\t{lokasi_ba}, {tanggal_ba}"
            for run in p.runs:
                run.font.name = "Arial"
                run.font.size = docx.shared.Pt(11)
                
    # Isian tabel di Surat Penyampaian BA
    if len(doc_sp.tables) > 0:
        table = doc_sp.tables[0]
        if len(table.rows) > 1:
            # Kolom Jumlah Usulan (Kolom 3 / indeks 3)
            cell_js = table.cell(1, 3)
            cell_js.text = str(total_verifikasi_aktif)
            for p in cell_js.paragraphs:
                p.alignment = 1  # Center
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = docx.shared.Pt(10)
            
            # Kolom Rincian Usulan per Kabupaten/Kota (Kolom 4 / indeks 4)
            cell_rinci = table.cell(1, 4)
            # Bersihkan isi lama
            p_first = cell_rinci.paragraphs[0]
            p_first.text = ""
            while len(cell_rinci.paragraphs) > 1:
                p_to_del = cell_rinci.paragraphs[-1]
                p_to_del._element.getparent().remove(p_to_del._element)
                
            is_first = True
            for kab_row in kab_details:
                kab_name = (kab_row['kabupaten_kota'] or "LAINNYA").upper().strip()
                total = kab_row['total']
                lolos = kab_row['lolos']
                tl = kab_row['tidak_lolos']
                
                if is_first:
                    p_kab = p_first
                    is_first = False
                else:
                    p_kab = cell_rinci.add_paragraph()
                    
                p_kab.paragraph_format.left_indent = docx.shared.Pt(12)
                run_kab = p_kab.add_run(f"•  {kab_name} (Total Verifikasi: {total} unit)")
                run_kab.bold = True
                run_kab.font.name = "Arial"
                run_kab.font.size = docx.shared.Pt(10)
                
                p_lolos = cell_rinci.add_paragraph()
                p_lolos.paragraph_format.left_indent = docx.shared.Pt(24)
                run_lolos = p_lolos.add_run(f"- Lolos: {lolos} unit")
                run_lolos.font.name = "Arial"
                run_lolos.font.size = docx.shared.Pt(9.5)
                
                p_tl = cell_rinci.add_paragraph()
                p_tl.paragraph_format.left_indent = docx.shared.Pt(24)
                run_tl = p_tl.add_run(f"- Tidak Lolos: {tl} unit")
                run_tl.font.name = "Arial"
                run_tl.font.size = docx.shared.Pt(9.5)
                
            # Kolom Keterangan (Kolom 5 / indeks 5)
            cell_ket = table.cell(1, 5)
            cell_ket.text = f"Sebanyak {total_verifikasi_aktif} unit dari total {total_invers} unit telah diverifikasi, sehingga yang masih proses verifikasi sebanyak {sisa_belum_verif} unit"
            for p in cell_ket.paragraphs:
                for run in p.runs:
                    run.font.name = "Arial"
                    run.font.size = docx.shared.Pt(10)
                    
    # Simpan ke byte stream
    stream_ba = io.BytesIO()
    doc_ba.save(stream_ba)
    stream_ba.seek(0)
    
    stream_sp = io.BytesIO()
    doc_sp.save(stream_sp)
    stream_sp.seek(0)
    
    name_suffix = (batch_name or stage_name).replace(' ', '_')
    stage_upper = stage_name.upper()
    
    # Buat file ZIP berisi kedua dokumen
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{stage_upper}_BAHV_{batch_name or stage_name}.docx", stream_ba.getvalue())
        z.writestr(f"{stage_upper}_PENYAMPAIAN_BAHV_{batch_name or stage_name}.docx", stream_sp.getvalue())
    zip_buffer.seek(0)
    
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=DRAFT_SURAT_BA_{name_suffix}.zip"}
    )

@app.post("/api/export/pdf")
async def export_pdf_files(
    stage_id: int = Form(...),
    batch_id: int = Form(None),
    nomor_ba: str = Form(""),
    nomor_surat: str = Form(""),
    tanggal_ba: str = Form(""),
    lokasi_ba: str = Form(""),
    no_surat_dirjen: str = Form(""),
    tgl_surat_dirjen: str = Form(""),
    hal_surat_dirjen: str = Form("")
):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("SELECT name FROM invers_stages WHERE id = ?", (stage_id,))
    stage_name = cursor.fetchone()['name']
    
    batch_name = None
    if batch_id:
        cursor.execute("SELECT name FROM verified_batches WHERE id = ?", (batch_id,))
        batch_row = cursor.fetchone()
        if batch_row:
            batch_name = batch_row['name']
        cursor.execute("""
            UPDATE verified_batches 
            SET nomor_ba = ?, tanggal_ba = ? 
            WHERE id = ?
        """, (nomor_ba, tanggal_ba, batch_id))
        conn.commit()
    
    cursor.execute("""
        SELECT COUNT(*) as cnt FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.stage_id = ? AND irv.is_active = 1
    """, (stage_id,))
    total_invers = cursor.fetchone()['cnt']
    
    if batch_id:
        cursor.execute("""
            SELECT COUNT(*) as cnt FROM verified_records vr
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = ?
            WHERE vr.batch_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
        """, (stage_id, batch_id))
    else:
        cursor.execute("""
            SELECT COUNT(*) as cnt FROM verified_records vr
            JOIN verified_batches vb ON vr.batch_id = vb.id
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
            WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
        """, (stage_id,))
    total_verifikasi_aktif = cursor.fetchone()['cnt']
    
    cursor.execute("""
        SELECT COUNT(*) as cnt FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
    """, (stage_id,))
    total_verif_tahap_seluruhnya = cursor.fetchone()['cnt']
    sisa_belum_verif = max(0, total_invers - total_verif_tahap_seluruhnya)
    
    if batch_id:
        cursor.execute("""
            SELECT vr.kabupaten_kota,
                   SUM(CASE WHEN vr.status = 'LOLOS' THEN 1 ELSE 0 END) as lolos,
                   SUM(CASE WHEN vr.status = 'TIDAK LOLOS' THEN 1 ELSE 0 END) as tidak_lolos,
                   COUNT(*) as total
            FROM verified_records vr
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = ?
            WHERE vr.batch_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
            GROUP BY UPPER(TRIM(vr.kabupaten_kota))
            ORDER BY vr.kabupaten_kota ASC
        """, (stage_id, batch_id))
    else:
        cursor.execute("""
            SELECT vr.kabupaten_kota,
                   SUM(CASE WHEN vr.status = 'LOLOS' THEN 1 ELSE 0 END) as lolos,
                   SUM(CASE WHEN vr.status = 'TIDAK LOLOS' THEN 1 ELSE 0 END) as tidak_lolos,
                   COUNT(*) as total
            FROM verified_records vr
            JOIN verified_batches vb ON vr.batch_id = vb.id
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
            WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
            GROUP BY UPPER(TRIM(vr.kabupaten_kota))
            ORDER BY vr.kabupaten_kota ASC
        """, (stage_id,))
    kab_details = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    perihal_surat = f"Penyampaian Hasil Verifikasi Calon Penerima Bantuan (CPB) Kegiatan Bantuan Stimulan Perumahan Swadaya (BSPS) {stage_name} TA 2026 Provinsi Sulawesi Selatan"
    
    styles = getSampleStyleSheet()
    
    style_title = ParagraphStyle('TitleBA', parent=styles['Title'], fontName='Times-Bold', fontSize=14, alignment=TA_CENTER, spaceAfter=6)
    style_nomor = ParagraphStyle('NomorBA', parent=styles['Normal'], fontName='Times-Bold', fontSize=11, alignment=TA_CENTER, spaceAfter=4)
    style_subtitle = ParagraphStyle('SubtitleBA', parent=styles['Normal'], fontName='Times-Bold', fontSize=11, alignment=TA_CENTER, spaceAfter=2)
    style_normal = ParagraphStyle('NormalBA', parent=styles['Normal'], fontName='Times-Roman', fontSize=11, alignment=TA_JUSTIFY, spaceAfter=6, leading=14)
    style_body = ParagraphStyle('BodyBA', parent=styles['Normal'], fontName='Times-Roman', fontSize=11, alignment=TA_JUSTIFY, spaceAfter=8, leading=15, firstLineIndent=2*cm)
    style_footer = ParagraphStyle('FooterBA', parent=styles['Normal'], fontName='Times-Roman', fontSize=11, alignment=TA_LEFT, spaceAfter=4)
    style_small = ParagraphStyle('SmallBA', parent=styles['Normal'], fontName='Times-Roman', fontSize=9.5, alignment=TA_LEFT, spaceAfter=2, leading=12)
    
    elements_ba = []
    elements_ba.append(Spacer(1, 1*cm))
    elements_ba.append(Paragraph("BERITA ACARA", style_title))
    elements_ba.append(Paragraph(f"NOMOR : {nomor_ba}", style_nomor))
    elements_ba.append(Spacer(1, 0.3*cm))
    elements_ba.append(Paragraph(f"HASIL VERIFIKASI CALON PENERIMA BANTUAN (CPB) {stage_name.upper()}", style_subtitle))
    elements_ba.append(Paragraph("KEGIATAN BANTUAN STIMULAN PERUMAHAN SWADAYA (BSPS) TA.2026", style_subtitle))
    elements_ba.append(Paragraph("PROVINSI SULAWESI SELATAN", style_subtitle))
    elements_ba.append(Spacer(1, 0.8*cm))
    elements_ba.append(Paragraph(
        f"Sehubungan dengan pelaksanaan verifikasi calon penerima bantuan yang telah dilakukan di "
        f"Provinsi Sulawesi Selatan dengan {stage_name}, maka bersama ini disampaikan Berita Acara "
        f"Hasil Verifikasi Calon Penerima Bantuan (CPB) Kegiatan Bantuan Stimulan Perumahan Swadaya "
        f"(BSPS) Tahun Anggaran 2026.", style_body))
    elements_ba.append(Paragraph(
        f"Data usulan calon penerima bantuan di Provinsi Sulawesi Selatan dengan Alokasi "
        f"{total_invers} unit, telah dilakukan verifikasi lapangan terhadap calon penerima bantuan "
        f"dengan hasil sebagaimana berikut:", style_body))
    elements_ba.append(Spacer(1, 0.3*cm))
    elements_ba.append(Paragraph("Lampiran I  Hasil Verifikasi Calon Penerima Bantuan Kegiatan BSPS Tahun Anggaran 2026 di Provinsi Sulawesi Selatan", style_normal))
    elements_ba.append(Paragraph("Lampiran II  Daftar Calon Penerima Bantuan Kegiatan BSPS Tahun Anggaran 2026 Provinsi Sulawesi Selatan", style_normal))
    elements_ba.append(Paragraph("Lampiran III  Daftar Calon Penerima Bantuan Pengganti Kegiatan BSPS Tahun Anggaran 2026 Provinsi Sulawesi Selatan", style_normal))
    elements_ba.append(Spacer(1, 0.3*cm))
    elements_ba.append(Paragraph(
        "Penggantian data calon penerima bantuan bersumber dari data pengusul K/L, aplikasi e-RTLH, "
        "DTKS, sistem informasi lainnya, serta hasil verifikasi dan validasi lapangan.", style_body))
    elements_ba.append(Paragraph(
        "Demikian Berita Acara ini dibuat dengan sebenarnya dan dapat dipergunakan sebagaimana mestinya.", style_body))
    elements_ba.append(Spacer(1, 1.5*cm))
    elements_ba.append(Paragraph(f"{lokasi_ba}, {tanggal_ba}", style_footer))
    elements_ba.append(Spacer(1, 1.5*cm))
    
    elements_ba.append(Paragraph("Membuat Berita Acara,", style_footer))
    elements_ba.append(Spacer(1, 1.5*cm))
    elements_ba.append(Paragraph("_________________________", style_footer))
    elements_ba.append(Paragraph("Operator Verifikasi", style_small))
    
    stream_ba = io.BytesIO()
    doc_ba = SimpleDocTemplate(stream_ba, pagesize=A4, leftMargin=2.5*cm, rightMargin=2.5*cm, topMargin=2*cm, bottomMargin=2*cm)
    doc_ba.build(elements_ba)
    stream_ba.seek(0)
    
    elements_sp = []
    elements_sp.append(Spacer(1, 1*cm))
    elements_sp.append(Paragraph(f"Nomor\t\t: {nomor_surat}\t\t\t\t\t{lokasi_ba}, {tanggal_ba}", style_normal))
    elements_sp.append(Paragraph("Sifat\t\t: Penting", style_normal))
    elements_sp.append(Paragraph("Lampiran\t: 1 (Satu) Berkas", style_normal))
    elements_sp.append(Paragraph(f"Perihal\t\t: {perihal_surat}.", style_normal))
    elements_sp.append(Spacer(1, 0.8*cm))
    elements_sp.append(Paragraph("Yth,", style_normal))
    elements_sp.append(Paragraph("Direktur Jenderal Kawasan Permukiman", style_normal))
    elements_sp.append(Paragraph("di-", style_normal))
    elements_sp.append(Paragraph("          Jakarta", style_normal))
    elements_sp.append(Spacer(1, 0.5*cm))
    elements_sp.append(Paragraph(
        f"Menindaklanjuti Surat Direktur Jenderal Kawasan Permukiman Nomor {no_surat_dirjen} "
        f"Tanggal {tgl_surat_dirjen} {hal_surat_dirjen}.", style_body))
    elements_sp.append(Paragraph(
        f"Bersama surat ini kami sampaikan Hasil Verifikasi Calon Penerima Bantuan (CPB) "
        f"Kegiatan BSPS {stage_name} Tahun Anggaran 2026 Provinsi Sulawesi Selatan.", style_body))
    elements_sp.append(Spacer(1, 0.5*cm))
    elements_sp.append(Paragraph("Rincian usulan per Kabupaten/Kota:", style_normal))
    elements_sp.append(Spacer(1, 0.2*cm))
    
    table_data = [['No.', 'Kabupaten/Kota', 'Lolos', 'Tidak Lolos', 'Total']]
    for idx, kab in enumerate(kab_details, 1):
        kab_name = (kab['kabupaten_kota'] or "LAINNYA").upper().strip()
        table_data.append([str(idx), kab_name, str(kab['lolos']), str(kab['tidak_lolos']), str(kab['total'])])
    table_data.append(['', 'TOTAL', '', '', str(total_verifikasi_aktif)])
    
    t = Table(table_data, colWidths=[1*cm, 6*cm, 2.5*cm, 2.5*cm, 2.5*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.Color(0.9, 0.9, 0.9)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Times-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('FONTNAME', (0, 1), (-1, -1), 'Times-Roman'),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('ALIGN', (2, 1), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('FONTNAME', (0, -1), (-1, -1), 'Times-Bold'),
        ('BACKGROUND', (0, -1), (-1, -1), colors.Color(0.95, 0.95, 0.95)),
    ]))
    elements_sp.append(t)
    elements_sp.append(Spacer(1, 0.5*cm))
    elements_sp.append(Paragraph(
        f"Sebanyak {total_verifikasi_aktif} unit dari total {total_invers} unit telah diverifikasi, "
        f"sehingga yang masih dalam proses verifikasi sebanyak {sisa_belum_verif} unit.", style_body))
    elements_sp.append(Paragraph(
        "Demikian surat ini disampaikan, atas perhatian dan perkenaan Bapak, kami ucapkan terima kasih.", style_body))
    elements_sp.append(Spacer(1, 1.5*cm))
    elements_sp.append(Paragraph("Kepala Balai,", style_footer))
    elements_sp.append(Spacer(1, 2*cm))
    elements_sp.append(Paragraph("_________________________", style_footer))
    elements_sp.append(Paragraph("Bakhtiar", style_small))
    elements_sp.append(Paragraph("NIP. 19711009 200212 1 003", style_small))
    elements_sp.append(Spacer(1, 0.8*cm))
    elements_sp.append(Paragraph("Tembusan:", style_small))
    elements_sp.append(Paragraph("1. Direktur Jenderal Tata Kelola dan Pengendalian Resiko;", style_small))
    elements_sp.append(Paragraph("2. Kepala Pusat Data dan Informasi, Sekretariat Jenderal Kementerian PKP.", style_small))
    
    stream_sp = io.BytesIO()
    doc_sp = SimpleDocTemplate(stream_sp, pagesize=A4, leftMargin=2.5*cm, rightMargin=2.5*cm, topMargin=2*cm, bottomMargin=2*cm)
    doc_sp.build(elements_sp)
    stream_sp.seek(0)
    
    name_suffix = (batch_name or stage_name).replace(' ', '_')
    stage_upper = stage_name.upper()
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr(f"{stage_upper}_BAHV_{batch_name or stage_name}.pdf", stream_ba.getvalue())
        z.writestr(f"{stage_upper}_PENYAMPAIAN_BAHV_{batch_name or stage_name}.pdf", stream_sp.getvalue())
    zip_buffer.seek(0)
    
    return StreamingResponse(
        zip_buffer,
        media_type="application/zip",
        headers={"Content-Disposition": f"attachment; filename=DRAFT_SURAT_BA_{name_suffix}.zip"}
    )

@app.post("/api/verified/record/{record_id}/delete")
def delete_verified_record(record_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        # Check if record exists
        cursor.execute("SELECT id, no_ktp FROM verified_records WHERE id = ?", (record_id,))
        row = cursor.fetchone()
        if not row:
            raise HTTPException(status_code=404, detail="Data verifikasi tidak ditemukan")
        
        nik = row['no_ktp']
        # Delete replacement events if any
        cursor.execute("DELETE FROM replacement_events WHERE disqualified_record_id = ?", (record_id,))
        # Delete verified record
        cursor.execute("DELETE FROM verified_records WHERE id = ?", (record_id,))
        
        # Also clean up reconciliation override if any associated with this NIK
        cursor.execute("DELETE FROM reconciliation_overrides WHERE original_no_ktp = ?", (nik,))
        
        conn.commit()
    except Exception as e:
        conn.rollback()
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal menghapus data: {str(e)}")
    conn.close()
    return {"message": "Data verifikasi berhasil dihapus"}

# --- BULK OPERATIONS ---

@app.post("/api/verified/records/bulk-delete")
async def bulk_delete_verified_records(record_ids: str = Form(...)):
    """Bulk delete verified records by comma-separated IDs."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        ids = [int(rid.strip()) for rid in record_ids.split(',') if rid.strip()]
        if not ids:
            raise HTTPException(status_code=400, detail="Tidak ada ID yang diberikan")
        
        placeholders = ','.join(['?' for _ in ids])
        
        # Get NIKs before deleting
        cursor.execute(f"SELECT DISTINCT no_ktp FROM verified_records WHERE id IN ({placeholders})", ids)
        niks = [row['no_ktp'] for row in cursor.fetchall()]
        
        # Delete replacement events
        cursor.execute(f"DELETE FROM replacement_events WHERE disqualified_record_id IN ({placeholders})", ids)
        
        # Delete verified records
        cursor.execute(f"DELETE FROM verified_records WHERE id IN ({placeholders})", ids)
        deleted_count = cursor.rowcount
        
        # Clean up reconciliation overrides
        for nik in niks:
            cursor.execute("DELETE FROM reconciliation_overrides WHERE original_no_ktp = ?", (nik,))
        
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal menghapus data secara bulk: {str(e)}")
    conn.close()
    return {"message": f"{deleted_count} data verifikasi berhasil dihapus", "deleted_count": deleted_count}


@app.post("/api/verified/records/bulk-delete-by-nik")
async def bulk_delete_verified_records_by_nik(niks: str = Form(...), stage_id: int = Form(None)):
    """Bulk delete verified records by comma-separated NIKs. Optionally scoped to a stage."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        nik_list = [nik.strip() for nik in niks.split(',') if nik.strip()]
        if not nik_list:
            raise HTTPException(status_code=400, detail="Tidak ada NIK yang diberikan")
        
        placeholders = ','.join(['?' for _ in nik_list])
        
        if stage_id:
            cursor.execute(f"""
                SELECT vr.id FROM verified_records vr
                JOIN verified_batches vb ON vb.id = vr.batch_id
                WHERE vr.no_ktp IN ({placeholders}) AND vb.stage_id = ?
            """, nik_list + [stage_id])
        else:
            cursor.execute(f"SELECT id FROM verified_records WHERE no_ktp IN ({placeholders})", nik_list)
        
        record_ids = [str(row['id']) for row in cursor.fetchall()]
        
        if not record_ids:
            return {"message": "Tidak ada data yang cocok untuk dihapus", "deleted_count": 0}
        
        id_placeholders = ','.join(['?' for _ in record_ids])
        id_ints = [int(rid) for rid in record_ids]
        
        cursor.execute(f"DELETE FROM replacement_events WHERE disqualified_record_id IN ({id_placeholders})", id_ints)
        cursor.execute(f"DELETE FROM verified_records WHERE id IN ({id_placeholders})", id_ints)
        deleted_count = cursor.rowcount
        
        for nik in nik_list:
            cursor.execute("DELETE FROM reconciliation_overrides WHERE original_no_ktp = ?", (nik,))
        
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal menghapus data secara bulk: {str(e)}")
    conn.close()
    return {"message": f"{deleted_count} data verifikasi berhasil dihapus", "deleted_count": deleted_count}


@app.post("/api/reconciliation/bulk-override")
async def bulk_reconciliation_override(
    stage_id: int = Form(...),
    niks: str = Form(...),
    override_type: str = Form(...)
):
    """Bulk apply reconciliation override to multiple NIKs."""
    conn = get_db_connection()
    cursor = conn.cursor()
    try:
        nik_list = [nik.strip() for nik in niks.split(',') if nik.strip()]
        if not nik_list:
            raise HTTPException(status_code=400, detail="Tidak ada NIK yang diberikan")
        
        count = 0
        for nik in nik_list:
            cursor.execute("""
                INSERT OR REPLACE INTO reconciliation_overrides (
                    stage_id, original_no_ktp, override_type, corrected_nama, corrected_no_ktp, corrected_no_kk
                ) VALUES (?, ?, ?, NULL, NULL, NULL)
            """, (stage_id, nik, override_type))
            count += cursor.rowcount
        
        conn.commit()
    except HTTPException:
        conn.close()
        raise
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=f"Gagal melakukan rekonsiliasi bulk: {str(e)}")
    conn.close()
    return {"message": f"{count} data berhasil direkonsiliasi", "reconciled_count": count}


# --- INVERS MANUAL PAIRS (rekonsiliasi manual invers ↔ verified) ---

@app.get("/api/reconciliation/unmatched-invers/{stage_id}")
def get_unmatched_invers(stage_id: int):
    """Ambil invers records yang belum punya pasangan di verified_records."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT ir.nama, ir.no_ktp, ir.no_kk,
               UPPER(TRIM(COALESCE(ir.kabupaten_kota, ''))) as kabupaten_kota,
               UPPER(TRIM(COALESCE(ir.kecamatan, ''))) as kecamatan,
               UPPER(TRIM(COALESCE(ir.desa_kelurahan, ''))) as desa_kelurahan
        FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.stage_id = ? AND irv.is_active = 1
    """, (stage_id,))
    invers_rows = [dict(r) for r in cursor.fetchall()]

    cursor.execute("""
        SELECT DISTINCT vr.no_ktp FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
    """, (stage_id,))
    verified_niks = {r['no_ktp'].strip() for r in cursor.fetchall()}

    cursor.execute("""
        SELECT invers_nik FROM invers_manual_pairs WHERE stage_id = ?
    """, (stage_id,))
    paired_niks = {r['invers_nik'].strip() for r in cursor.fetchall()}

    conn.close()

    unmatched = []
    for ir in invers_rows:
        nik = ir['no_ktp'].strip()
        if nik not in verified_niks and nik not in paired_niks:
            unmatched.append(ir)

    return {"records": unmatched, "total": len(unmatched)}


@app.get("/api/reconciliation/unmatched-verified/{stage_id}")
def get_unmatched_verified(stage_id: int):
    """Ambil verified records yang belum dipasangkan dengan invers record manapun."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT vr.id, vr.no_ktp, vr.no_kk, vr.nama, vr.desa_kelurahan,
               vr.kecamatan, vr.kabupaten_kota, vr.status
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE vb.stage_id = ?
          AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
          AND vr.id NOT IN (
            SELECT mp.verified_record_id FROM invers_manual_pairs mp WHERE mp.stage_id = ?
          )
        ORDER BY vr.nama
    """, (stage_id, stage_id))
    records = [dict(r) for r in cursor.fetchall()]
    conn.close()

    return {"records": records, "total": len(records)}


@app.get("/api/reconciliation/manual-pairs/{stage_id}")
def get_manual_pairs(stage_id: int):
    """Ambil semua manual pairs untuk suatu tahap."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("""
        SELECT mp.id, mp.invers_nik, mp.invers_nama, mp.invers_kabupaten,
               vr.nama as verified_nama, vr.no_ktp as verified_nik,
               vr.kabupaten_kota as verified_kabupaten, vr.status as verified_status
        FROM invers_manual_pairs mp
        LEFT JOIN verified_records vr ON vr.id = mp.verified_record_id
        WHERE mp.stage_id = ?
        ORDER BY mp.created_at DESC
    """, (stage_id,))
    pairs = [dict(r) for r in cursor.fetchall()]
    conn.close()

    return {"pairs": pairs, "total": len(pairs)}


@app.post("/api/reconciliation/pair-invers")
async def pair_invers_with_verified(
    stage_id: int = Form(...),
    invers_nik: str = Form(...),
    invers_nama: str = Form(...),
    invers_kabupaten: str = Form(''),
    verified_record_id: int = Form(...)
):
    """Pasangkan invers record dengan verified record."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id, no_ktp, nama, status FROM verified_records WHERE id = ?", (verified_record_id,))
    vr = cursor.fetchone()
    if not vr:
        conn.close()
        raise HTTPException(status_code=404, detail="Verified record tidak ditemukan")

    try:
        cursor.execute("""
            INSERT OR REPLACE INTO invers_manual_pairs
                (stage_id, invers_nik, invers_nama, invers_kabupaten, verified_record_id)
            VALUES (?, ?, ?, ?, ?)
        """, (stage_id, invers_nik.strip(), invers_nama, invers_kabupaten, verified_record_id))
        conn.commit()
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=400, detail=f"Gagal menyimpan pasangan: {str(e)}")

    conn.close()
    return {
        "message": f"Berhasil memasangkan NIK {invers_nik} dengan verified record ({vr['nama']}, {vr['status']})",
        "verified_nama": vr['nama'],
        "verified_status": vr['status']
    }


@app.post("/api/reconciliation/auto-pair-nik/{stage_id}")
def auto_pair_by_nik(stage_id: int):
    """Otomatis pasangkan semua invers records yang NIK-nya cocok dengan verified_records (NIK atau KK)."""
    conn = get_db_connection()
    cursor = conn.cursor()

    # Ambil semua verified records untuk stage ini (yang aktif, bukan duplicate)
    cursor.execute("""
        SELECT vr.id, vr.no_ktp, vr.no_kk, vr.nama, vr.desa_kelurahan, vr.status
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
    """, (stage_id,))
    verified_map = {}
    verified_by_kk = {}
    for r in cursor.fetchall():
        d = dict(r)
        nik = r['no_ktp'].strip()
        verified_map[nik] = d
        kk = r['no_kk'].strip() if r['no_kk'] else ''
        if kk:
            verified_by_kk[kk] = d

    # Ambil NIK yang sudah dipasang
    cursor.execute("SELECT invers_nik FROM invers_manual_pairs WHERE stage_id = ?", (stage_id,))
    paired_niks = {r['invers_nik'].strip() for r in cursor.fetchall()}

    # Ambil semua invers records
    cursor.execute("""
        SELECT ir.nama, ir.no_ktp, ir.no_kk,
               UPPER(TRIM(COALESCE(ir.kabupaten_kota, ''))) as kabupaten_kota,
               UPPER(TRIM(COALESCE(ir.kecamatan, ''))) as kecamatan,
               UPPER(TRIM(COALESCE(ir.desa_kelurahan, ''))) as desa_kelurahan
        FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.stage_id = ? AND irv.is_active = 1
    """, (stage_id,))
    invers_rows = [dict(r) for r in cursor.fetchall()]

    paired_records = []
    already_paired = 0
    no_match = 0

    for ir in invers_rows:
        nik = ir['no_ktp'].strip()
        if nik in paired_niks:
            already_paired += 1
            continue
        vr = verified_map.get(nik) or verified_by_kk.get(nik)
        if vr:
            try:
                cursor.execute("""
                    INSERT OR REPLACE INTO invers_manual_pairs
                        (stage_id, invers_nik, invers_nama, invers_kabupaten, verified_record_id)
                    VALUES (?, ?, ?, ?, ?)
                """, (stage_id, nik, ir['nama'], ir['kabupaten_kota'], vr['id']))
                paired_records.append({
                    "invers_nik": nik,
                    "invers_nama": ir['nama'],
                    "verified_nama": vr['nama'],
                    "verified_status": vr['status'],
                    "desa": ir['desa_kelurahan']
                })
            except Exception:
                pass
        else:
            no_match += 1

    conn.commit()
    conn.close()

    return {
        "message": f"Berhasil auto-pair {len(paired_records)} data. {no_match} data belum cocok.",
        "paired_count": len(paired_records),
        "already_paired": already_paired,
        "no_match": no_match,
        "paired_records": paired_records
    }


@app.get("/api/reconciliation/suggest-pairs/{stage_id}")
def suggest_pairs_by_name_desa(stage_id: int):
    """Sarankan pasangan berdasarkan kecocokan nama + desa untuk invers yang belum terpasang."""
    conn = get_db_connection()
    cursor = conn.cursor()

    # Ambil verified records, build map by (nama, desa)
    cursor.execute("""
        SELECT vr.id, vr.no_ktp, vr.no_kk, vr.nama, vr.desa_kelurahan, vr.status
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
    """, (stage_id,))
    verified_by_name_desa = {}
    for r in cursor.fetchall():
        key = (r['nama'].upper().strip(), r['desa_kelurahan'].upper().strip() if r['desa_kelurahan'] else '')
        if key not in verified_by_name_desa:
            verified_by_name_desa[key] = []
        verified_by_name_desa[key].append(dict(r))

    # Ambil NIK yang sudah dipasang
    cursor.execute("SELECT invers_nik FROM invers_manual_pairs WHERE stage_id = ?", (stage_id,))
    paired_niks = {r['invers_nik'].strip() for r in cursor.fetchall()}

    # Ambil unmatched invers
    cursor.execute("""
        SELECT ir.nama, ir.no_ktp, ir.no_kk,
               UPPER(TRIM(COALESCE(ir.kabupaten_kota, ''))) as kabupaten_kota,
               UPPER(TRIM(COALESCE(ir.kecamatan, ''))) as kecamatan,
               UPPER(TRIM(COALESCE(ir.desa_kelurahan, ''))) as desa_kelurahan
        FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.stage_id = ? AND irv.is_active = 1
    """, (stage_id,))
    invers_rows = [dict(r) for r in cursor.fetchall()]

    conn.close()

    suggestions = []
    for ir in invers_rows:
        nik = ir['no_ktp'].strip()
        if nik in paired_niks:
            continue
        key = (ir['nama'].upper().strip(), ir['desa_kelurahan'].upper().strip() if ir['desa_kelurahan'] else '')
        candidates = verified_by_name_desa.get(key, [])
        if len(candidates) == 1:
            suggestions.append({
                "invers": ir,
                "verified": candidates[0]
            })

    return {"suggestions": suggestions, "total": len(suggestions)}


@app.post("/api/reconciliation/batch-pair")
def batch_pair(body: dict = Body(...)):
    """Pasangkan beberapa invers sekaligus dengan verified records."""
    stage_id = body.get("stage_id")
    pairs = body.get("pairs", [])
    if not stage_id or not pairs:
        raise HTTPException(status_code=400, detail="stage_id dan pairs harus diisi")

    conn = get_db_connection()
    cursor = conn.cursor()
    paired_count = 0
    errors = []

    for p in pairs:
        invers_nik = p.get("invers_nik", "").strip()
        invers_nama = p.get("invers_nama", "")
        invers_kabupaten = p.get("invers_kabupaten", "")
        verified_record_id = p.get("verified_record_id")
        if not invers_nik or not verified_record_id:
            errors.append(f"Data tidak lengkap: {invers_nik}")
            continue
        try:
            cursor.execute("""
                INSERT OR REPLACE INTO invers_manual_pairs
                    (stage_id, invers_nik, invers_nama, invers_kabupaten, verified_record_id)
                VALUES (?, ?, ?, ?, ?)
            """, (stage_id, invers_nik, invers_nama, invers_kabupaten, verified_record_id))
            paired_count += 1
        except Exception as e:
            errors.append(f"Gagal pair {invers_nik}: {str(e)}")

    conn.commit()
    conn.close()

    return {
        "message": f"Berhasil memasangkan {paired_count} data.",
        "paired_count": paired_count,
        "errors": errors
    }


@app.delete("/api/reconciliation/unpair-invers/{pair_id}")
async def unpair_invers(pair_id: int):
    """Hapus pasangan invers ↔ verified."""
    conn = get_db_connection()
    cursor = conn.cursor()

    cursor.execute("SELECT id FROM invers_manual_pairs WHERE id = ?", (pair_id,))
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=404, detail="Pasangan tidak ditemukan")

    cursor.execute("DELETE FROM invers_manual_pairs WHERE id = ?", (pair_id,))
    conn.commit()
    conn.close()

    return {"message": "Pasangan berhasil dihapus"}


# --- PENGUSUL TREE EXPORT ---

@app.get("/api/stage/{stage_id}/pengusul-tree/export")
def export_pengusul_tree(stage_id: int):
    """Export pengusul hierarchy tree to Excel."""
    import re as re_mod
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT name FROM invers_stages WHERE id = ?", (stage_id,))
    stage_row = cursor.fetchone()
    stage_name = stage_row['name'] if stage_row else "Tahap"
    
    # Get tree data
    tree_data = get_pengusul_tree(stage_id)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hirarki Pengusul"
    
    header_font = Font(name='Bookman Old Style', size=11, bold=True)
    data_font = Font(name='Bookman Old Style', size=10)
    header_fill = PatternFill(start_color="0F5132", end_color="0F5132", fill_type="solid")
    header_font_white = Font(name='Bookman Old Style', size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    center_align = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # Headers
    headers = ['No', 'Pengusul', 'Kabupaten/Kota', 'Kecamatan', 'Desa/Kelurahan', 'CPB', 'Lolos', 'Tidak Lolos', 'Belum']
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border
    
    row_num = 2
    counter = 1
    
    for pengusul in tree_data:
        p_name = pengusul['name']
        for kab in pengusul.get('children', []):
            kb_name = kab['name']
            for kec in kab.get('children', []):
                kc_name = kec['name']
                for desa in kec.get('children', []):
                    ds_name = desa['name']
                    values = [counter, p_name, kb_name, kc_name, ds_name,
                              desa['cpb'], desa['lolos'], desa['tidak_lolos'], desa['belum_verifikasi']]
                    for col, val in enumerate(values, 1):
                        cell = ws.cell(row=row_num, column=col, value=val)
                        cell.font = data_font
                        cell.border = thin_border
                        cell.alignment = center_align if col in (1, 6, 7, 8, 9) else left_align
                    row_num += 1
                    counter += 1
        
        # Add summary row for pengusul
        summary_values = ['', f'TOTAL {p_name}', '', '', '',
                          pengusul['cpb'], pengusul['lolos'], pengusul['tidak_lolos'], pengusul['belum_verifikasi']]
        for col, val in enumerate(summary_values, 1):
            cell = ws.cell(row=row_num, column=col, value=val)
            cell.font = Font(name='Bookman Old Style', size=10, bold=True)
            cell.fill = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
            cell.border = thin_border
            cell.alignment = center_align if col in (1, 6, 7, 8, 9) else left_align
        row_num += 1
    
    # Column widths
    col_widths = [6, 25, 25, 25, 25, 8, 8, 12, 8]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    # Auto filter
    ws.auto_filter.ref = f"A1:I{row_num - 1}"
    
    conn.close()
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"Hirarki_Pengusul_{stage_name.replace(' ', '_')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{filename}"'}
    )


@app.get("/api/stage/{stage_id}/kabupaten-pengusul-tree")
def get_kabupaten_pengusul_tree(stage_id: int):
    """Ambil data rekap per kabupaten dengan children berupa daftar pengusul."""
    tree_data = get_pengusul_tree(stage_id)

    kab_map = {}
    for pengusul in tree_data:
        p_name = pengusul['name']
        for kab in pengusul.get('children', []):
            kb_name = kab['name']
            if kb_name not in kab_map:
                kab_map[kb_name] = {
                    "name": kb_name,
                    "cpb": 0, "lolos": 0, "tidak_lolos": 0, "belum_verifikasi": 0,
                    "children": []
                }
            node = kab_map[kb_name]
            node["cpb"] += kab['cpb']
            node["lolos"] += kab['lolos']
            node["tidak_lolos"] += kab['tidak_lolos']
            node["belum_verifikasi"] += kab['belum_verifikasi']
            node["children"].append({
                "name": p_name,
                "cpb": kab['cpb'],
                "lolos": kab['lolos'],
                "tidak_lolos": kab['tidak_lolos'],
                "belum_verifikasi": kab['belum_verifikasi']
            })

    result = sorted(kab_map.values(), key=lambda x: x['name'])
    return result


# --- END OF NEW ENDPOINTS ---

# --- REKAP KESELURUHAN (All stages, per kabupaten) ---
@app.get("/api/rekap-keseluruhan/export")
def export_rekap_keseluruhan():
    import re
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, name FROM invers_stages ORDER BY id ASC")
    all_stages = [dict(r) for r in cursor.fetchall()]
    def get_stage_num(s):
        match = re.search(r'\d+', s['name'])
        return int(match.group()) if match else 999
    all_stages = sorted(all_stages, key=get_stage_num)
    
    published_filter = "AND vb.is_published = 1"
    
    cursor.execute(f"""
        SELECT DISTINCT UPPER(TRIM(COALESCE(kabupaten_kota, ''))) as kab 
        FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.is_active = 1 AND TRIM(COALESCE(kabupaten_kota, '')) != ''
        UNION
        SELECT DISTINCT UPPER(TRIM(COALESCE(vr.kabupaten_kota, ''))) as kab
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL) AND TRIM(COALESCE(vr.kabupaten_kota, '')) != ''
        {published_filter}
        ORDER BY kab ASC
    """)
    all_kabupaten = [r['kab'] for r in cursor.fetchall()]
    
    cursor.execute("""
        SELECT 
            m.verified_stage_id as stage_id,
            UPPER(TRIM(COALESCE(vr.kabupaten_kota, ''))) as kab,
            COUNT(*) as cnt
        FROM sk_dirjen_matches m
        JOIN verified_records vr ON vr.id = m.verified_record_id
        WHERE m.verified_stage_id IS NOT NULL
        AND m.verified_record_id IS NOT NULL
        AND (m.match_type = 'PERFECT' 
            OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED')
            OR m.match_type = 'MANUAL_PAIR')
        GROUP BY m.verified_stage_id, kab
    """)
    sk_by_stage_kab = {}
    for row in cursor.fetchall():
        sid = row['stage_id']
        kab = row['kab']
        if sid not in sk_by_stage_kab:
            sk_by_stage_kab[sid] = {}
        sk_by_stage_kab[sid][kab] = row['cnt']
    
    stages_data = []
    
    for stage in all_stages:
        stage_id = stage['id']
        stage_name = stage['name']
        
        cursor.execute("""
            SELECT ir.no_ktp, UPPER(TRIM(COALESCE(ir.kabupaten_kota, ''))) as kab
            FROM invers_records ir
            JOIN invers_revisions irv ON ir.revision_id = irv.id
            WHERE irv.stage_id = ? AND irv.is_active = 1
        """, (stage_id,))
        invers_recs = [dict(r) for r in cursor.fetchall()]
        
        cursor.execute(f"""
            SELECT vr.no_ktp, vr.status, UPPER(TRIM(COALESCE(vr.kabupaten_kota, ''))) as kab
            FROM verified_records vr
            JOIN verified_batches vb ON vr.batch_id = vb.id
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
            WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
            {published_filter}
        """, (stage_id,))
        verified_recs = [dict(r) for r in cursor.fetchall()]
        
        alokasi_by_kab = {}
        for ir in invers_recs:
            kab = ir['kab'] if ir['kab'] else 'TIDAK DIKETAHUI'
            alokasi_by_kab[kab] = alokasi_by_kab.get(kab, 0) + 1
            
        verif_by_kab = {}
        for vr in verified_recs:
            kab = vr['kab'] if vr['kab'] else 'TIDAK DIKETAHUI'
            if kab not in verif_by_kab:
                verif_by_kab[kab] = {"lolos": 0, "tidak_lolos": 0}
            if vr['status'] == 'LOLOS':
                verif_by_kab[kab]['lolos'] += 1
            else:
                verif_by_kab[kab]['tidak_lolos'] += 1
                
        sk_data = sk_by_stage_kab.get(stage_id, {})
        
        kab_data = {}
        total_sk_sudah = 0
        total_sk_belum = 0
        for kab in all_kabupaten:
            alokasi = alokasi_by_kab.get(kab, 0)
            lolos = verif_by_kab.get(kab, {}).get('lolos', 0)
            tidak_lolos = verif_by_kab.get(kab, {}).get('tidak_lolos', 0)
            verifikasi = lolos + tidak_lolos
            belum = max(0, alokasi - verifikasi)
            
            sk_sudah = sk_data.get(kab, 0)
            sk_belum = max(0, lolos - sk_sudah)
            total_sk_sudah += sk_sudah
            total_sk_belum += sk_belum
            
            kab_data[kab] = {
                "alokasi": alokasi,
                "verifikasi": verifikasi,
                "lolos": lolos,
                "tidak_lolos": tidak_lolos,
                "belum": belum,
                "sk_dirjen_sudah": sk_sudah,
                "sk_dirjen_belum": sk_belum
            }
            
        stages_data.append({
            "stage_name": stage_name,
            "data": kab_data,
            "totals": {
                "alokasi": sum(alokasi_by_kab.values()),
                "lolos": sum(x['lolos'] for x in verif_by_kab.values()),
                "tidak_lolos": sum(x['tidak_lolos'] for x in verif_by_kab.values()),
                "sk_dirjen_sudah": total_sk_sudah,
                "sk_dirjen_belum": total_sk_belum
            }
        })
    
    conn.close()

    murni_stages = [s for s in stages_data if "pengganti" not in s['stage_name'].lower()]
    pengganti_stages = [s for s in stages_data if "pengganti" in s['stage_name'].lower()]

    wb = openpyxl.Workbook()

    font_title = Font(name='Segoe UI', size=15, bold=True, color='1F4E78')
    font_subtitle = Font(name='Segoe UI', size=10, italic=True, color='595959')
    font_header = Font(name='Segoe UI', size=9, bold=True, color='FFFFFF')
    font_sub_header = Font(name='Segoe UI', size=9, bold=True, color='2C3E50')
    font_data = Font(name='Segoe UI', size=9)
    font_total = Font(name='Segoe UI', size=9, bold=True)

    fill_title_group = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_stage_group = PatternFill(start_color='293241', end_color='293241', fill_type='solid')
    fill_total_row = PatternFill(start_color='EAF0F6', end_color='EAF0F6', fill_type='solid')
    fill_summary_col = PatternFill(start_color='F4F8FB', end_color='F4F8FB', fill_type='solid')

    border_thin = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    border_thick_right = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='medium', color='7F8C8D'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    border_double_top = Border(
        top=Side(style='double', color='2C3E50'),
        bottom=Side(style='thin', color='D3D3D3'),
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3')
    )

    fill_sk_sudah = PatternFill(start_color='D5F5E3', end_color='D5F5E3', fill_type='solid')
    fill_sk_belum = PatternFill(start_color='EAECEE', end_color='EAECEE', fill_type='solid')
    font_sk_sudah = Font(name='Segoe UI', size=9, bold=True, color='1E8449')
    font_sk_belum = Font(name='Segoe UI', size=9, bold=True, color='7F8C8D')

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')
    sub_headers = ["ALOKASI", "VERIFIKASI", "LOLOS", "TIDAK LOLOS", "BELUM", "SUDAH", "BELUM"]

    def build_worksheet(ws, title, subtitle, group_stages):
        ws['A1'] = title
        ws['A1'].font = font_title
        ws['A2'] = subtitle
        ws['A2'].font = font_subtitle

        ws.cell(row=4, column=1, value="No").font = font_header
        ws.cell(row=4, column=1).fill = fill_stage_group
        ws.cell(row=4, column=1).alignment = align_center
        ws.merge_cells(start_row=4, start_column=1, end_row=5, end_column=1)

        ws.cell(row=4, column=2, value="Kabupaten / Kota").font = font_header
        ws.cell(row=4, column=2).fill = fill_stage_group
        ws.cell(row=4, column=2).alignment = align_center
        ws.merge_cells(start_row=4, start_column=2, end_row=5, end_column=2)

        ws.cell(row=4, column=3, value="REKAP TOTAL").font = font_header
        ws.cell(row=4, column=3).fill = fill_title_group
        ws.cell(row=4, column=3).alignment = align_center
        ws.merge_cells(start_row=4, start_column=3, end_row=4, end_column=9)

        col_idx = 10
        for stage in group_stages:
            ws.cell(row=4, column=col_idx, value=stage['stage_name'].upper()).font = font_header
            ws.cell(row=4, column=col_idx).fill = fill_stage_group
            ws.cell(row=4, column=col_idx).alignment = align_center
            ws.merge_cells(start_row=4, start_column=col_idx, end_row=4, end_column=col_idx+6)
            col_idx += 7

        for i, sh in enumerate(sub_headers):
            cell = ws.cell(row=5, column=3+i, value=sh)
            cell.font = font_sub_header
            cell.alignment = align_center
            if sh == "SUDAH":
                cell.fill = fill_sk_sudah
                cell.font = font_sk_sudah
            elif sh == "BELUM" and i == 6:
                cell.fill = fill_sk_belum
                cell.font = font_sk_belum
            else:
                cell.fill = PatternFill(start_color='D4E2EE', end_color='D4E2EE', fill_type='solid')

        col_idx = 10
        for stage in group_stages:
            for i, sh in enumerate(sub_headers):
                cell = ws.cell(row=5, column=col_idx+i, value=sh)
                cell.font = font_sub_header
                cell.alignment = align_center
                if sh == "SUDAH":
                    cell.fill = fill_sk_sudah
                    cell.font = font_sk_sudah
                elif sh == "BELUM" and i == 6:
                    cell.fill = fill_sk_belum
                    cell.font = font_sk_belum
                else:
                    cell.fill = PatternFill(start_color='E8F0F8', end_color='E8F0F8', fill_type='solid')
            col_idx += 7

        row_idx = 6
        for idx, kab in enumerate(all_kabupaten):
            ws.cell(row=row_idx, column=1, value=idx+1).alignment = align_center
            ws.cell(row=row_idx, column=1).font = font_data
            ws.cell(row=row_idx, column=2, value=kab).font = font_total

            sum_a = sum(s['data'].get(kab, {}).get('alokasi', 0) for s in group_stages)
            sum_v = sum(s['data'].get(kab, {}).get('verifikasi', 0) for s in group_stages)
            sum_l = sum(s['data'].get(kab, {}).get('lolos', 0) for s in group_stages)
            sum_tl = sum(s['data'].get(kab, {}).get('tidak_lolos', 0) for s in group_stages)
            sum_b = sum(s['data'].get(kab, {}).get('belum', 0) for s in group_stages)
            sum_sks = sum(s['data'].get(kab, {}).get('sk_dirjen_sudah', 0) for s in group_stages)
            sum_skb = sum(s['data'].get(kab, {}).get('sk_dirjen_belum', 0) for s in group_stages)

            ws.cell(row=row_idx, column=3, value=sum_a).font = font_total
            ws.cell(row=row_idx, column=4, value=sum_v).font = font_total
            ws.cell(row=row_idx, column=5, value=sum_l).font = font_total
            ws.cell(row=row_idx, column=6, value=sum_tl).font = font_total
            ws.cell(row=row_idx, column=7, value=sum_b).font = font_total
            ws.cell(row=row_idx, column=8, value=sum_sks).font = font_sk_sudah
            ws.cell(row=row_idx, column=9, value=sum_skb).font = font_sk_belum

            for c in range(3, 10):
                cell = ws.cell(row=row_idx, column=c)
                if c == 8:
                    cell.fill = fill_sk_sudah
                elif c == 9:
                    cell.fill = fill_sk_belum
                else:
                    cell.fill = fill_summary_col
                cell.alignment = align_center
                if c == 9:
                    cell.border = border_thick_right
                else:
                    cell.border = border_thin

            col_idx = 10
            for stage in group_stages:
                kd = stage["data"].get(kab, {"alokasi": 0, "verifikasi": 0, "lolos": 0, "tidak_lolos": 0, "belum": 0, "sk_dirjen_sudah": 0, "sk_dirjen_belum": 0})
                ws.cell(row=row_idx, column=col_idx, value=kd["alokasi"] or "-").alignment = align_center
                ws.cell(row=row_idx, column=col_idx+1, value=kd["verifikasi"] or "-").alignment = align_center
                ws.cell(row=row_idx, column=col_idx+2, value=kd["lolos"] or "-").alignment = align_center
                ws.cell(row=row_idx, column=col_idx+3, value=kd["tidak_lolos"] or "-").alignment = align_center
                ws.cell(row=row_idx, column=col_idx+4, value=kd["belum"] or "-").alignment = align_center
                ws.cell(row=row_idx, column=col_idx+5, value=kd.get("sk_dirjen_sudah", 0) or "-").alignment = align_center
                ws.cell(row=row_idx, column=col_idx+6, value=kd.get("sk_dirjen_belum", 0) or "-").alignment = align_center

                for offset in range(7):
                    cell = ws.cell(row=row_idx, column=col_idx+offset)
                    cell.font = font_data
                    if offset == 5:
                        cell.fill = fill_sk_sudah
                        cell.font = font_sk_sudah
                    elif offset == 6:
                        cell.fill = fill_sk_belum
                        cell.font = font_sk_belum
                    if offset == 6:
                        cell.border = border_thick_right
                    else:
                        cell.border = border_thin
                col_idx += 7

            ws.cell(row=row_idx, column=1).border = border_thin
            ws.cell(row=row_idx, column=2).border = border_thin
            row_idx += 1

        ws.cell(row=row_idx, column=2, value="TOTAL").font = font_total
        ws.cell(row=row_idx, column=2).alignment = align_left
        ws.cell(row=row_idx, column=2).fill = fill_total_row
        ws.cell(row=row_idx, column=2).border = border_double_top
        ws.cell(row=row_idx, column=1).fill = fill_total_row
        ws.cell(row=row_idx, column=1).border = border_double_top

        tot_a = sum(sum(s['data'].get(kab, {}).get('alokasi', 0) for s in group_stages) for kab in all_kabupaten)
        tot_v = sum(sum(s['data'].get(kab, {}).get('verifikasi', 0) for s in group_stages) for kab in all_kabupaten)
        tot_l = sum(sum(s['data'].get(kab, {}).get('lolos', 0) for s in group_stages) for kab in all_kabupaten)
        tot_tl = sum(sum(s['data'].get(kab, {}).get('tidak_lolos', 0) for s in group_stages) for kab in all_kabupaten)
        tot_b = sum(sum(s['data'].get(kab, {}).get('belum', 0) for s in group_stages) for kab in all_kabupaten)
        tot_sks = sum(sum(s['data'].get(kab, {}).get('sk_dirjen_sudah', 0) for s in group_stages) for kab in all_kabupaten)
        tot_skb = sum(sum(s['data'].get(kab, {}).get('sk_dirjen_belum', 0) for s in group_stages) for kab in all_kabupaten)

        total_vals = [tot_a, tot_v, tot_l, tot_tl, tot_b, tot_sks, tot_skb]
        for i, val in enumerate(total_vals):
            cell = ws.cell(row=row_idx, column=3+i, value=val)
            cell.font = font_total
            cell.alignment = align_center
            cell.fill = fill_total_row
            if i == 5:
                cell.fill = fill_sk_sudah
                cell.font = Font(name='Segoe UI', size=9, bold=True, color='1E8449')
            elif i == 6:
                cell.fill = fill_sk_belum
                cell.font = Font(name='Segoe UI', size=9, bold=True, color='7F8C8D')
            if i == 6:
                cell.border = Border(top=Side(style='double', color='2C3E50'), bottom=Side(style='thin', color='D3D3D3'), right=Side(style='medium', color='7F8C8D'))
            else:
                cell.border = border_double_top

        col_idx = 10
        for stage in group_stages:
            st_a = sum(kd["alokasi"] for kd in stage["data"].values())
            st_v = sum(kd["verifikasi"] for kd in stage["data"].values())
            st_l = sum(kd["lolos"] for kd in stage["data"].values())
            st_tl = sum(kd["tidak_lolos"] for kd in stage["data"].values())
            st_b = sum(kd["belum"] for kd in stage["data"].values())
            st_sks = sum(kd.get("sk_dirjen_sudah", 0) for kd in stage["data"].values())
            st_skb = sum(kd.get("sk_dirjen_belum", 0) for kd in stage["data"].values())

            vals = [st_a, st_v, st_l, st_tl, st_b, st_sks, st_skb]
            for i, val in enumerate(vals):
                cell = ws.cell(row=row_idx, column=col_idx+i, value=val)
                cell.font = font_total
                cell.alignment = align_center
                cell.fill = fill_total_row
                if i == 5:
                    cell.fill = fill_sk_sudah
                    cell.font = Font(name='Segoe UI', size=9, bold=True, color='1E8449')
                elif i == 6:
                    cell.fill = fill_sk_belum
                    cell.font = Font(name='Segoe UI', size=9, bold=True, color='7F8C8D')
                if i == 6:
                    cell.border = Border(top=Side(style='double', color='2C3E50'), bottom=Side(style='thin', color='D3D3D3'), right=Side(style='medium', color='7F8C8D'))
                else:
                    cell.border = border_double_top
            col_idx += 7

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val = str(cell.value or '')
                if cell.row in [1, 2]:
                    continue
                if len(val) > max_len:
                    max_len = len(val)
            ws.column_dimensions[col_letter].width = max(max_len + 4, 10)
        ws.column_dimensions['B'].width = 28
        ws.freeze_panes = 'C6'

    ws_murni = wb.active
    ws_murni.title = "Rekap Invers Murni"
    build_worksheet(ws_murni, "REKAPITULASI INVERS MURNI", "Sistem Verifikasi Perumahan Swadaya — Tahap Invers Murni", murni_stages)

    ws_pengganti = wb.create_sheet("Rekap Invers Pengganti")
    build_worksheet(ws_pengganti, "REKAPITULASI INVERS PENGGANTI", "Sistem Verifikasi Perumahan Swadaya — Tahap Invers Pengganti", pengganti_stages)

    ws_keseluruhan = wb.create_sheet("Rekap Keseluruhan")
    build_worksheet(ws_keseluruhan, "REKAPITULASI KESELURUHAN INVERS", "Sistem Verifikasi Perumahan Swadaya — Semua Tahap", stages_data)

    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)

    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=REKAP_KESELURUHAN_INVERS.xlsx"}
    )

# --- REKAP KESELURUHAN (All stages, per kabupaten) ---
@app.get("/api/rekap-keseluruhan")
def get_rekap_keseluruhan(published_only: int = 0):
    import re
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # Get all stages
    cursor.execute("SELECT id, name FROM invers_stages ORDER BY id ASC")
    all_stages = [dict(r) for r in cursor.fetchall()]
    def get_stage_num(s):
        match = re.search(r'\d+', s['name'])
        return int(match.group()) if match else 999
    all_stages = sorted(all_stages, key=get_stage_num)
    
    # Build optional published-only filter for verified queries
    published_filter = "AND vb.is_published = 1" if published_only else ""
    
    # Get the full list of unique kabupaten across all stages
    cursor.execute(f"""
        SELECT DISTINCT UPPER(TRIM(COALESCE(kabupaten_kota, ''))) as kab 
        FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.is_active = 1 AND TRIM(COALESCE(kabupaten_kota, '')) != ''
        UNION
        SELECT DISTINCT UPPER(TRIM(COALESCE(vr.kabupaten_kota, ''))) as kab
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL) AND TRIM(COALESCE(vr.kabupaten_kota, '')) != ''
        {published_filter}
        ORDER BY kab ASC
    """)
    all_kabupaten = [r['kab'] for r in cursor.fetchall()]
    
    # Pre-fetch SK Dirjen match counts per stage per kabupaten
    cursor.execute("""
        SELECT 
            m.verified_stage_id as stage_id,
            UPPER(TRIM(COALESCE(vr.kabupaten_kota, ''))) as kab,
            COUNT(*) as cnt
        FROM sk_dirjen_matches m
        JOIN verified_records vr ON vr.id = m.verified_record_id
        WHERE m.verified_stage_id IS NOT NULL
        AND m.verified_record_id IS NOT NULL
        AND (m.match_type = 'PERFECT' 
            OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED')
            OR m.match_type = 'MANUAL_PAIR')
        GROUP BY m.verified_stage_id, kab
    """)
    sk_by_stage_kab = {}
    for row in cursor.fetchall():
        sid = row['stage_id']
        kab = row['kab']
        if sid not in sk_by_stage_kab:
            sk_by_stage_kab[sid] = {}
        sk_by_stage_kab[sid][kab] = row['cnt']
    
    stages_data = []
    
    for stage in all_stages:
        stage_id = stage['id']
        stage_name = stage['name']
        
        # Get INVERS records for this stage (alokasi)
        cursor.execute("""
            SELECT ir.no_ktp, UPPER(TRIM(COALESCE(ir.kabupaten_kota, ''))) as kab
            FROM invers_records ir
            JOIN invers_revisions irv ON ir.revision_id = irv.id
            WHERE irv.stage_id = ? AND irv.is_active = 1
        """, (stage_id,))
        invers_recs = [dict(r) for r in cursor.fetchall()]
        
        # Get verified records for this stage (non-duplicate OR reconciled)
        cursor.execute(f"""
            SELECT vr.no_ktp, vr.status, UPPER(TRIM(COALESCE(vr.kabupaten_kota, ''))) as kab,
                   vr.is_duplicate_in_previous
            FROM verified_records vr
            JOIN verified_batches vb ON vr.batch_id = vb.id
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
            WHERE vb.stage_id = ? AND (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL)
            {published_filter}
        """, (stage_id,))
        verified_recs = [dict(r) for r in cursor.fetchall()]
        
        # Build alokasi per kabupaten
        alokasi_by_kab = {}
        invers_niks = set()
        for ir in invers_recs:
            kab = ir['kab'] if ir['kab'] else 'TIDAK DIKETAHUI'
            alokasi_by_kab[kab] = alokasi_by_kab.get(kab, 0) + 1
            invers_niks.add(ir['no_ktp'])
        
        # Build verifikasi stats per kabupaten
        verif_by_kab = {}
        verified_niks = set()
        for vr in verified_recs:
            kab = vr['kab'] if vr['kab'] else 'TIDAK DIKETAHUI'
            if kab not in verif_by_kab:
                verif_by_kab[kab] = {"lolos": 0, "tidak_lolos": 0}
            if vr['status'] == 'LOLOS':
                verif_by_kab[kab]['lolos'] += 1
            else:
                verif_by_kab[kab]['tidak_lolos'] += 1
            verified_niks.add(vr['no_ktp'])
        
        # Per-kabupaten table data
        kab_data = []
        total_alokasi = 0
        total_verifikasi = 0
        total_lolos = 0
        total_tidak_lolos = 0
        total_belum = 0
        total_sk_sudah = 0
        total_sk_belum = 0
        
        sk_data = sk_by_stage_kab.get(stage_id, {})
        
        for kab in all_kabupaten:
            alokasi = alokasi_by_kab.get(kab, 0)
            lolos = verif_by_kab.get(kab, {}).get('lolos', 0)
            tidak_lolos = verif_by_kab.get(kab, {}).get('tidak_lolos', 0)
            verifikasi = lolos + tidak_lolos
            belum = max(0, alokasi - verifikasi)
            
            sk_sudah = sk_data.get(kab, 0)
            sk_belum = max(0, lolos - sk_sudah)
            
            total_alokasi += alokasi
            total_verifikasi += verifikasi
            total_lolos += lolos
            total_tidak_lolos += tidak_lolos
            total_belum += belum
            total_sk_sudah += sk_sudah
            total_sk_belum += sk_belum
            
            kab_data.append({
                "kabupaten": kab,
                "alokasi": alokasi,
                "verifikasi": verifikasi,
                "lolos": lolos,
                "tidak_lolos": tidak_lolos,
                "belum_verifikasi": belum,
                "sk_dirjen_sudah": sk_sudah,
                "sk_dirjen_belum": sk_belum
            })
        
        # Progress bar segments for this stage
        stage_type = "pengganti" if "pengganti" in stage_name.lower() else "murni"
        stages_data.append({
            "stage_id": stage_id,
            "stage_name": stage_name,
            "stage_type": stage_type,
            "kabupaten_data": kab_data,
            "totals": {
                "alokasi": total_alokasi,
                "verifikasi": total_verifikasi,
                "lolos": total_lolos,
                "tidak_lolos": total_tidak_lolos,
                "belum_verifikasi": total_belum,
                "sk_dirjen_sudah": total_sk_sudah,
                "sk_dirjen_belum": total_sk_belum
            }
        })
    
    return {
        "all_kabupaten": all_kabupaten,
        "stages": stages_data
    }

# --- REKAP BATCH BERITA ACARA ---
@app.get("/api/rekap-batch-ba")
def get_rekap_batch_ba(published_only: int = 1):
    import re
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Get all stages
    cursor.execute("SELECT id, name FROM invers_stages ORDER BY id ASC")
    all_stages = [dict(r) for r in cursor.fetchall()]
    def get_stage_num(s):
        match = re.search(r'\d+', s['name'])
        return int(match.group()) if match else 999
    all_stages = sorted(all_stages, key=get_stage_num)
    
    # 2. Get list of kabupaten/kota
    published_filter = "AND vb.is_published = 1" if published_only else ""
    cursor.execute(f"""
        SELECT DISTINCT UPPER(TRIM(COALESCE(kabupaten_kota, ''))) as kab 
        FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.is_active = 1 AND TRIM(COALESCE(kabupaten_kota, '')) != ''
        UNION
        SELECT DISTINCT UPPER(TRIM(COALESCE(vr.kabupaten_kota, ''))) as kab
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL) AND TRIM(COALESCE(vr.kabupaten_kota, '')) != ''
        {published_filter}
        ORDER BY kab ASC
    """)
    all_kabupaten = [r['kab'] for r in cursor.fetchall()]
    
    # 3. For each stage, get its batches
    stage_batch_filter = "AND is_published = 1" if published_only else ""
    stages_data = []
    
    for stage in all_stages:
        stage_id = stage['id']
        stage_name = stage['name']
        
        cursor.execute(f"""
            SELECT id, name, is_published, nomor_ba, tanggal_ba, sort_order 
            FROM verified_batches 
            WHERE stage_id = ? {stage_batch_filter}
            ORDER BY sort_order ASC, uploaded_at ASC, id ASC
        """, (stage_id,))
        batches = [dict(r) for r in cursor.fetchall()]
        
        if not batches:
            continue
            
        batches_data = []
        for batch in batches:
            batch_id = batch['id']
            batch_name = batch['name']
            
            cursor.execute("""
                SELECT UPPER(TRIM(COALESCE(kabupaten_kota, ''))) as kab,
                       SUM(CASE WHEN status = 'LOLOS' THEN 1 ELSE 0 END) as lolos,
                       SUM(CASE WHEN status = 'TIDAK LOLOS' THEN 1 ELSE 0 END) as tidak_lolos
                FROM verified_records
                WHERE batch_id = ? AND TRIM(COALESCE(kabupaten_kota, '')) != ''
                GROUP BY kab
            """, (batch_id,))
            
            stats_by_kab = {}
            for row in cursor.fetchall():
                stats_by_kab[row['kab']] = {
                    "lolos": row['lolos'],
                    "tidak_lolos": row['tidak_lolos'],
                    "verifikasi": row['lolos'] + row['tidak_lolos']
                }
            
            kab_data = []
            total_verifikasi = 0
            total_lolos = 0
            total_tidak_lolos = 0
            
            for kab in all_kabupaten:
                stats = stats_by_kab.get(kab, {"lolos": 0, "tidak_lolos": 0, "verifikasi": 0})
                kab_data.append({
                    "kabupaten": kab,
                    "verifikasi": stats["verifikasi"],
                    "lolos": stats["lolos"],
                    "tidak_lolos": stats["tidak_lolos"]
                })
                total_verifikasi += stats["verifikasi"]
                total_lolos += stats["lolos"]
                total_tidak_lolos += stats["tidak_lolos"]
                
            batches_data.append({
                "batch_id": batch_id,
                "batch_name": batch_name,
                "is_published": batch["is_published"],
                "nomor_ba": batch.get("nomor_ba"),
                "tanggal_ba": batch.get("tanggal_ba"),
                "kabupaten_data": kab_data,
                "totals": {
                    "verifikasi": total_verifikasi,
                    "lolos": total_lolos,
                    "tidak_lolos": total_tidak_lolos
                }
            })
            
        stage_type = "pengganti" if "pengganti" in stage_name.lower() else "murni"
        stages_data.append({
            "stage_id": stage_id,
            "stage_name": stage_name,
            "stage_type": stage_type,
            "batches": batches_data
        })
        
    conn.close()
    return {
        "all_kabupaten": all_kabupaten,
        "stages": stages_data
    }

@app.get("/api/rekap-batch-ba/export")
def export_rekap_batch_ba(published_only: int = 1):
    import re
    conn = get_db_connection()
    cursor = conn.cursor()
    
    # 1. Get all stages
    cursor.execute("SELECT id, name FROM invers_stages ORDER BY id ASC")
    all_stages = [dict(r) for r in cursor.fetchall()]
    def get_stage_num(s):
        match = re.search(r'\d+', s['name'])
        return int(match.group()) if match else 999
    all_stages = sorted(all_stages, key=get_stage_num)
    
    # 2. Get list of kabupaten/kota
    published_filter = "AND vb.is_published = 1" if published_only else ""
    cursor.execute(f"""
        SELECT DISTINCT UPPER(TRIM(COALESCE(kabupaten_kota, ''))) as kab 
        FROM invers_records ir
        JOIN invers_revisions irv ON ir.revision_id = irv.id
        WHERE irv.is_active = 1 AND TRIM(COALESCE(kabupaten_kota, '')) != ''
        UNION
        SELECT DISTINCT UPPER(TRIM(COALESCE(vr.kabupaten_kota, ''))) as kab
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
        WHERE (vr.is_duplicate_in_previous = 0 OR ro.id IS NOT NULL) AND TRIM(COALESCE(vr.kabupaten_kota, '')) != ''
        {published_filter}
        ORDER BY kab ASC
    """)
    all_kabupaten = [r['kab'] for r in cursor.fetchall()]
    
    stage_batch_filter = "AND is_published = 1" if published_only else ""
    stages_data = []
    
    for stage in all_stages:
        stage_id = stage['id']
        stage_name = stage['name']
        
        cursor.execute(f"""
            SELECT id, name, nomor_ba, tanggal_ba 
            FROM verified_batches 
            WHERE stage_id = ? {stage_batch_filter}
            ORDER BY sort_order ASC, uploaded_at ASC, id ASC
        """, (stage_id,))
        batches = [dict(r) for r in cursor.fetchall()]
        
        if not batches:
            continue
            
        batches_data = []
        for batch in batches:
            batch_id = batch['id']
            batch_name = batch['name']
            
            cursor.execute("""
                SELECT UPPER(TRIM(COALESCE(kabupaten_kota, ''))) as kab,
                       SUM(CASE WHEN status = 'LOLOS' THEN 1 ELSE 0 END) as lolos,
                       SUM(CASE WHEN status = 'TIDAK LOLOS' THEN 1 ELSE 0 END) as tidak_lolos
                FROM verified_records
                WHERE batch_id = ? AND TRIM(COALESCE(kabupaten_kota, '')) != ''
                GROUP BY kab
            """, (batch_id,))
            
            stats_by_kab = {}
            for row in cursor.fetchall():
                stats_by_kab[row['kab']] = {
                    "lolos": row['lolos'],
                    "tidak_lolos": row['tidak_lolos'],
                    "verifikasi": row['lolos'] + row['tidak_lolos']
                }
            
            kab_data = {}
            total_verifikasi = 0
            total_lolos = 0
            total_tidak_lolos = 0
            for kab in all_kabupaten:
                stats = stats_by_kab.get(kab, {"lolos": 0, "tidak_lolos": 0, "verifikasi": 0})
                kab_data[kab] = {
                    "verifikasi": stats["verifikasi"],
                    "lolos": stats["lolos"],
                    "tidak_lolos": stats["tidak_lolos"]
                }
                total_verifikasi += stats["verifikasi"]
                total_lolos += stats["lolos"]
                total_tidak_lolos += stats["tidak_lolos"]
                
            batches_data.append({
                "batch_id": batch_id,
                "batch_name": batch_name,
                "nomor_ba": batch.get("nomor_ba"),
                "tanggal_ba": batch.get("tanggal_ba"),
                "data": kab_data,
                "totals": {
                    "verifikasi": total_verifikasi,
                    "lolos": total_lolos,
                    "tidak_lolos": total_tidak_lolos
                }
            })
            
        stage_type = "pengganti" if "pengganti" in stage_name.lower() else "murni"
        stages_data.append({
            "stage_name": stage_name,
            "stage_type": stage_type,
            "batches": batches_data
        })
        
    conn.close()
    
    murni_stages = [s for s in stages_data if s['stage_type'] == 'murni']
    pengganti_stages = [s for s in stages_data if s['stage_type'] == 'pengganti']
    
    wb = openpyxl.Workbook()
    
    font_title = Font(name='Segoe UI', size=15, bold=True, color='1F4E78')
    font_subtitle = Font(name='Segoe UI', size=10, italic=True, color='595959')
    font_header_corner = Font(name='Segoe UI', size=9, bold=True, color='2C3E50')
    font_header_l1 = Font(name='Segoe UI', size=9, bold=True, color='FFFFFF')
    font_header_l2 = Font(name='Segoe UI', size=8, bold=True, color='2C3E50')
    font_header_l3 = Font(name='Segoe UI', size=8, bold=True, color='595959')
    
    font_data = Font(name='Segoe UI', size=9)
    font_total = Font(name='Segoe UI', size=9, bold=True)
    
    fill_corner = PatternFill(start_color='EAF0F6', end_color='EAF0F6', fill_type='solid')
    fill_l1 = PatternFill(start_color='1F4E78', end_color='1F4E78', fill_type='solid')
    fill_l2 = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
    fill_l3 = PatternFill(start_color='F2F4F4', end_color='F2F4F4', fill_type='solid')
    fill_total_row = PatternFill(start_color='EAF0F6', end_color='EAF0F6', fill_type='solid')
    
    fill_lolos_cell = PatternFill(start_color='E8F5E9', end_color='E8F5E9', fill_type='solid')
    fill_tidak_lolos_cell = PatternFill(start_color='FDEDEC', end_color='FDEDEC', fill_type='solid')
    font_lolos_cell = Font(name='Segoe UI', size=9, color='2E7D32')
    font_tidak_lolos_cell = Font(name='Segoe UI', size=9, color='C0392B')
    
    border_thin = Border(
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3'),
        top=Side(style='thin', color='D3D3D3'),
        bottom=Side(style='thin', color='D3D3D3')
    )
    border_double_top = Border(
        top=Side(style='double', color='2C3E50'),
        bottom=Side(style='thin', color='D3D3D3'),
        left=Side(style='thin', color='D3D3D3'),
        right=Side(style='thin', color='D3D3D3')
    )
    
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center')
    
    def build_worksheet(ws, title, subtitle, group_stages):
        ws['A1'] = title
        ws['A1'].font = font_title
        ws['A2'] = subtitle
        ws['A2'].font = font_subtitle
        
        cell_no = ws.cell(row=4, column=1, value="No")
        cell_no.font = font_header_corner
        cell_no.fill = fill_corner
        cell_no.alignment = align_center
        ws.merge_cells(start_row=4, start_column=1, end_row=6, end_column=1)
        
        cell_kab = ws.cell(row=4, column=2, value="Kabupaten / Kota")
        cell_kab.font = font_header_corner
        cell_kab.fill = fill_corner
        cell_kab.alignment = align_center
        ws.merge_cells(start_row=4, start_column=2, end_row=6, end_column=2)
        
        col_idx = 3
        for stage in group_stages:
            num_batches = len(stage['batches'])
            if num_batches == 0:
                continue
                
            stage_width = num_batches * 3
            cell_stage = ws.cell(row=4, column=col_idx, value=stage['stage_name'].upper())
            cell_stage.font = font_header_l1
            cell_stage.fill = fill_l1
            cell_stage.alignment = align_center
            ws.merge_cells(start_row=4, start_column=col_idx, end_row=4, end_column=col_idx + stage_width - 1)
            
            b_col_idx = col_idx
            for batch in stage['batches']:
                val_str = f"{batch['batch_name'].upper()}\nNo: {batch.get('nomor_ba') or '—'}\nTgl: {batch.get('tanggal_ba') or '—'}"
                cell_batch = ws.cell(row=5, column=b_col_idx, value=val_str)
                cell_batch.font = font_header_l2
                cell_batch.fill = fill_l2
                cell_batch.alignment = align_center
                ws.merge_cells(start_row=5, start_column=b_col_idx, end_row=5, end_column=b_col_idx + 2)
                ws.row_dimensions[5].height = 42
                
                metrics = ["VERIFIKASI", "LOLOS", "TIDAK LOLOS"]
                for i, m in enumerate(metrics):
                    cell_m = ws.cell(row=6, column=b_col_idx + i, value=m)
                    cell_m.font = font_header_l3
                    cell_m.fill = fill_l3
                    cell_m.alignment = align_center
                    
                b_col_idx += 3
            col_idx += stage_width
            
        max_col = 2
        for stage in group_stages:
            max_col += len(stage['batches']) * 3
            
        for r in range(4, 7):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = border_thin
                
        row_idx = 7
        for idx, kab in enumerate(all_kabupaten):
            ws.cell(row=row_idx, column=1, value=idx + 1).alignment = align_center
            ws.cell(row=row_idx, column=1).font = font_data
            ws.cell(row=row_idx, column=1).border = border_thin
            
            ws.cell(row=row_idx, column=2, value=kab).font = font_total
            ws.cell(row=row_idx, column=2).border = border_thin
            
            c_idx = 3
            for stage in group_stages:
                for batch in stage['batches']:
                    kd = batch['data'].get(kab, {"verifikasi": 0, "lolos": 0, "tidak_lolos": 0})
                    val_v = kd['verifikasi']
                    val_l = kd['lolos']
                    val_tl = kd['tidak_lolos']
                    
                    cell_v = ws.cell(row=row_idx, column=c_idx, value=val_v or "-")
                    cell_l = ws.cell(row=row_idx, column=c_idx + 1, value=val_l or "-")
                    cell_tl = ws.cell(row=row_idx, column=c_idx + 2, value=val_tl or "-")
                    
                    cell_v.alignment = align_center
                    cell_l.alignment = align_center
                    cell_tl.alignment = align_center
                    
                    cell_v.font = font_data
                    cell_l.font = font_lolos_cell if val_l > 0 else font_data
                    cell_tl.font = font_tidak_lolos_cell if val_tl > 0 else font_data
                    
                    if val_l > 0:
                        cell_l.fill = fill_lolos_cell
                    if val_tl > 0:
                        cell_tl.fill = fill_tidak_lolos_cell
                        
                    cell_v.border = border_thin
                    cell_l.border = border_thin
                    cell_tl.border = border_thin
                    c_idx += 3
            row_idx += 1
            
        cell_tot_label = ws.cell(row=row_idx, column=2, value="TOTAL")
        cell_tot_label.font = font_total
        cell_tot_label.alignment = align_left
        cell_tot_label.fill = fill_total_row
        cell_tot_label.border = border_double_top
        
        cell_tot_no = ws.cell(row=row_idx, column=1)
        cell_tot_no.fill = fill_total_row
        cell_tot_no.border = border_double_top
        
        c_idx = 3
        for stage in group_stages:
            for batch in stage['batches']:
                t = batch['totals']
                cell_tot_v = ws.cell(row=row_idx, column=c_idx, value=t['verifikasi'])
                cell_tot_l = ws.cell(row=row_idx, column=c_idx + 1, value=t['lolos'])
                cell_tot_tl = ws.cell(row=row_idx, column=c_idx + 2, value=t['tidak_lolos'])
                
                cell_tot_v.font = font_total
                cell_tot_l.font = font_total
                cell_tot_tl.font = font_total
                
                cell_tot_v.alignment = align_center
                cell_tot_l.alignment = align_center
                cell_tot_tl.alignment = align_center
                
                cell_tot_v.fill = fill_total_row
                cell_tot_l.fill = fill_total_row
                cell_tot_tl.fill = fill_total_row
                
                cell_tot_v.border = border_double_top
                cell_tot_l.border = border_double_top
                cell_tot_tl.border = border_double_top
                c_idx += 3
                
        for col in ws.columns:
            col_idx = col[0].column
            col_letter = get_column_letter(col_idx)
            if col_idx == 1:
                ws.column_dimensions[col_letter].width = 6
            elif col_idx == 2:
                ws.column_dimensions[col_letter].width = 28
            else:
                ws.column_dimensions[col_letter].width = 14
        ws.freeze_panes = 'C7'
        
    ws_murni = wb.active
    ws_murni.title = "Rekap Batch Murni"
    build_worksheet(ws_murni, "REKAPITULASI BATCH INVERS MURNI", "Sistem Verifikasi Perumahan Swadaya — Laporan per Batch Murni", murni_stages)
    
    if pengganti_stages:
        ws_pengganti = wb.create_sheet("Rekap Batch Pengganti")
        build_worksheet(ws_pengganti, "REKAPITULASI BATCH INVERS PENGGANTI", "Sistem Verifikasi Perumahan Swadaya — Laporan per Batch Pengganti", pengganti_stages)
        
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    
    filename = "REKAP_BATCH_BERITA_ACARA.xlsx"
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# --- GLOBAL SEARCH (Cross-stage search for invers + verified records) ---
@app.get("/api/global-search")
def global_search(
    q: str = "",
    kabupaten: str = "",
    kecamatan: str = "",
    desa: str = "",
    status: str = "",
    tahap: str = "",
    record_type: str = "all",
    page: int = 1,
    limit: int = 30,
    export_all: bool = False
):
    conn = get_db_connection()
    cursor = conn.cursor()

    conditions = []
    params = []

    # --- Build WHERE clauses ---
    if q and len(q.strip()) >= 1:
        term = f"%{q.strip().upper()}%"
        conditions.append("(UPPER(vr.nama) LIKE ? OR vr.no_ktp LIKE ? OR vr.no_kk LIKE ?)")
        params.extend([term, term, term])

    if kabupaten:
        conditions.append("UPPER(vr.kabupaten_kota) = UPPER(?)")
        params.append(kabupaten)
    if kecamatan:
        conditions.append("UPPER(vr.kecamatan) = UPPER(?)")
        params.append(kecamatan)
    if desa:
        conditions.append("UPPER(vr.desa_kelurahan) = UPPER(?)")
        params.append(desa)
    if tahap:
        conditions.append("ist.id = ?")
        params.append(int(tahap))

    where_clause = " AND ".join(conditions) if conditions else "1=1"

    # Status filter for verified records
    verified_status_clause = ""
    verified_status_params = []
    if status == "LOLOS":
        verified_status_clause = "AND vr.status = 'LOLOS'"
    elif status == "TIDAK_LOLOS":
        verified_status_clause = "AND vr.status = 'TIDAK LOLOS'"

    # Status filter for invers (unverified)
    invers_status_clause = ""
    if status == "BELUM":
        invers_status_clause = "AND ir.no_ktp NOT IN (SELECT no_ktp FROM verified_records WHERE batch_id IN (SELECT id FROM verified_batches WHERE stage_id = ist.id))"
    elif status:
        # If status is set but not BELUM, no invers records match
        invers_status_clause = "AND 1=0"

    all_records = []
    summary = {"total_alokasi": 0, "total_verifikasi": 0, "total_lolos": 0, "total_tidak_lolos": 0, "total_belum": 0}

    # --- Verified records ---
    if record_type in ("all", "verified"):
        query = f"""
            SELECT vr.id, vr.nama, vr.no_ktp, vr.no_kk, vr.kabupaten_kota, vr.kecamatan,
                   vr.desa_kelurahan, vr.status, ist.id as tahap_id, ist.name as tahap_name,
                   vb.name as batch_name, 'verified' as record_type
            FROM verified_records vr
            JOIN verified_batches vb ON vb.id = vr.batch_id
            JOIN invers_stages ist ON ist.id = vb.stage_id
            WHERE {where_clause}
            {verified_status_clause}
            ORDER BY vr.nama ASC
        """
        cursor.execute(query, params)
        for row in cursor.fetchall():
            rec = dict(row)
            rec['status'] = rec['status'] if rec['status'] else 'SELESAI'
            all_records.append(rec)

    # --- Invers records (belum diverifikasi) ---
    if record_type in ("all", "invers"):
        invers_conditions = []
        invers_params = []
        if q and len(q.strip()) >= 1:
            term = f"%{q.strip().upper()}%"
            invers_conditions.append("(UPPER(ir.nama) LIKE ? OR ir.no_ktp LIKE ? OR ir.no_kk LIKE ?)")
            invers_params.extend([term, term, term])
        if kabupaten:
            invers_conditions.append("UPPER(ir.kabupaten_kota) = UPPER(?)")
            invers_params.append(kabupaten)
        if kecamatan:
            invers_conditions.append("UPPER(ir.kecamatan) = UPPER(?)")
            invers_params.append(kecamatan)
        if desa:
            invers_conditions.append("UPPER(ir.desa_kelurahan) = UPPER(?)")
            invers_params.append(desa)
        if tahap:
            invers_conditions.append("irv.stage_id = ?")
            invers_params.append(int(tahap))

        invers_where = " AND ".join(invers_conditions) if invers_conditions else "1=1"

        invers_query = f"""
            SELECT ir.id, ir.nama, ir.no_ktp, ir.no_kk, ir.kabupaten_kota, ir.kecamatan,
                   ir.desa_kelurahan, irv.stage_id as tahap_id, ist.name as tahap_name,
                   'Belum Diverifikasi' as batch_name, 'invers' as record_type
            FROM invers_records ir
            JOIN invers_revisions irv ON ir.revision_id = irv.id
            JOIN invers_stages ist ON ist.id = irv.stage_id
            WHERE irv.is_active = 1
              AND {invers_where}
              AND ir.no_ktp NOT IN (
                  SELECT vr2.no_ktp FROM verified_records vr2
                  JOIN verified_batches vb2 ON vb2.id = vr2.batch_id
                  WHERE vb2.stage_id = irv.stage_id
              )
            ORDER BY ir.nama ASC
        """
        cursor.execute(invers_query, invers_params)
        for row in cursor.fetchall():
            rec = dict(row)
            rec['status'] = 'BELUM'
            all_records.append(rec)

    conn.close()

    # Sort: verified first, then by name
    all_records.sort(key=lambda r: (0 if r['record_type'] == 'verified' else 1, r['nama'] or ''))

    # Summary from ALL records (before pagination)
    summary['total_alokasi'] = len(all_records)
    summary['total_verifikasi'] = len([r for r in all_records if r['record_type'] == 'verified'])
    summary['total_lolos'] = len([r for r in all_records if r.get('status') == 'LOLOS'])
    summary['total_tidak_lolos'] = len([r for r in all_records if r.get('status') == 'TIDAK LOLOS'])
    summary['total_belum'] = len([r for r in all_records if r['record_type'] == 'invers'])

    total = len(all_records)
    total_pages = max(1, (total + limit - 1) // limit)

    if export_all:
        return {"records": all_records, "total": total, "page": 1, "total_pages": 1, "summary": summary}

    start = (page - 1) * limit
    paginated = all_records[start:start + limit]

    # Fetch filter options (cascading) based on current records
    all_kabs = sorted(set(r['kabupaten_kota'] for r in all_records if r.get('kabupaten_kota')))
    all_kecs = sorted(set(r['kecamatan'] for r in all_records if r.get('kecamatan')))
    all_desas = sorted(set(r['desa_kelurahan'] for r in all_records if r.get('desa_kelurahan')))
    all_tahaps_raw = {}
    for r in all_records:
        tid = r.get('tahap_id')
        tname = r.get('tahap_name')
        if tid and tid not in all_tahaps_raw:
            all_tahaps_raw[tid] = tname
    all_tahaps = [{"id": k, "name": v} for k, v in sorted(all_tahaps_raw.items())]

    return {
        "records": paginated,
        "total": total,
        "page": page,
        "total_pages": total_pages,
        "summary": summary,
        "filters": {
            "kabupatens": all_kabs,
            "kecamatans": all_kecs,
            "desas": all_desas,
            "tahaps": all_tahaps
        }
    }

@app.get("/api/global-search/export")
def global_search_export(
    q: str = "",
    kabupaten: str = "",
    kecamatan: str = "",
    desa: str = "",
    status: str = "",
    tahap: str = "",
    record_type: str = "all"
):
    result = global_search(q=q, kabupaten=kabupaten, kecamatan=kecamatan, desa=desa,
                           status=status, tahap=tahap, record_type=record_type,
                           page=1, limit=99999, export_all=True)
    records = result["records"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Pencarian Global"

    headers = ["No", "Tahap", "Nama", "NIK", "No KK", "Kabupaten/Kota", "Kecamatan", "Desa/Kelurahan", "Status", "Asal"]
    header_font = openpyxl.styles.Font(bold=True, color="FFFFFF")
    header_fill = openpyxl.styles.PatternFill(start_color="1A3C40", end_color="1A3C40", fill_type="solid")
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = openpyxl.styles.Alignment(horizontal="center")

    for row_idx, rec in enumerate(records, 2):
        ws.cell(row=row_idx, column=1, value=row_idx - 1)
        ws.cell(row=row_idx, column=2, value=rec.get("tahap_name", ""))
        ws.cell(row=row_idx, column=3, value=rec.get("nama", ""))
        ws.cell(row=row_idx, column=4, value=rec.get("no_ktp", ""))
        ws.cell(row=row_idx, column=5, value=rec.get("no_kk", ""))
        ws.cell(row=row_idx, column=6, value=rec.get("kabupaten_kota", ""))
        ws.cell(row=row_idx, column=7, value=rec.get("kecamatan", ""))
        ws.cell(row=row_idx, column=8, value=rec.get("desa_kelurahan", ""))
        ws.cell(row=row_idx, column=9, value=rec.get("status", ""))
        ws.cell(row=row_idx, column=10, value="Terverifikasi" if rec["record_type"] == "verified" else "Belum Diverifikasi")

    for col in ws.columns:
        max_len = 0
        col_letter = col[0].column_letter
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 40)

    import io
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    filename = f"Pencarian_Global_{kabupaten or 'Semua'}_{status or 'Semua'}_{page}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/api/stage/{stage_id}/export")
def export_excel(stage_id: int, batch_id: int = None):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT name FROM invers_stages WHERE id = ?", (stage_id,))
    stage_name = cursor.fetchone()['name']
    
    if batch_id:
        cursor.execute("SELECT name FROM verified_batches WHERE id = ?", (batch_id,))
        batch_row = cursor.fetchone()
        batch_title = batch_row['name'] if batch_row else "Berita Acara"
        
        cursor.execute("""
            SELECT vr.*, re.nama_pengganti, re.no_ktp_pengganti, re.no_kk_pengganti, re.alamat_pengganti,
                   re.desa_kelurahan_pengganti, re.kecamatan_pengganti, re.kabupaten_pengganti,
                   re.jenis_kelamin_pengganti,
                   ro.id as override_id
            FROM verified_records vr
            LEFT JOIN replacement_events re ON re.disqualified_record_id = vr.id
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = ?
            WHERE vr.batch_id = ?
        """, (stage_id, batch_id))
    else:
        batch_title = "Semua Berita Acara"
        cursor.execute("""
            SELECT vr.*, re.nama_pengganti, re.no_ktp_pengganti, re.no_kk_pengganti, re.alamat_pengganti,
                   re.desa_kelurahan_pengganti, re.kecamatan_pengganti, re.kabupaten_pengganti,
                   re.jenis_kelamin_pengganti,
                   ro.id as override_id
            FROM verified_records vr
            JOIN verified_batches vb ON vr.batch_id = vb.id
            LEFT JOIN replacement_events re ON re.disqualified_record_id = vr.id
            LEFT JOIN reconciliation_overrides ro ON ro.original_no_ktp = vr.no_ktp AND ro.stage_id = vb.stage_id
            WHERE vb.stage_id = ?
        """, (stage_id,))
        
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    # Filter: include non-duplicates OR reconciled duplicates
    def should_include_in_ba(r):
        if r['is_duplicate_in_previous'] == 0:
            return True
        return r.get('override_id') is not None
    
    lolos_records = [r for r in records if r['status'] == 'LOLOS' and should_include_in_ba(r)]
    tidak_lolos_records = [r for r in records if r['status'] == 'TIDAK LOLOS' and should_include_in_ba(r)]
    
    summary_by_kab = {}
    for r in records:
        if not should_include_in_ba(r):
            continue
        kab = (r['kabupaten_kota'] or "KAB. LUWU UTARA").upper().strip()
        if kab not in summary_by_kab:
            summary_by_kab[kab] = {"cpb": 0, "lolos": 0, "tidak_lolos": 0, "pengganti": 0}
            
        summary_by_kab[kab]["cpb"] += 1
        if r['status'] == 'LOLOS':
            summary_by_kab[kab]["lolos"] += 1
        else:
            summary_by_kab[kab]["tidak_lolos"] += 1
            if r['nama_pengganti']:
                summary_by_kab[kab]["pengganti"] += 1
                
    wb = openpyxl.Workbook()
    
    font_family = "Bookman Old Style"
    f_body = Font(name=font_family, size=12)
    f_body_bold = Font(name=font_family, size=12, bold=True)
    f_title = Font(name=font_family, size=14, bold=True)
    f_subtitle = Font(name=font_family, size=12, italic=True)
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    fill_header_blue = PatternFill(start_color="8EA9DB", end_color="8EA9DB", fill_type="solid")
    fill_header_red = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
    fill_header_green = PatternFill(start_color="375623", end_color="375623", fill_type="solid")
    
    thin_side = Side(border_style="thin", color="000000")
    double_bottom_side = Side(border_style="double", color="000000")
    
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom_side)
    
    # --- Sheet 1: Lamp.IA (Summary of Kabupaten) ---
    ws_ia = wb.active
    ws_ia.title = "Lamp.IA"
    ws_ia.views.sheetView[0].showGridLines = True
    ws_ia.page_setup.orientation = ws_ia.ORIENTATION_PORTRAIT
    ws_ia.page_margins.top = ws_ia.page_margins.bottom = ws_ia.page_margins.left = ws_ia.page_margins.right = 0.5
    ws_ia.sheet_properties.pageSetUpPr.fitToPage = True
    ws_ia.page_setup.fitToWidth = 1
    ws_ia.page_setup.fitToHeight = 0
    
    ws_ia.append([])
    ws_ia.merge_cells("A2:F2")
    ws_ia.merge_cells("A3:F3")
    ws_ia.merge_cells("A4:F4")
    ws_ia.cell(row=2, column=1, value="HASIL VERIFIKASI CALON PENERIMA BANTUAN").font = f_title
    ws_ia.cell(row=2, column=1).alignment = align_center
    ws_ia.cell(row=3, column=1, value=f"KEGIATAN BANTUAN STIMULAN PERUMAHAN SWADAYA {stage_name.upper()} TAHUN 2026").font = f_title
    ws_ia.cell(row=3, column=1).alignment = align_center
    ws_ia.cell(row=4, column=1, value="PROVINSI SULAWESI SELATAN").font = f_title
    ws_ia.cell(row=4, column=1).alignment = align_center
    
    headers_ia_row1 = ["No.", "KABUPATEN/KOTA", "CPB (unit)", "Hasil Verifikasi", "", ""]
    headers_ia_row2 = ["", "", "", "Lolos (unit)", "Tidak Lolos (Unit)", "CPB Pengganti (unit)"]
    
    ws_ia.append([])
    ws_ia.append(headers_ia_row1)
    ws_ia.append(headers_ia_row2)
    
    ws_ia.merge_cells("A6:A7")
    ws_ia.merge_cells("B6:B7")
    ws_ia.merge_cells("C6:C7")
    ws_ia.merge_cells("D6:F6")
    
    for r in [6, 7]:
        for c in range(1, 7):
            cell = ws_ia.cell(row=r, column=c)
            cell.font = f_body_bold
            cell.alignment = align_center
            cell.border = border_data
            cell.fill = fill_header_blue
            if r == 6 and c >= 4:
                cell.font = Font(name=font_family, size=12, bold=True, color="000000")
                cell.fill = fill_header_blue if c == 4 else fill_header_red
                
    row_num = 8
    tot_cpb = tot_lolos = tot_tl = tot_peng = 0
    for idx, (kab, val) in enumerate(summary_by_kab.items()):
        ws_ia.append([idx+1, kab, val["cpb"], val["lolos"], val["tidak_lolos"], val["pengganti"]])
        tot_cpb += val["cpb"]
        tot_lolos += val["lolos"]
        tot_tl += val["tidak_lolos"]
        tot_peng += val["pengganti"]
        
        for c in range(1, 7):
            cell = ws_ia.cell(row=row_num, column=c)
            cell.font = f_body
            cell.border = border_data
            cell.alignment = align_center if c != 2 else align_left
        row_num += 1
        
    ws_ia.append(["", "TOTAL", tot_cpb, tot_lolos, tot_tl, tot_peng])
    ws_ia.merge_cells(start_row=row_num, start_column=1, end_row=row_num, end_column=2)
    for c in range(1, 7):
        cell = ws_ia.cell(row=row_num, column=c)
        cell.font = f_body_bold
        cell.border = border_total
        cell.alignment = align_center if c != 2 else align_left
        
    ws_ia.cell(row=row_num+2, column=5, value="PARAF").font = f_body_bold
    ws_ia.cell(row=row_num+2, column=5).alignment = align_center
    
    # Border box di bawah PARAF untuk kolom tanda tangan
    paraf_box_start_row = row_num + 3
    paraf_box_end_row = row_num + 5
    for r_box in range(paraf_box_start_row, paraf_box_end_row + 1):
        for c_box in [5, 6]:
            cell = ws_ia.cell(row=r_box, column=c_box)
            cell.border = border_data
    
    ws_ia.column_dimensions['A'].width = 8
    ws_ia.column_dimensions['B'].width = 30
    ws_ia.column_dimensions['C'].width = 11.5
    ws_ia.column_dimensions['D'].width = 15
    ws_ia.column_dimensions['E'].width = 18
    ws_ia.column_dimensions['F'].width = 22

    # --- Sheet 2: Lamp.IIA (Lolos) ---
    ws_iia = wb.create_sheet(title="Lamp.IIA")
    ws_iia.views.sheetView[0].showGridLines = True
    ws_iia.page_setup.orientation = ws_iia.ORIENTATION_LANDSCAPE
    ws_iia.page_margins.top = ws_iia.page_margins.bottom = ws_iia.page_margins.left = ws_iia.page_margins.right = 0.5
    ws_iia.page_margins.header = ws_iia.page_margins.footer = 0.25
    ws_iia.sheet_properties.pageSetUpPr.fitToPage = True
    ws_iia.page_setup.fitToWidth = 1
    ws_iia.page_setup.fitToHeight = 0
    
    ws_iia.append([])
    ws_iia.merge_cells("A2:P2")
    ws_iia.merge_cells("A3:P3")
    ws_iia.merge_cells("A4:P4")
    ws_iia.cell(row=2, column=1, value="DAFTAR CALON PENERIMA BANTUAN KEGIATAN BANTUAN STIMULAN").font = f_title
    ws_iia.cell(row=2, column=1).alignment = align_center
    ws_iia.cell(row=3, column=1, value=f"KEGIATAN BANTUAN STIMULAN PERUMAHAN SWADAYA {stage_name.upper()} TAHUN 2026").font = f_title
    ws_iia.cell(row=3, column=1).alignment = align_center
    ws_iia.cell(row=4, column=1, value="PROVINSI SULAWESI SELATAN").font = f_title
    ws_iia.cell(row=4, column=1).alignment = align_center
    
    headers_iia = [
        "NO. URUT", "KODE DESA / KEL", "NAMA", "JENIS KELAMIN (L/P)", "NO.KTP", "NO.KK",
        "ALAMAT TEMPAT TINGGAL", "DESA / KELURAHAN", "KECAMATAN", "KABUPATEN / KOTA",
        "*) LOLOS / PENGGANTI", "LATITUDE", "LONGITUDE", "TAHAP", "TANGGAL", "KETERANGAN"
    ]
    ws_iia.append([])
    ws_iia.append(headers_iia)
    
    for c in range(1, len(headers_iia)+1):
        cell = ws_iia.cell(row=6, column=c)
        cell.font = f_body_bold
        cell.alignment = align_center
        cell.border = border_data
        cell.fill = fill_header_blue
        
    row_num = 7
    for idx, r in enumerate(lolos_records):
        ws_iia.append([
            idx+1, r.get('kode_desa') or "", r.get('nama') or "", r.get('jenis_kelamin') or "",
            f"'{r.get('no_ktp')}" if r.get('no_ktp') else "", f"'{r.get('no_kk')}" if r.get('no_kk') else "",
            r.get('alamat') or "", r.get('desa_kelurahan') or "", r.get('kecamatan') or "",
            r.get('kabupaten_kota') or "", "LOLOS", r.get('latitude') or "", r.get('longitude') or "",
            r.get('tahap') or "", r.get('tanggal') or "", r.get('keterangan') or ""
        ])
        for c in range(1, len(headers_iia)+1):
            cell = ws_iia.cell(row=row_num, column=c)
            cell.font = f_body
            cell.border = border_data
            if c in [1, 2, 4, 5, 6, 11, 12, 13, 14, 15]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        ws_iia.row_dimensions[row_num].height = 28
        row_num += 1
        
    for col in ws_iia.columns:
        col_letter = get_column_letter(col[0].column)
        ws_iia.column_dimensions[col_letter].width = 18
    # Column widths sesuai permintaan user
    ws_iia.column_dimensions['A'].width = 10      # NO. URUT
    ws_iia.column_dimensions['B'].width = 18       # KODE DESA
    ws_iia.column_dimensions['C'].width = 36.5     # NAMA
    ws_iia.column_dimensions['D'].width = 18       # JENIS KELAMIN
    ws_iia.column_dimensions['E'].width = 27       # NO.KTP (NIK)
    ws_iia.column_dimensions['F'].width = 27       # NO.KK
    ws_iia.column_dimensions['G'].width = 30       # ALAMAT TEMPAT TINGGAL
    ws_iia.column_dimensions['H'].width = 27       # DESA / KELURAHAN
    ws_iia.column_dimensions['I'].width = 27       # KECAMATAN
    ws_iia.column_dimensions['J'].width = 42       # KABUPATEN / KOTA
    
    # --- Sheet 3: Lamp.IIIA (Tidak Lolos / Pengganti) ---
    ws_iiia = wb.create_sheet(title="Lamp.IIIA")
    ws_iiia.views.sheetView[0].showGridLines = True
    ws_iiia.page_setup.orientation = ws_iiia.ORIENTATION_LANDSCAPE
    ws_iiia.page_margins.top = ws_iiia.page_margins.bottom = ws_iiia.page_margins.left = ws_iiia.page_margins.right = 0.5
    ws_iiia.page_margins.header = ws_iiia.page_margins.footer = 0.25
    ws_iiia.sheet_properties.pageSetUpPr.fitToPage = True
    ws_iiia.page_setup.fitToWidth = 1
    ws_iiia.page_setup.fitToHeight = 0
    
    ws_iiia.append([])
    ws_iiia.merge_cells("A2:V2")
    ws_iiia.merge_cells("A3:V3")
    ws_iiia.merge_cells("A4:V4")
    ws_iiia.cell(row=2, column=1, value="DAFTAR CALON PENERIMA BANTUAN PENGGANTI KEGIATAN BANTUAN").font = f_title
    ws_iiia.cell(row=2, column=1).alignment = align_center
    ws_iiia.cell(row=3, column=1, value=f"STIMULAN PERUMAHAN SWADAYA {stage_name.upper()} TAHUN 2026").font = f_title
    ws_iiia.cell(row=3, column=1).alignment = align_center
    ws_iiia.cell(row=4, column=1, value="PROVINSI SULAWESI SELATAN").font = f_title
    ws_iiia.cell(row=4, column=1).alignment = align_center
    
    headers_iiia_top = ["NO.", "TIDAK LOLOS", "", "", "", "", "", "", "", "", "SESUDAH", "", "", "", "", "", "", "", "", "INSTRUKSI VERIFIKASI", "", "KETERANGAN"]
    headers_iiia_bottom = [
        "NO.", "NAMA", "JENIS KELAMIN (L/P)", "NO.KTP", "NO.KK", "ALAMAT TEMPAT TINGGAL",
        "DESA / KELURAHAN", "KECAMATAN", "KABUPATEN", "ALASAN TIDAK LOLOS *)",
        "BNBA", "NAMA (PENGGANTI)", "JENIS KELAMIN (L/P)", "NO.KTP (PENGGANTI)", "NO.KK (PENGGANTI)",
        "ALAMAT TEMPAT TINGGAL", "DESA / KELURAHAN", "KECAMATAN", "KABUPATEN",
        "TAHAP", "TANGGAL", "KETERANGAN"
    ]
    
    ws_iiia.append([])
    ws_iiia.append(headers_iiia_top)
    ws_iiia.append(headers_iiia_bottom)
    
    ws_iiia.merge_cells("A6:A7")
    ws_iiia.merge_cells("B6:J6")
    ws_iiia.merge_cells("K6:S6")
    ws_iiia.merge_cells("T6:U6")
    ws_iiia.merge_cells("V6:V7")
    
    for r in [6, 7]:
        for c in range(1, len(headers_iiia_bottom)+1):
            cell = ws_iiia.cell(row=r, column=c)
            cell.font = f_body_bold
            cell.border = border_data
            cell.alignment = align_center
            
            if c >= 2 and c <= 10:
                cell.fill = fill_header_red
                cell.font = Font(name=font_family, size=12, bold=True, color="FFFFFF")
            elif c >= 11 and c <= 19:
                cell.fill = fill_header_green
                cell.font = Font(name=font_family, size=12, bold=True, color="FFFFFF")
            else:
                cell.fill = fill_header_blue
                
    row_num = 8
    for idx, r in enumerate(tidak_lolos_records):
        ws_iiia.append([
            idx+1, r.get('nama') or "", r.get('jenis_kelamin') or "", f"'{r.get('no_ktp')}" if r.get('no_ktp') else "", f"'{r.get('no_kk')}" if r.get('no_kk') else "",
            r.get('alamat') or "", r.get('desa_kelurahan') or "", r.get('kecamatan') or "", r.get('kabupaten_kota') or "",
            r.get('alasan_tidak_lolos') or "", "", r.get('nama_pengganti') or "", r.get('jenis_kelamin_pengganti') or "",
            f"'{r.get('no_ktp_pengganti')}" if r.get('no_ktp_pengganti') else "", f"'{r.get('no_kk_pengganti')}" if r.get('no_kk_pengganti') else "",
            r.get('alamat_pengganti') or "", r.get('desa_kelurahan_pengganti') or "", r.get('kecamatan_pengganti') or "", r.get('kabupaten_pengganti') or "",
            r.get('tahap') or "", r.get('tanggal') or "", r.get('keterangan') or ""
        ])
        for c in range(1, len(headers_iiia_bottom)+1):
            cell = ws_iiia.cell(row=row_num, column=c)
            cell.font = f_body
            cell.border = border_data
            if c in [1, 3, 4, 5, 10, 11, 13, 14, 15, 20, 21]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        ws_iiia.row_dimensions[row_num].height = 28
        row_num += 1

    # Catatan section di bawah tabel Lamp.IIIA
    catatan_start = row_num + 1
    ws_iiia.cell(row=catatan_start, column=1, value="Catatan:").font = Font(name=font_family, size=10, bold=True)
    catatan_items = [
        "*) Alasan Tidak Lolos, diisi dengan angka (1-8) sebagai berikut:",
        "1. Belum memiliki KK sendiri;",
        "2. Tanah bersengketa;",
        "3. Rumah dalam kondisi layak;",
        "4. Memiliki rumah lebih dari 1;",
        "5. Pernah memperoleh bantuan dari APBN/APBD/CSR/anggaran lainnya;",
        "6. Penghasilan lebih dari UMP;",
        "7. Memilih untuk dibantu dengan sumber anggaran lain;",
        "8. Menghuni kurang dari 3 tahun;",
        "9. Lainnya (diisi pada kolom keterangan);"
    ]
    for i, item in enumerate(catatan_items):
        ws_iiia.cell(row=catatan_start + 1 + i, column=1, value=item).font = Font(name=font_family, size=9)
        
    for col in ws_iiia.columns:
        col_letter = get_column_letter(col[0].column)
        ws_iiia.column_dimensions[col_letter].width = 18
    # Column widths sesuai permintaan user
    ws_iiia.column_dimensions['A'].width = 8        # NO.
    ws_iiia.column_dimensions['B'].width = 25       # NAMA (Tidak Lolos)
    ws_iiia.column_dimensions['D'].width = 27       # NO.KTP (Tidak Lolos)
    ws_iiia.column_dimensions['E'].width = 27       # NO.KK (Tidak Lolos)
    ws_iiia.column_dimensions['F'].width = 30       # ALAMAT (Tidak Lolos)
    ws_iiia.column_dimensions['G'].width = 27       # DESA / KELURAHAN (Tidak Lolos)
    ws_iiia.column_dimensions['H'].width = 27       # KECAMATAN (Tidak Lolos)
    ws_iiia.column_dimensions['I'].width = 50       # KABUPATEN (Tidak Lolos)
    ws_iiia.column_dimensions['J'].width = 60       # ALASAN TIDAK LOLOS
    ws_iiia.column_dimensions['L'].width = 25       # NAMA (Pengganti)
    ws_iiia.column_dimensions['N'].width = 27       # NO.KTP (Pengganti)
    ws_iiia.column_dimensions['O'].width = 27       # NO.KK (Pengganti)
    ws_iiia.column_dimensions['P'].width = 30       # ALAMAT (Pengganti)
    ws_iiia.column_dimensions['Q'].width = 27       # DESA / KELURAHAN (Pengganti)
    ws_iiia.column_dimensions['R'].width = 27       # KECAMATAN (Pengganti)
    ws_iiia.column_dimensions['S'].width = 50       # KABUPATEN (Pengganti)
    ws_iiia.column_dimensions['V'].width = 45       # KETERANGAN
    
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    
    filename = f"{stage_name.upper()}_LAMPIRAN_BAHV_{batch_title}.xlsx"
    
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# ============================================================
# SK DIRJEN ENDPOINTS
# ============================================================

@app.post("/api/sk-dirjen/upload")
async def upload_sk_dirjen(
    stage_name: str = Form(...),
    file: UploadFile = File(...)
):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="File harus berformat .xlsx atau .xls")
    
    content = await file.read()
    wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
    ws = wb.active
    
    rows = list(ws.iter_rows(values_only=True))
    headers = None
    header_idx = None
    for i, row in enumerate(rows):
        cleaned = [str(c).strip().replace('\n', ' ').upper() if c else '' for c in row]
        if 'NAMA' in cleaned and any('KTP' in c or 'NIK' in c for c in cleaned):
            header_idx = i
            headers = [str(c).strip().replace('\n', ' ') if c else f'COL_{j}' for j, c in enumerate(row)]
            break
    
    if header_idx is None:
        raise HTTPException(status_code=400, detail="Header tidak ditemukan. Pastikan kolom NAMA dan NO. KTP ada.")
    
    data_rows = []
    for row in rows[header_idx+1:]:
        if any(c is not None for c in row):
            data_rows.append(row)
    
    if not data_rows:
        raise HTTPException(status_code=400, detail="Tidak ada data ditemukan di file")
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id FROM sk_dirjen_batches WHERE stage_name = ?", (stage_name,))
    existing = cursor.fetchone()
    if existing:
        cursor.execute("DELETE FROM sk_dirjen_matches WHERE sk_record_id IN (SELECT id FROM sk_dirjen_records WHERE batch_id = ?)", (existing['id'],))
        cursor.execute("DELETE FROM sk_dirjen_records WHERE batch_id = ?", (existing['id'],))
        cursor.execute("DELETE FROM sk_dirjen_batches WHERE id = ?", (existing['id'],))
        conn.commit()
    
    cursor.execute("INSERT INTO sk_dirjen_batches (stage_name, filename) VALUES (?, ?)", (stage_name, file.filename))
    batch_id = cursor.lastrowid
    
    header_lower = [h.lower().replace(' ', '').replace('.', '') for h in headers]
    
    def find_col(keywords):
        for i, h in enumerate(header_lower):
            if any(k in h for k in keywords):
                return i
        return None
    
    col_nama = find_col(['nama'])
    col_nik = find_col(['ktp', 'nik'])
    col_kk = find_col(['kk', 'kartu'])
    col_jk = find_col(['jenis', 'kelamin', 'jk'])
    col_alamat = find_col(['alamat', 'tempat'])
    col_desa = find_col(['kelurahan'])
    if col_desa is None:
        col_desa = find_col(['desa'])
    col_kec = find_col(['kecamatan'])
    col_kab = find_col(['kabupaten', 'kota'])
    col_kode_desa = find_col(['kode'])
    col_no = find_col(['no', 'urut'])
    col_ket = find_col(['keterangan'])
    
    inserted = 0
    for row in data_rows:
        cells = list(row)
        def safe(idx):
            if idx is None or idx >= len(cells):
                return None
            v = cells[idx]
            if v is None:
                return None
            if isinstance(v, float):
                return str(int(v)) if v == int(v) else str(v)
            return str(v).strip()
        
        nama = safe(col_nama)
        nik = safe(col_nik)
        if not nama or not nik:
            continue
        
        cursor.execute("""
            INSERT INTO sk_dirjen_records 
            (batch_id, no_urut, kode_desa, nama, jenis_kelamin, no_ktp, no_kk, alamat, desa_kelurahan, kecamatan, kabupaten_kota, keterangan)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            batch_id, safe(col_no), safe(col_kode_desa), nama, safe(col_jk),
            nik, safe(col_kk), safe(col_alamat), safe(col_desa),
            safe(col_kec), safe(col_kab), safe(col_ket)
        ))
        sk_rec_id = cursor.lastrowid
        
        cursor.execute("""
            SELECT vr.id, vr.nama, vr.no_ktp, vr.no_kk 
            FROM verified_records vr
            WHERE vr.no_ktp = ? AND vr.status = 'LOLOS'
            ORDER BY (SELECT id FROM verified_batches WHERE id = vr.batch_id) DESC NULLS LAST
            LIMIT 1
        """, (nik,))
        matched = cursor.fetchone()
        
        if matched:
            nama_match = (nama.upper().strip() == (matched['nama'] or '').upper().strip())
            kk_match = (safe(col_kk) or '') == (matched['no_kk'] or '')
            
            if nama_match and kk_match:
                match_type = 'PERFECT'
            else:
                match_type = 'NEEDS_APPROVAL'
            
            cursor.execute("""
                SELECT vb.id as batch_id, vb.name as batch_name, vb.stage_id, ist.name as stage_name
                FROM verified_batches vb
                JOIN invers_stages ist ON vb.stage_id = ist.id
                WHERE vb.id = (SELECT batch_id FROM verified_records WHERE id = ?)
            """, (matched['id'],))
            batch_info = cursor.fetchone()
            
            if not batch_info:
                cursor.execute("""
                    SELECT vr.batch_id, vb.name as batch_name, vb.stage_id, ist.name as stage_name
                    FROM verified_records vr
                    LEFT JOIN verified_batches vb ON vb.id = vr.batch_id
                    LEFT JOIN invers_stages ist ON ist.id = vb.stage_id
                    WHERE vr.id = ?
                """, (matched['id'],))
                batch_info = cursor.fetchone()
            
            cursor.execute("""
                INSERT INTO sk_dirjen_matches (sk_record_id, verified_record_id, verified_batch_id, verified_stage_id, match_type)
                VALUES (?, ?, ?, ?, ?)
            """, (sk_rec_id, matched['id'], 
                  batch_info['batch_id'] if batch_info and batch_info['batch_id'] else None, 
                  batch_info['stage_id'] if batch_info and batch_info['stage_id'] else None, 
                  match_type))
        else:
            cursor.execute("INSERT INTO sk_dirjen_matches (sk_record_id, match_type) VALUES (?, 'NO_MATCH')", (sk_rec_id,))
        
        inserted += 1
    
    conn.commit()
    conn.close()
    
    return {
        "batch_id": batch_id,
        "stage_name": stage_name,
        "inserted_records": inserted,
        "total_rows": len(data_rows)
    }

@app.get("/api/sk-dirjen/batches")
def get_sk_dirjen_batches():
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("""
        SELECT b.*, 
            (SELECT COUNT(*) FROM sk_dirjen_records WHERE batch_id = b.id) as total_records,
            (SELECT COUNT(*) FROM sk_dirjen_matches m JOIN sk_dirjen_records r ON m.sk_record_id = r.id WHERE r.batch_id = b.id AND m.match_type = 'PERFECT') as perfect_count,
            (SELECT COUNT(*) FROM sk_dirjen_matches m JOIN sk_dirjen_records r ON m.sk_record_id = r.id WHERE r.batch_id = b.id AND m.match_type = 'NEEDS_APPROVAL') as needs_approval_count,
            (SELECT COUNT(*) FROM sk_dirjen_matches m JOIN sk_dirjen_records r ON m.sk_record_id = r.id WHERE r.batch_id = b.id AND m.match_type = 'NO_MATCH') as no_match_count
        FROM sk_dirjen_batches b ORDER BY b.id DESC
    """)
    batches = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return {"batches": batches}

@app.get("/api/sk-dirjen/search-verified")
def search_verified_for_pairing(q: str = "", desa: str = ""):
    if not q or len(q.strip()) < 2:
        return {"records": []}
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    term = f"%{q.upper()}%"
    if desa and desa.strip():
        desa_term = f"%{desa.upper().strip()}%"
        cursor.execute("""
            SELECT vr.id, vr.nama, vr.no_ktp, vr.no_kk, vr.desa_kelurahan, 
                   vr.kecamatan, vr.kabupaten_kota, vr.status, vb.name as batch_name, ist.name as tahap
            FROM verified_records vr
            JOIN verified_batches vb ON vb.id = vr.batch_id
            JOIN invers_stages ist ON ist.id = vb.stage_id
            WHERE UPPER(vr.nama) LIKE ? AND UPPER(vr.desa_kelurahan) LIKE ? AND vr.status = 'LOLOS'
            ORDER BY vr.nama
            LIMIT 20
        """, (term, desa_term))
    else:
        cursor.execute("""
            SELECT vr.id, vr.nama, vr.no_ktp, vr.no_kk, vr.desa_kelurahan, 
                   vr.kecamatan, vr.kabupaten_kota, vr.status, vb.name as batch_name, ist.name as tahap
            FROM verified_records vr
            JOIN verified_batches vb ON vb.id = vr.batch_id
            JOIN invers_stages ist ON ist.id = vb.stage_id
            WHERE (UPPER(vr.nama) LIKE ? OR vr.no_ktp LIKE ? OR vr.no_kk LIKE ?
                   OR UPPER(vr.desa_kelurahan) LIKE ?) AND vr.status = 'LOLOS'
            ORDER BY vr.nama
            LIMIT 20
        """, (term, term, term, term))
    
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return {"records": records}

@app.post("/api/sk-dirjen/pair/{sk_record_id}")
def pair_sk_dirjen_record(sk_record_id: int, body: dict = Body(...)):
    verified_record_id = body.get("verified_record_id")
    if not verified_record_id:
        raise HTTPException(status_code=400, detail="verified_record_id harus diisi")
    
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id FROM sk_dirjen_records WHERE id = ?", (sk_record_id,))
    if not cursor.fetchone():
        conn.close()
        raise HTTPException(status_code=404, detail="Record SK Dirjen tidak ditemukan")
    
    cursor.execute("SELECT id, status FROM verified_records WHERE id = ?", (verified_record_id,))
    vr_check = cursor.fetchone()
    if not vr_check:
        conn.close()
        raise HTTPException(status_code=404, detail="Record Terverifikasi tidak ditemukan")
    if vr_check['status'] != 'LOLOS':
        conn.close()
        raise HTTPException(status_code=400, detail="Hanya bisa memasangkan dengan record yang LOLOS")
    
    cursor.execute("""
        SELECT vb.id as vb_id, vb.stage_id as vs_id
        FROM verified_records vr
        JOIN verified_batches vb ON vb.id = vr.batch_id
        WHERE vr.id = ?
    """, (verified_record_id,))
    vb_row = cursor.fetchone()
    
    cursor.execute("""
        INSERT INTO sk_dirjen_matches (sk_record_id, verified_record_id, verified_batch_id, verified_stage_id, match_type, override_status)
        VALUES (?, ?, ?, ?, 'MANUAL_PAIR', 'APPROVED')
        ON CONFLICT(sk_record_id) DO UPDATE SET
            verified_record_id = excluded.verified_record_id,
            verified_batch_id = excluded.verified_batch_id,
            verified_stage_id = excluded.verified_stage_id,
            match_type = 'MANUAL_PAIR',
            override_status = 'APPROVED'
    """, (sk_record_id, verified_record_id, vb_row['vb_id'] if vb_row else None, vb_row['vs_id'] if vb_row else None))
    
    conn.commit()
    conn.close()
    
    return {"success": True, "message": "Berhasil memasangkan data"}

@app.get("/api/sk-dirjen/{batch_id}/records")
def get_sk_dirjen_records(batch_id: int, kabupaten: str = None, kecamatan: str = None, desa: str = None, tahap: str = None, status: str = None, q: str = None, asal_batch: str = None):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    query = """
        SELECT r.*, m.match_type, m.override_status, m.verified_record_id, m.verified_batch_id, m.verified_stage_id,
            COALESCE(vb.name, vb2.name) as verified_batch_name, COALESCE(ist.name, ist2.name) as verified_stage_name,
            vr.nama as verified_nama, vr.no_ktp as verified_no_ktp, vr.no_kk as verified_no_kk,
            vr.alamat as verified_alamat, vr.desa_kelurahan as verified_desa_kelurahan,
            vr.kecamatan as verified_kecamatan, vr.kabupaten_kota as verified_kabupaten_kota
        FROM sk_dirjen_records r
        LEFT JOIN sk_dirjen_matches m ON m.sk_record_id = r.id
        LEFT JOIN verified_batches vb ON vb.id = m.verified_batch_id
        LEFT JOIN invers_stages ist ON ist.id = m.verified_stage_id
        LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
        LEFT JOIN verified_batches vb2 ON vb2.id = vr.batch_id
        LEFT JOIN invers_stages ist2 ON ist2.id = vb2.stage_id
        WHERE r.batch_id = ?
    """
    params = [batch_id]
    
    if q and q.strip():
        search = f"%{q.strip().upper()}%"
        query += " AND (UPPER(r.nama) LIKE ? OR r.no_ktp LIKE ? OR r.no_kk LIKE ? OR UPPER(r.desa_kelurahan) LIKE ? OR UPPER(r.kecamatan) LIKE ? OR UPPER(r.kabupaten_kota) LIKE ?)"
        params.extend([search, search, search, search, search, search])
    
    if kabupaten:
        query += " AND UPPER(r.kabupaten_kota) = ?"
        params.append(kabupaten.upper())
    if kecamatan:
        query += " AND UPPER(r.kecamatan) = ?"
        params.append(kecamatan.upper())
    if desa:
        query += " AND UPPER(r.desa_kelurahan) = ?"
        params.append(desa.upper())
    if tahap:
        query += " AND UPPER(ist.name) = ?"
        params.append(tahap.upper())
    if asal_batch:
        query += " AND UPPER(COALESCE(vb.name, vb2.name)) = ?"
        params.append(asal_batch.upper())
    if status:
        if status == 'PERFECT':
            query += " AND m.match_type = 'PERFECT'"
        elif status == 'NEEDS_APPROVAL':
            query += " AND m.match_type = 'NEEDS_APPROVAL'"
        elif status == 'NO_MATCH':
            query += " AND m.match_type = 'NO_MATCH'"
        elif status == 'APPROVED':
            query += " AND m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED'"
        elif status == 'MANUAL_PAIR':
            query += " AND m.match_type = 'MANUAL_PAIR'"
    
    query += " ORDER BY r.no_urut, r.id"
    cursor.execute(query, params)
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return {"records": records}

@app.get("/api/sk-dirjen/all-records")
def get_sk_dirjen_all_records(q: str = None, kabupaten: str = None, kecamatan: str = None, desa: str = None, status: str = None, tahap: str = None, asal_batch: str = None):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    query = """
        SELECT r.*, m.match_type, m.override_status, m.verified_record_id, m.verified_batch_id, m.verified_stage_id,
            COALESCE(vb.name, vb2.name) as verified_batch_name, COALESCE(ist.name, ist2.name) as verified_stage_name,
            vr.nama as verified_nama, vr.no_ktp as verified_no_ktp, vr.no_kk as verified_no_kk,
            vr.alamat as verified_alamat, vr.desa_kelurahan as verified_desa_kelurahan,
            vr.kecamatan as verified_kecamatan, vr.kabupaten_kota as verified_kabupaten_kota,
            sb.stage_name as batch_stage_name
        FROM sk_dirjen_records r
        JOIN sk_dirjen_batches sb ON sb.id = r.batch_id
        LEFT JOIN sk_dirjen_matches m ON m.sk_record_id = r.id
        LEFT JOIN verified_batches vb ON vb.id = m.verified_batch_id
        LEFT JOIN invers_stages ist ON ist.id = m.verified_stage_id
        LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
        LEFT JOIN verified_batches vb2 ON vb2.id = vr.batch_id
        LEFT JOIN invers_stages ist2 ON ist2.id = vb2.stage_id
        WHERE 1=1
    """
    params = []
    
    if q and q.strip():
        search = f"%{q.strip().upper()}%"
        query += " AND (UPPER(r.nama) LIKE ? OR r.no_ktp LIKE ? OR r.no_kk LIKE ? OR UPPER(r.desa_kelurahan) LIKE ? OR UPPER(r.kecamatan) LIKE ? OR UPPER(r.kabupaten_kota) LIKE ?)"
        params.extend([search, search, search, search, search, search])
    
    if kabupaten:
        query += " AND UPPER(r.kabupaten_kota) = ?"
        params.append(kabupaten.upper())
    if kecamatan:
        query += " AND UPPER(r.kecamatan) = ?"
        params.append(kecamatan.upper())
    if desa:
        query += " AND UPPER(r.desa_kelurahan) = ?"
        params.append(desa.upper())
    if status:
        if status == 'PERFECT':
            query += " AND m.match_type = 'PERFECT'"
        elif status == 'NEEDS_APPROVAL':
            query += " AND m.match_type = 'NEEDS_APPROVAL'"
        elif status == 'NO_MATCH':
            query += " AND m.match_type = 'NO_MATCH'"
        elif status == 'APPROVED':
            query += " AND m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED'"
        elif status == 'MANUAL_PAIR':
            query += " AND m.match_type = 'MANUAL_PAIR'"
    
    if tahap:
        query += " AND UPPER(COALESCE(ist.name, ist2.name)) = ?"
        params.append(tahap.upper())
    if asal_batch:
        query += " AND UPPER(COALESCE(vb.name, vb2.name)) = ?"
        params.append(asal_batch.upper())
    
    query += " ORDER BY sb.id DESC, r.no_urut, r.id"
    cursor.execute(query, params)
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return {"records": records}

@app.post("/api/sk-dirjen/record/{record_id}/approve")
def approve_sk_dirjen_record(record_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE sk_dirjen_matches SET override_status = 'APPROVED' WHERE sk_record_id = ?", (record_id,))
    conn.commit()
    conn.close()
    return {"success": True, "message": "Persetujuan berhasil disimpan"}

@app.post("/api/sk-dirjen/record/{record_id}/reject")
def reject_sk_dirjen_record(record_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    cursor.execute("UPDATE sk_dirjen_matches SET override_status = 'REJECTED' WHERE sk_record_id = ?", (record_id,))
    conn.commit()
    conn.close()
    return {"success": True, "message": "Penolakan berhasil disimpan"}

@app.delete("/api/sk-dirjen/batch/{batch_id}")
def delete_sk_dirjen_batch(batch_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, stage_name FROM sk_dirjen_batches WHERE id = ?", (batch_id,))
    batch = cursor.fetchone()
    if not batch:
        conn.close()
        raise HTTPException(status_code=404, detail="Batch SK Dirjen tidak ditemukan")
    
    cursor.execute("DELETE FROM sk_dirjen_matches WHERE sk_record_id IN (SELECT id FROM sk_dirjen_records WHERE batch_id = ?)", (batch_id,))
    cursor.execute("DELETE FROM sk_dirjen_records WHERE batch_id = ?", (batch_id,))
    cursor.execute("DELETE FROM sk_dirjen_batches WHERE id = ?", (batch_id,))
    conn.commit()
    conn.close()
    
    return {"success": True, "message": f"Batch '{batch['stage_name']}' dan semua data terkait berhasil dihapus"}

@app.get("/api/sk-dirjen/rekap-per-tahap")
def get_sk_dirjen_rekap_per_tahap():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, stage_name FROM sk_dirjen_batches ORDER BY id")
    sk_batches = [dict(row) for row in cursor.fetchall()]
    
    cursor.execute("SELECT id, name FROM invers_stages ORDER BY id")
    invers_stages_raw = [dict(row) for row in cursor.fetchall()]
    
    def sort_stages(stages):
        import re
        murni = []
        pengganti = []
        for s in stages:
            if 'pengganti' in s['name'].lower():
                pengganti.append(s)
            else:
                murni.append(s)
        def stage_num(s):
            nums = re.findall(r'\d+', s['name'])
            return int(nums[0]) if nums else 999
        murni.sort(key=stage_num)
        pengganti.sort(key=stage_num)
        return murni + pengganti
    
    invers_stages = sort_stages(invers_stages_raw)
    
    rekap = []
    for sk_batch in sk_batches:
        row_data = {"batch_id": sk_batch['id'], "stage_name": sk_batch['stage_name'], "total": 0, "tahap": {}}
        
        for inv_stage in invers_stages:
            cursor.execute("""
                SELECT COUNT(*) as cnt FROM sk_dirjen_matches m
                JOIN sk_dirjen_records r ON m.sk_record_id = r.id
                LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
                LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
                WHERE r.batch_id = ? AND vb.stage_id = ?
                AND (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
            """, (sk_batch['id'], inv_stage['id']))
            cnt = cursor.fetchone()['cnt']
            if cnt > 0:
                row_data['tahap'][inv_stage['name']] = cnt
                row_data['total'] += cnt
        
        rekap.append(row_data)
    
    conn.close()
    return {"rekap": rekap, "invers_stages": [s['name'] for s in invers_stages]}

@app.get("/api/sk-dirjen/rekap-per-kabupaten/all")
def get_sk_dirjen_rekap_per_kabupaten_all():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT DISTINCT UPPER(kabupaten_kota) as kab FROM sk_dirjen_records ORDER BY kab")
    kabupatens = [row['kab'] for row in cursor.fetchall()]
    
    cursor.execute("""
        SELECT DISTINCT ist.name as inv_stage_name
        FROM sk_dirjen_matches m
        JOIN sk_dirjen_records r ON m.sk_record_id = r.id
        LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
        LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
        JOIN invers_stages ist ON ist.id = vb.stage_id
        WHERE (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
        ORDER BY ist.name
    """)
    inv_stages_raw = [row['inv_stage_name'] for row in cursor.fetchall()]
    
    import re
    def sort_stage_names(names):
        murni = []
        pengganti = []
        for n in names:
            if 'pengganti' in n.lower():
                pengganti.append(n)
            else:
                murni.append(n)
        def stage_num(n):
            nums = re.findall(r'\d+', n)
            return int(nums[0]) if nums else 999
        murni.sort(key=stage_num)
        pengganti.sort(key=stage_num)
        return murni + pengganti
    
    inv_stages = sort_stage_names(inv_stages_raw)
    
    cursor.execute("""
        SELECT UPPER(TRIM(COALESCE(vr.kabupaten_kota, ''))) as kab, COUNT(*) as cnt
        FROM verified_records vr
        JOIN verified_batches vb ON vb.id = vr.batch_id
        WHERE vr.status = 'LOLOS' AND vb.is_published = 1
        GROUP BY kab
    """)
    lolos_by_kab = {row['kab']: row['cnt'] for row in cursor.fetchall()}
    
    table_data = []
    for kab in kabupatens:
        kab_row = {"kabupaten": kab}
        total = 0
        
        for inv_stage in inv_stages:
            cursor.execute("""
                SELECT COUNT(*) as cnt FROM sk_dirjen_matches m
                JOIN sk_dirjen_records r ON m.sk_record_id = r.id
                LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
                LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
                JOIN invers_stages ist ON ist.id = vb.stage_id
                WHERE UPPER(r.kabupaten_kota) = ? AND ist.name = ?
                AND (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
            """, (kab, inv_stage))
            cnt = cursor.fetchone()['cnt']
            kab_row[inv_stage] = cnt
            total += cnt
        
        cpb_lolos = lolos_by_kab.get(kab, 0)
        kab_row['cpb_lolos'] = cpb_lolos
        kab_row['total'] = total
        kab_row['selisih_sk'] = cpb_lolos - total
        
        sumber_list = []
        if cpb_lolos - total > 0:
            cursor.execute("""
                SELECT vb.name as batch_name, ist.name as stage_name, COUNT(*) as cnt
                FROM verified_records vr
                JOIN verified_batches vb ON vr.batch_id = vb.id
                JOIN invers_stages ist ON vb.stage_id = ist.id
                WHERE vr.status = 'LOLOS' 
                AND vb.is_published = 1
                AND UPPER(TRIM(vr.kabupaten_kota)) = ?
                AND vr.id NOT IN (
                    SELECT m.verified_record_id FROM sk_dirjen_matches m
                    WHERE m.verified_record_id IS NOT NULL
                    AND (m.match_type = 'PERFECT' 
                        OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED')
                        OR m.match_type = 'MANUAL_PAIR')
                )
                GROUP BY vb.name, ist.name
                ORDER BY ist.name, vb.name
            """, (kab,))
            sumber_list = [dict(row) for row in cursor.fetchall()]
        kab_row['sumber_selisih'] = sumber_list
        
        table_data.append(kab_row)
    
    conn.close()
    return {
        "batch_id": None,
        "stage_name": "SEMUA BATCH",
        "kabupatens": table_data,
        "invers_stages": inv_stages
    }

@app.get("/api/sk-dirjen/sumber-selisih-detail")
def get_sumber_selisih_detail(kabupaten: str, batch_name: str, stage_name: str):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("""
        SELECT vr.nama, vr.no_ktp, vr.no_kk, vr.desa_kelurahan, 
               vr.kecamatan, vr.kabupaten_kota
        FROM verified_records vr
        JOIN verified_batches vb ON vr.batch_id = vb.id
        JOIN invers_stages ist ON vb.stage_id = ist.id
        WHERE vr.status = 'LOLOS' 
        AND vb.is_published = 1
        AND UPPER(TRIM(vr.kabupaten_kota)) = UPPER(TRIM(?))
        AND vb.name = ?
        AND ist.name = ?
        AND vr.id NOT IN (
            SELECT m.verified_record_id FROM sk_dirjen_matches m
            WHERE m.verified_record_id IS NOT NULL
            AND (m.match_type = 'PERFECT' 
                OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED')
                OR m.match_type = 'MANUAL_PAIR')
        )
        ORDER BY vr.nama
    """, (kabupaten, batch_name, stage_name))
    
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    return {"records": records}

@app.get("/api/sk-dirjen/rekap-per-kabupaten/{batch_id}")
def get_sk_dirjen_rekap_per_kabupaten(batch_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, stage_name FROM sk_dirjen_batches ORDER BY id")
    all_sk_batches = [dict(row) for row in cursor.fetchall()]
    
    cursor.execute("SELECT DISTINCT UPPER(kabupaten_kota) as kab FROM sk_dirjen_records WHERE batch_id = ?", (batch_id,))
    kabupatens = sorted([row['kab'] for row in cursor.fetchall()])
    
    current_batch = None
    for b in all_sk_batches:
        if b['id'] == batch_id:
            current_batch = b
            break
    
    cursor.execute("""
        SELECT DISTINCT ist.name as inv_stage_name
        FROM sk_dirjen_matches m
        JOIN sk_dirjen_records r ON m.sk_record_id = r.id
        LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
        LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
        JOIN invers_stages ist ON ist.id = vb.stage_id
        WHERE r.batch_id = ? AND (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
        ORDER BY ist.name
    """, (batch_id,))
    inv_stages_in_match_raw = [row['inv_stage_name'] for row in cursor.fetchall()]
    
    import re
    def sort_stage_names(names):
        murni = []
        pengganti = []
        for n in names:
            if 'pengganti' in n.lower():
                pengganti.append(n)
            else:
                murni.append(n)
        def stage_num(n):
            nums = re.findall(r'\d+', n)
            return int(nums[0]) if nums else 999
        murni.sort(key=stage_num)
        pengganti.sort(key=stage_num)
        return murni + pengganti
    
    inv_stages_in_match = sort_stage_names(inv_stages_in_match_raw)
    
    table_data = []
    for kab in kabupatens:
        kab_row = {"kabupaten": kab}
        total = 0
        
        for inv_stage in inv_stages_in_match:
            cursor.execute("""
                SELECT COUNT(*) as cnt FROM sk_dirjen_matches m
                JOIN sk_dirjen_records r ON m.sk_record_id = r.id
                LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
                LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
                JOIN invers_stages ist ON ist.id = vb.stage_id
                WHERE r.batch_id = ? AND UPPER(r.kabupaten_kota) = ? AND ist.name = ?
                AND (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
            """, (batch_id, kab, inv_stage))
            cnt = cursor.fetchone()['cnt']
            kab_row[inv_stage] = cnt
            total += cnt
        
        kab_row['total'] = total
        table_data.append(kab_row)
    
    conn.close()
    return {
        "batch_id": batch_id,
        "stage_name": current_batch['stage_name'] if current_batch else "",
        "kabupatens": table_data,
        "invers_stages": inv_stages_in_match
    }

@app.get("/api/sk-dirjen/export/{batch_id}")
def export_sk_dirjen_daftar_pb(batch_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT stage_name FROM sk_dirjen_batches WHERE id = ?", (batch_id,))
    batch = cursor.fetchone()
    if not batch:
        conn.close()
        raise HTTPException(status_code=404, detail="Batch SK Dirjen tidak ditemukan")
    stage_name = batch['stage_name']
    
    cursor.execute("""
        SELECT r.*, m.match_type, m.override_status, vb.name as verified_batch_name, ist.name as verified_stage_name
        FROM sk_dirjen_records r
        LEFT JOIN sk_dirjen_matches m ON m.sk_record_id = r.id
        LEFT JOIN verified_batches vb ON vb.id = m.verified_batch_id
        LEFT JOIN invers_stages ist ON ist.id = m.verified_stage_id
        WHERE r.batch_id = ?
        ORDER BY r.no_urut, r.id
    """, (batch_id,))
    records = [dict(row) for row in cursor.fetchall()]
    conn.close()
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DAFTAR PB SK DIRJEN"
    
    headers_excel = ["NO", "KODE DESA/KEL", "NAMA", "JENIS KELAMIN", "NO. KTP", "NO. KK", "ALAMAT TEMPAT TINGGAL", "DESA/KELURAHAN", "KECAMATAN", "KABUPATEN/KOTA", "KETERANGAN", "STATUS", "ASAL VERIFIKASI"]
    
    f_header = Font(name='Arial', size=11, bold=True)
    fill_header = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    a_header = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for col_idx, h in enumerate(headers_excel, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = a_header
        cell.border = thin_border
    
    fill_perfect = PatternFill(start_color='DCfce7', end_color='DCfce7', fill_type='solid')
    fill_needs = PatternFill(start_color='fef9c3', end_color='fef9c3', fill_type='solid')
    fill_none = PatternFill(start_color='fee2e2', end_color='fee2e2', fill_type='solid')
    f_body = Font(name='Arial', size=11)
    
    for i, r in enumerate(records, 2):
        status_label = {'PERFECT': 'Cocok', 'NEEDS_APPROVAL': 'Perlu Persetujuan', 'NO_MATCH': 'Tidak Ditemukan'}.get(r['match_type'], '')
        asal = ''
        if r['match_type'] in ('PERFECT', 'NEEDS_APPROVAL') and r['verified_batch_name']:
            asal = f"BA {r['verified_batch_name']} Tahap {r['verified_stage_name']}"
        
        vals = [r['no_urut'], r['kode_desa'], r['nama'], r['jenis_kelamin'], r['no_ktp'], r['no_kk'], r['alamat'], r['desa_kelurahan'], r['kecamatan'], r['kabupaten_kota'], r['keterangan'], status_label, asal]
        
        row_fill = fill_perfect if r['match_type'] == 'PERFECT' else (fill_needs if r['match_type'] == 'NEEDS_APPROVAL' else fill_none)
        
        for col_idx, v in enumerate(vals, 1):
            cell = ws.cell(row=i, column=col_idx, value=v)
            cell.font = f_body
            cell.border = thin_border
            if col_idx == 12:
                cell.fill = row_fill
    
    col_widths = [6, 18, 30, 16, 20, 20, 30, 22, 22, 22, 25, 18, 35]
    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    
    safe_stage = stage_name.replace(' ', '_').upper()
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={safe_stage}_DAFTAR_PB_SK_DIRJEN.xlsx"}
    )

@app.get("/api/sk-dirjen/rekap-per-tahap/export")
def export_sk_dirjen_rekap_per_tahap():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT id, stage_name FROM sk_dirjen_batches ORDER BY id")
    sk_batches = [dict(row) for row in cursor.fetchall()]
    
    cursor.execute("SELECT id, name FROM invers_stages ORDER BY id")
    invers_stages = [dict(row) for row in cursor.fetchall()]
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "REKAP PB PER TAHAP"
    
    f_header = Font(name='Arial', size=11, bold=True)
    fill_header = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    a_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    f_total = Font(name='Arial', size=11, bold=True)
    
    headers_excel = ["TAHAP SK DIRJEN"] + [s['name'] for s in invers_stages] + ["TOTAL"]
    for col_idx, h in enumerate(headers_excel, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = a_center
        cell.border = thin_border
    
    for row_idx, sk_batch in enumerate(sk_batches, 2):
        ws.cell(row=row_idx, column=1, value=sk_batch['stage_name']).font = f_header
        ws.cell(row=row_idx, column=1).border = thin_border
        total = 0
        for col_idx, inv_stage in enumerate(invers_stages, 2):
            cursor.execute("""
                SELECT COUNT(*) as cnt FROM sk_dirjen_matches m
                JOIN sk_dirjen_records r ON m.sk_record_id = r.id
                LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
                LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
                WHERE r.batch_id = ? AND vb.stage_id = ?
                AND (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
            """, (sk_batch['id'], inv_stage['id']))
            cnt = cursor.fetchone()['cnt']
            ws.cell(row=row_idx, column=col_idx, value=cnt if cnt > 0 else None).font = f_body
            ws.cell(row=row_idx, column=col_idx).border = thin_border
            ws.cell(row=row_idx, column=col_idx).alignment = a_center
            total += cnt
        ws.cell(row=row_idx, column=len(invers_stages)+2, value=total).font = f_total
        ws.cell(row=row_idx, column=len(invers_stages)+2).border = thin_border
        ws.cell(row=row_idx, column=len(invers_stages)+2).alignment = a_center
    
    total_row = len(sk_batches) + 2
    ws.cell(row=total_row, column=1, value="TOTAL").font = f_total
    ws.cell(row=total_row, column=1).border = thin_border
    for col_idx, inv_stage in enumerate(invers_stages, 2):
        cursor.execute("""
            SELECT COUNT(*) as cnt FROM sk_dirjen_matches m
            JOIN sk_dirjen_records r ON m.sk_record_id = r.id
            LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
            LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
            WHERE vb.stage_id = ?
            AND (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
        """, (inv_stage['id'],))
        cnt = cursor.fetchone()['cnt']
        ws.cell(row=total_row, column=col_idx, value=cnt if cnt > 0 else None).font = f_total
        ws.cell(row=total_row, column=col_idx).border = thin_border
        ws.cell(row=total_row, column=col_idx).alignment = a_center
    grand_total = sum(ws.cell(row=r, column=len(invers_stages)+2).value or 0 for r in range(2, total_row))
    ws.cell(row=total_row, column=len(invers_stages)+2, value=grand_total).font = f_total
    ws.cell(row=total_row, column=len(invers_stages)+2).border = thin_border
    
    ws.column_dimensions['A'].width = 25
    for i in range(2, len(invers_stages)+3):
        ws.column_dimensions[get_column_letter(i)].width = 18
    
    conn.close()
    
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=REKAP_PB_SK_DIRJEN_PER_TAHAP.xlsx"}
    )

@app.get("/api/sk-dirjen/rekap-per-kabupaten/{batch_id}/export")
def export_sk_dirjen_rekap_per_kabupaten(batch_id: int):
    conn = get_db_connection()
    cursor = conn.cursor()
    
    cursor.execute("SELECT stage_name FROM sk_dirjen_batches WHERE id = ?", (batch_id,))
    batch = cursor.fetchone()
    if not batch:
        conn.close()
        raise HTTPException(status_code=404, detail="Batch SK Dirjen tidak ditemukan")
    stage_name = batch['stage_name']
    
    cursor.execute("SELECT DISTINCT UPPER(kabupaten_kota) as kab FROM sk_dirjen_records WHERE batch_id = ?", (batch_id,))
    kabupatens = sorted([row['kab'] for row in cursor.fetchall()])
    
    cursor.execute("""
        SELECT DISTINCT ist.name as inv_stage_name
        FROM sk_dirjen_matches m
        JOIN sk_dirjen_records r ON m.sk_record_id = r.id
        LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
        LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
        JOIN invers_stages ist ON ist.id = vb.stage_id
        WHERE r.batch_id = ? AND (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
        ORDER BY ist.name
    """, (batch_id,))
    inv_stages_raw = [row['inv_stage_name'] for row in cursor.fetchall()]
    
    import re
    def sort_stage_names(names):
        murni = []
        pengganti = []
        for n in names:
            if 'pengganti' in n.lower():
                pengganti.append(n)
            else:
                murni.append(n)
        def stage_num(n):
            nums = re.findall(r'\d+', n)
            return int(nums[0]) if nums else 999
        murni.sort(key=stage_num)
        pengganti.sort(key=stage_num)
        return murni + pengganti
    
    inv_stages = sort_stage_names(inv_stages_raw)
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "REKAP PB PER KABUPATEN"
    
    f_header = Font(name='Arial', size=11, bold=True)
    fill_header = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
    a_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    f_total = Font(name='Arial', size=11, bold=True)
    f_body = Font(name='Arial', size=11)
    
    ws.merge_cells(start_row=1, start_column=1, end_row=2, end_column=1)
    ws.merge_cells(start_row=1, start_column=2, end_row=2, end_column=2)
    ws.merge_cells(start_row=1, start_column=len(inv_stages)+3, end_row=2, end_column=len(inv_stages)+3)
    
    cell_no = ws.cell(row=1, column=1, value="No.")
    cell_no.font = f_header
    cell_no.fill = fill_header
    cell_no.alignment = a_center
    cell_no.border = thin_border
    
    cell_kab = ws.cell(row=1, column=2, value="KABUPATEN/KOTA")
    cell_kab.font = f_header
    cell_kab.fill = fill_header
    cell_kab.alignment = a_center
    cell_kab.border = thin_border
    
    cell_total_header = ws.cell(row=1, column=len(inv_stages)+3, value="TOTAL")
    cell_total_header.font = f_header
    cell_total_header.fill = fill_header
    cell_total_header.alignment = a_center
    cell_total_header.border = thin_border
    
    ws.cell(row=2, column=1).border = thin_border
    ws.cell(row=2, column=2).border = thin_border
    ws.cell(row=2, column=len(inv_stages)+3).border = thin_border
    
    for col_idx, inv_name in enumerate(inv_stages, 3):
        cell = ws.cell(row=1, column=col_idx, value=f"SK DIRJEN {stage_name.upper()}")
        cell.font = f_header
        cell.fill = fill_header
        cell.alignment = a_center
        cell.border = thin_border
        
        cell2 = ws.cell(row=2, column=col_idx, value=inv_name)
        cell2.font = f_header
        cell2.fill = fill_header
        cell2.alignment = a_center
        cell2.border = thin_border
    
    for row_idx, kab in enumerate(kabupatens, 3):
        ws.cell(row=row_idx, column=1, value=row_idx-2).font = f_body
        ws.cell(row=row_idx, column=1).border = thin_border
        ws.cell(row=row_idx, column=1).alignment = a_center
        
        ws.cell(row=row_idx, column=2, value=kab).font = f_body
        ws.cell(row=row_idx, column=2).border = thin_border
        
        total = 0
        for col_idx, inv_name in enumerate(inv_stages, 3):
            cursor.execute("""
                SELECT COUNT(*) as cnt FROM sk_dirjen_matches m
                JOIN sk_dirjen_records r ON m.sk_record_id = r.id
                LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
                LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
                JOIN invers_stages ist ON ist.id = vb.stage_id
                WHERE r.batch_id = ? AND UPPER(r.kabupaten_kota) = ? AND ist.name = ?
                AND (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
            """, (batch_id, kab, inv_name))
            cnt = cursor.fetchone()['cnt']
            ws.cell(row=row_idx, column=col_idx, value=cnt if cnt > 0 else None).font = f_body
            ws.cell(row=row_idx, column=col_idx).border = thin_border
            ws.cell(row=row_idx, column=col_idx).alignment = a_center
            total += cnt
        ws.cell(row=row_idx, column=len(inv_stages)+3, value=total).font = f_total
        ws.cell(row=row_idx, column=len(inv_stages)+3).border = thin_border
        ws.cell(row=row_idx, column=len(inv_stages)+3).alignment = a_center
    
    total_row = len(kabupatens) + 3
    ws.cell(row=total_row, column=1, value="").border = thin_border
    ws.cell(row=total_row, column=2, value="TOTAL").font = f_total
    ws.cell(row=total_row, column=2).border = thin_border
    for col_idx, inv_name in enumerate(inv_stages, 3):
        cursor.execute("""
            SELECT COUNT(*) as cnt FROM sk_dirjen_matches m
            JOIN sk_dirjen_records r ON m.sk_record_id = r.id
            LEFT JOIN verified_records vr ON vr.id = m.verified_record_id
            LEFT JOIN verified_batches vb ON vb.id = COALESCE(vr.batch_id, m.verified_batch_id)
            JOIN invers_stages ist ON ist.id = vb.stage_id
            WHERE r.batch_id = ? AND ist.name = ?
            AND (m.match_type = 'PERFECT' OR (m.match_type = 'NEEDS_APPROVAL' AND m.override_status = 'APPROVED') OR m.match_type = 'MANUAL_PAIR')
        """, (batch_id, inv_name))
        cnt = cursor.fetchone()['cnt']
        ws.cell(row=total_row, column=col_idx, value=cnt if cnt > 0 else None).font = f_total
        ws.cell(row=total_row, column=col_idx).border = thin_border
        ws.cell(row=total_row, column=col_idx).alignment = a_center
    grand_total = sum(ws.cell(row=r, column=len(inv_stages)+3).value or 0 for r in range(3, total_row))
    ws.cell(row=total_row, column=len(inv_stages)+3, value=grand_total).font = f_total
    ws.cell(row=total_row, column=len(inv_stages)+3).border = thin_border
    
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 30
    for i in range(3, len(inv_stages)+4):
        ws.column_dimensions[get_column_letter(i)].width = 22
    
    conn.close()
    
    file_stream = io.BytesIO()
    wb.save(file_stream)
    file_stream.seek(0)
    
    safe_stage = stage_name.replace(' ', '_').upper()
    return StreamingResponse(
        file_stream,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={safe_stage}_REKAP_PB_PER_KABUPATEN.xlsx"}
    )
