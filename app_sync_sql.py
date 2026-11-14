# -*- coding: utf-8 -*-
"""
HỆ THỐNG ĐỒNG BỘ DỮ LIỆU TỪ GOOGLE SHEETS VÀO SQLITE DATABASE
- Giao diện trực quan (Tkinter GUI App)
- Tự động kéo dữ liệu 40 phút / lần toàn bộ 2 link:
  1. Link 1 ('data VNDL'):
     + Sheet 'Put MTI': Dán tiếp vào SQL, xóa dữ liệu ngày D (theo cột N: Putaway Time) và nạp ngày D mới, giữ nguyên các ngày cũ.
     + Sheet 'Put IB': Dán tiếp vào SQL, xóa dữ liệu ngày D (theo cột I: Time) và nạp ngày D mới, giữ nguyên các ngày cũ.
     + Sheet 'ONH', 'Cell': Xóa và dán toàn bộ dữ liệu mới.
  2. Link 2 ('VNDL File Nhập MTI'):
     + Sheet 'File Nhập hàng': Dán tiếp vào bảng 'data_mti', không xóa data cũ, tự động lọc chống trùng lặp (dữ liệu đã có thì không dán lại).
- Tích hợp cửa sổ xem trực tiếp dữ liệu SQL (tìm kiếm, lọc, phân trang).
- Tích hợp tính năng NẠP DỮ LIỆU TỪ FILE EXCEL / CSV VÀO SQL.
"""

import os
import sys
import re
import time
import datetime
import sqlite3
import threading
import queue
import logging
from typing import List, Set, Optional, Tuple, Dict
from collections import Counter
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import google.auth.transport.requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

# UTF-8 cho console Windows nếu có
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ==================== CẤU HÌNH ĐƯỜNG DẪN ====================
if getattr(sys, 'frozen', False):
    EXE_DIR = os.path.dirname(sys.executable)
    BUNDLE_DIR = getattr(sys, '_MEIPASS', EXE_DIR)
else:
    EXE_DIR = os.path.dirname(os.path.abspath(__file__))
    BUNDLE_DIR = EXE_DIR

def get_credentials_path() -> str:
    candidates = [
        os.path.join(EXE_DIR, 'credentials.json'),
        os.path.join(BUNDLE_DIR, 'credentials.json'),
        r'C:\Users\srvn0514\Desktop\mới\credentials.json',
        r'C:\Users\srvn0514\Desktop\SQL\credentials.json',
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return os.path.join(EXE_DIR, 'credentials.json')

CREDENTIALS_FILE = get_credentials_path()

# Link 1: data VNDL
SPREADSHEET_ID_1 = '1-mUiQsabDNWQ_HiwS4VQvbLpPl5xGFW4LqvM72-koiA'
TARGET_SHEETS_1 = ['Put MTI', 'ONH', 'Cell', 'Put IB', 'RTS', 'Data IB', 'Over view VNDL']

# Link 2: VNDL File Nhập MTI
SPREADSHEET_ID_2 = '1NCIRWxiwEOSp9PBmx8h8LI4JnOME5SuSE7jaQ_--BOk'
SHEET_NAME_2 = 'FIle Nhập hàng'
TABLE_NAME_2 = 'data_mti'

# Link 3: Daily truck tracker (Thông tin xe đến)
SPREADSHEET_ID_3 = '1LlMGkx2FM9lmvREUqsm35QNNjEXpj4MifbZwPZ8rMjw'
SHEET_NAME_3 = 'Daily truck tracker'
TABLE_NAME_3 = 'Daily truck tracker'

# Link 4: Xe đến VNDL
SPREADSHEET_ID_4 = '1vOUEjQGjmbYn0jEYJZZy2AdZchetBiIdUFIj47NxMAI'
SHEET_NAME_4 = 'Trang tính1'
TABLE_NAME_4 = 'xe đến VNDL'

DB_PATH = r'C:\Users\srvn0514\Desktop\SQL\Data MTI.db'
INTERVAL_SECONDS = 40 * 60  # 40 phút
LOG_FILE = os.path.join(EXE_DIR, 'sync_sheets.log')


def get_db_connection() -> sqlite3.Connection:
    db_dir = os.path.dirname(DB_PATH)
    if db_dir and not os.path.exists(db_dir):
        os.makedirs(db_dir, exist_ok=True)
    return sqlite3.connect(DB_PATH)


def init_db(conn: sqlite3.Connection):
    with conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS _sync_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                synced_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                sheet_name TEXT,
                table_name TEXT,
                row_count INTEGER,
                col_count INTEGER,
                duration_seconds REAL,
                status TEXT,
                message TEXT
            )
        """)
        # Đảm bảo bảng data_mti có cấu trúc chuẩn nếu chưa có
        conn.execute("""
            CREATE TABLE IF NOT EXISTS data_mti (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT,
                mtb TEXT,
                sku TEXT,
                qty INTEGER,
                ship_order TEXT,
                box TEXT,
                kho_chuyen TEXT,
                asn TEXT,
                pallet TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_ship_order ON data_mti (ship_order)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_mtb ON data_mti (mtb)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_sku ON data_mti (sku)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_asn ON data_mti (asn)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_date ON data_mti (date)")
        
        # Bảng Daily truck tracker (Thông tin xe đến - quản lý trực tiếp trên SQL)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS [Daily truck tracker] (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                [Ngày] TEXT,
                [Chuyến] INTEGER,
                [Kho đi] TEXT,
                [Pick thêm] TEXT,
                [Kho đến] TEXT,
                [Drop thêm] TEXT,
                [Số xe] TEXT,
                [Loại xe] TEXT,
                [Vendor] TEXT,
                [Tài xế] TEXT,
                [Thời gian lấy hàng] TEXT,
                [Note] TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                [Shipping ID] TEXT
            )
        """)
        try:
            cur_chk = conn.cursor()
            cur_chk.execute("PRAGMA table_info([Daily truck tracker])")
            _cols_chk = [c[1] for c in cur_chk.fetchall()]
            if _cols_chk and 'Shipping ID' not in _cols_chk and 'Shiping ID' not in _cols_chk:
                conn.execute("ALTER TABLE [Daily truck tracker] ADD COLUMN [Shipping ID] TEXT")
        except Exception:
            pass
        
        # Dọn dẹp các view cũ nếu còn sót
        try:
            conn.execute("DROP VIEW IF EXISTS put_mti")
            conn.execute("DROP VIEW IF EXISTS put_ib")
            conn.execute("DROP VIEW IF EXISTS v_inbound_backlog")
            conn.execute("DROP TABLE IF EXISTS KET_QUA_GOP_KHONG_TRUNG")
            conn.execute("DROP TABLE IF EXISTS bao_cao_tong_hop")
        except Exception:
            pass


def _get_compute_module():
    import importlib.util
    for p in [
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\compute_summary.py',
        r'c:\Users\srvn0514\Desktop\mới\compute_summary.py',
        r'C:\Users\srvn0514\Desktop\SQL\compute_summary.py',
        r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\compute_summary.py',
        os.path.join(EXE_DIR, 'compute_summary.py'),
    ]:
        if os.path.exists(p):
            spec = importlib.util.spec_from_file_location("compute_summary_disk", p)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    import compute_summary
    return compute_summary


def _get_compute_ib_module():
    import importlib.util
    for p in [
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\compute_summary_ib.py',
        r'c:\Users\srvn0514\Desktop\mới\compute_summary_ib.py',
        r'C:\Users\srvn0514\Desktop\SQL\compute_summary_ib.py',
        r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\compute_summary_ib.py',
        os.path.join(EXE_DIR, 'compute_summary_ib.py'),
    ]:
        if os.path.exists(p):
            spec = importlib.util.spec_from_file_location("compute_summary_ib_disk", p)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    try:
        import compute_summary_ib
        return compute_summary_ib
    except Exception:
        return None


def _get_compute_rts_module():
    import importlib.util
    for p in [
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\compute_summary_rts.py',
        r'c:\Users\srvn0514\Desktop\mới\compute_summary_rts.py',
        r'C:\Users\srvn0514\Desktop\SQL\compute_summary_rts.py',
        r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\compute_summary_rts.py',
        os.path.join(EXE_DIR, 'compute_summary_rts.py'),
    ]:
        if os.path.exists(p):
            spec = importlib.util.spec_from_file_location("compute_summary_rts_disk", p)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    try:
        import compute_summary_rts
        return compute_summary_rts
    except Exception:
        return None


def _get_compute_onh_module():
    import importlib.util
    for p in [
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\compute_summary_onh.py',
        r'c:\Users\srvn0514\Desktop\mới\compute_summary_onh.py',
        r'C:\Users\srvn0514\Desktop\SQL\compute_summary_onh.py',
        r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\compute_summary_onh.py',
        os.path.join(EXE_DIR, 'compute_summary_onh.py'),
    ]:
        if os.path.exists(p):
            spec = importlib.util.spec_from_file_location("compute_summary_onh_disk", p)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    try:
        import compute_summary_onh
        return compute_summary_onh
    except Exception:
        return None


def _get_compute_overview_module():
    import importlib.util
    for p in [
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\compute_overview_vndl.py',
        r'c:\Users\srvn0514\Desktop\mới\compute_overview_vndl.py',
        r'C:\Users\srvn0514\Desktop\SQL\compute_overview_vndl.py',
        r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\compute_overview_vndl.py',
        os.path.join(EXE_DIR, 'compute_overview_vndl.py'),
    ]:
        if os.path.exists(p):
            spec = importlib.util.spec_from_file_location("compute_overview_vndl_disk", p)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    try:
        import compute_overview_vndl
        return compute_overview_vndl
    except Exception:
        return None


def refresh_summary_table(conn: sqlite3.Connection, table_name: str = "Bao_Cao_Tong_Hop (MTI)"):
    """
    Tổng hợp dữ liệu toàn bộ từ data_mti và đối chiếu với Put MTI -> [Bao_Cao_Tong_Hop (MTI)]
    Đồng thời tổng hợp data IB đối chiếu với Put IB -> [Bao_Cao_Tong_Hop (IB)]
    Đồng thời tổng hợp data RTS -> [Bao_Cao_Tong_Hop (RTS)]
    Đồng thời tổng hợp ONH và Sku size -> [Tổng hợp ONH (VNDL)] và đẩy lên Google Sheet "ONH (SKU size)"
    """
    try:
        mod = _get_compute_module()
        mod.compute_and_save_summary(DB_PATH)
    except Exception as e:
        print(f"Lỗi refresh_summary_table MTI: {e}")

    try:
        mod_ib = _get_compute_ib_module()
        if mod_ib:
            mod_ib.compute_and_save_summary_ib(DB_PATH)
    except Exception as e_ib:
        print(f"Lỗi refresh_summary_table IB: {e_ib}")

    try:
        mod_rts = _get_compute_rts_module()
        if mod_rts:
            mod_rts.compute_and_save_summary_rts(DB_PATH)
    except Exception as e_rts:
        print(f"Lỗi refresh_summary_table RTS: {e_rts}")

    try:
        mod_onh = _get_compute_onh_module()
        if mod_onh:
            mod_onh.compute_and_save_summary_onh(DB_PATH, push_to_gsheet=True)
    except Exception as e_onh:
        print(f"Lỗi refresh_summary_table ONH: {e_onh}")

    try:
        mod_ov = _get_compute_overview_module()
        if mod_ov:
            mod_ov.compute_and_save_overview_vndl(DB_PATH)
    except Exception as e_ov:
        print(f"Lỗi refresh_summary_table Overview VNDL: {e_ov}")


def export_dashboard_html():
    """Tự động cập nhật file index.html và rts.html mỗi khi dữ liệu SQL thay đổi"""
    try:
        for dash_py in [
            r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\dashboard.py',
            r'c:\Users\srvn0514\Desktop\mới\dashboard.py',
            r'C:\Users\srvn0514\Desktop\SQL\dashboard.py',
            r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\dashboard.py',
            os.path.join(EXE_DIR, 'dashboard.py')
        ]:
            if os.path.exists(dash_py):
                import importlib.util
                spec = importlib.util.spec_from_file_location("dash_module", dash_py)
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                if hasattr(mod, 'save_all_html_copies'):
                    mod.save_all_html_copies()
                else:
                    mod.save_html_file()
                return True
    except Exception as e:
        print(f"Lỗi export_dashboard_html: {e}")
    return False


def clean_headers(raw_headers: List[str]) -> List[str]:
    seen = {}
    clean = []
    for idx, col in enumerate(raw_headers):
        c = str(col).strip()
        if not c:
            c = f"column_{idx + 1}"
        if c in seen:
            seen[c] += 1
            c = f"{c}_{seen[c]}"
        else:
            seen[c] = 0
        clean.append(c)
    return clean


def extract_dates_from_col(df: pd.DataFrame, col: str) -> set:
    if col not in df.columns:
        return set()
    dates = set()
    for val in df[col].dropna():
        s = str(val).strip()
        if not s or s.lower() in ['nan', 'none', 'null', 'nat']:
            continue
        date_str = s[:10]
        if any(sep in date_str for sep in ['-', '/']) and any(c.isdigit() for c in date_str):
            dates.add(date_str)
    return dates


def fetch_sheet_df(ws) -> pd.DataFrame:
    raw = ws.get_all_values()
    if not raw:
        return pd.DataFrame()
    headers = clean_headers(raw[0])
    data = raw[1:]
    clean_data = [r for r in data if any(str(c).strip() for c in r)]
    col_len = len(headers)
    normalized = []
    for r in clean_data:
        if len(r) < col_len:
            r = r + [''] * (col_len - len(r))
        elif len(r) > col_len:
            r = r[:col_len]
        normalized.append(r)
    return pd.DataFrame(normalized, columns=headers)


def sync_append_dedup_table(
    conn: sqlite3.Connection,
    table_name: str,
    df: pd.DataFrame,
    dedup_cols: List[str],
    emit,
    view_name: Optional[str] = None
) -> Tuple[int, int, int]:
    """
    Quy tắc:
    - CHỈ CẬP NHẬT THÊM, KHÔNG XÓA DATA CŨ.
    - Những dòng đã có trong SQL thì KHÔNG dán lại, tránh bị lặp dữ liệu.
    - Trả về (total_in_db, new_added, skipped_existing).
    """
    cursor = conn.cursor()
    cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
    table_exists = cursor.fetchone()[0] > 0

    df_cols = list(df.columns)

    # Nếu bảng chưa có trong SQL -> Khởi tạo mới từ df
    if not table_exists:
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        total_in_db = len(df)
        emit(f"    - Bảng [{table_name}] chưa có trong SQL ➔ Khởi tạo mới với {total_in_db:,} dòng.")
        if view_name:
            with conn:
                conn.execute(f'DROP VIEW IF EXISTS "{view_name}"')
                conn.execute(f'CREATE VIEW "{view_name}" AS SELECT * FROM [{table_name}]')
        return total_in_db, total_in_db, 0

    # Lấy danh sách cột thực tế của bảng SQL
    cursor.execute(f"PRAGMA table_info([{table_name}])")
    sql_cols = [c[1] for c in cursor.fetchall()]

    # Kiểm tra nếu bảng cũ có cấu trúc bị lệch hoàn toàn (ví dụ thiếu cột chính)
    matching_dedup = [c for c in dedup_cols if c in sql_cols]
    if len(matching_dedup) == 0 and len(sql_cols) > 0:
        emit(f"    - [CẢNH BÁO] Cấu trúc bảng [{table_name}] trong SQL không khớp. Đang tạo lại bảng...")
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        total_in_db = len(df)
        if view_name:
            with conn:
                conn.execute(f'DROP VIEW IF EXISTS "{view_name}"')
                conn.execute(f'CREATE VIEW "{view_name}" AS SELECT * FROM [{table_name}]')
        return total_in_db, total_in_db, 0

    # Lọc ra các cột dùng làm khóa dedup
    valid_key_cols = [c for c in dedup_cols if c in df_cols and c in sql_cols]
    if not valid_key_cols:
        valid_key_cols = [c for c in df_cols if c in sql_cols]

    def _norm_val(v):
        if v is None:
            return ''
        s = str(v).strip()
        if s in ('None', 'nan', 'NaN'):
            return ''
        if s.endswith('.0') and s[:-2].replace('-', '').isdigit():
            s = s[:-2]
        return s

    # 1. Tự động dọn dẹp loại bỏ các dòng bị lặp cũ trong SQL (nếu có)
    try:
        group_cols_sql = ", ".join([f"[{c}]" for c in valid_key_cols])
        with conn:
            conn.execute(f"""
                DELETE FROM [{table_name}]
                WHERE rowid NOT IN (
                    SELECT MIN(rowid)
                    FROM [{table_name}]
                    GROUP BY {group_cols_sql}
                )
            """)
    except Exception as e_dedup:
        pass

    # 2. Đọc tất cả các khóa hiện có trong SQL (chuẩn hóa về chuỗi)
    select_keys_sql = ", ".join([f"COALESCE([{c}], '')" for c in valid_key_cols])
    cursor.execute(f"SELECT {select_keys_sql} FROM [{table_name}]")
    existing_keys = set(tuple(_norm_val(x) for x in row) for row in cursor.fetchall())
    count_existing = len(existing_keys)

    records = df.to_dict('records')
    new_rows = []
    seen_in_batch = set()

    for r in records:
        key = tuple(_norm_val(r.get(c, '')) for c in valid_key_cols)
        if key not in existing_keys and key not in seen_in_batch:
            seen_in_batch.add(key)
            new_rows.append(r)

    if new_rows:
        df_new = pd.DataFrame(new_rows)
        common_cols = [c for c in df_new.columns if c in sql_cols]
        df_new[common_cols].to_sql(table_name, conn, if_exists='append', index=False)
        new_count = len(new_rows)
    else:
        new_count = 0

    skipped_count = len(records) - new_count
    cursor.execute(f"SELECT count(*) FROM [{table_name}]")
    total_in_db = cursor.fetchone()[0]

    emit(f"    - Sheet nguồn có: {len(records):,} dòng.")
    emit(f"    - Dữ liệu đã có trong SQL: {skipped_count:,} dòng (giữ nguyên không dán lại, tránh trùng lặp).")
    emit(f"    - Đã dán tiếp thêm: {new_count:,} dòng MỚI vào bảng [{table_name}].")

    if view_name:
        with conn:
            conn.execute(f'DROP VIEW IF EXISTS "{view_name}"')
            conn.execute(f'CREATE VIEW "{view_name}" AS SELECT * FROM [{table_name}]')

    return total_in_db, new_count, skipped_count


def sync_upsert_dedup_table(
    conn: sqlite3.Connection,
    table_name: str,
    df: pd.DataFrame,
    key_cols: List[str],
    emit
) -> Tuple[int, int, int, int]:
    """
    Quy tắc cho RTS và data IB:
    - KHÔNG XÓA DATA CŨ trong SQL.
    - Tránh dán trùng data (những dòng không đổi thì giữ nguyên, không dán lại).
    - Nếu dòng đã có trong SQL nhưng trên Sheet có thay đổi dữ liệu (Status, Update Time, Putaway Quantity...) -> CẬP NHẬT (UPDATE) dòng đó.
    - Nếu là dòng mới chưa có trong SQL -> DÁN THÊM (INSERT) vào SQL.
    - Trả về (total_in_db, new_added, updated_count, skipped_existing).
    """
    cursor = conn.cursor()
    cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
    table_exists = cursor.fetchone()[0] > 0

    df = df.fillna('')
    for c in df.columns:
        df[c] = df[c].apply(lambda x: '' if str(x).strip() in ('None', 'nan', 'NaN') else str(x).strip())

    df_cols = list(df.columns)

    if not table_exists:
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        total_in_db = len(df)
        emit(f"    - Bảng [{table_name}] chưa có trong SQL ➔ Khởi tạo mới với {total_in_db:,} dòng.")
        return total_in_db, total_in_db, 0, 0

    cursor.execute(f"PRAGMA table_info([{table_name}])")
    sql_cols = [c[1] for c in cursor.fetchall()]

    matching_keys = [c for c in key_cols if c in sql_cols]
    if len(matching_keys) == 0 and len(sql_cols) > 0:
        # Bảng rỗng hoặc bảng tạm chưa đúng cấu trúc cột chuẩn của Sheet
        emit(f"    - Khởi tạo chuẩn cấu trúc {len(df_cols)} cột cho bảng [{table_name}]...")
        df.to_sql(table_name, conn, if_exists='replace', index=False)
        total_in_db = len(df)
        emit(f"    - Đã nạp {total_in_db:,} dòng vào bảng [{table_name}].")
        return total_in_db, total_in_db, 0, 0

    # Tự động bổ sung cột mới nếu Sheet có thêm cột
    for c in df_cols:
        if c not in sql_cols:
            with conn:
                conn.execute(f"ALTER TABLE [{table_name}] ADD COLUMN [{c}] TEXT DEFAULT ''")
            sql_cols.append(c)

    common_cols = [c for c in df_cols if c in sql_cols]
    valid_key_cols = [c for c in key_cols if c in common_cols]
    if not valid_key_cols:
        valid_key_cols = common_cols

    def _norm(val) -> str:
        if val is None:
            return ''
        s = str(val).strip()
        return '' if s in ('None', 'nan', 'NaN') else s

    quoted_cols = ", ".join([f"[{c}]" for c in common_cols])
    cursor.execute(f"SELECT rowid, {quoted_cols} FROM [{table_name}] ORDER BY rowid")
    sql_rows = cursor.fetchall()

    existing_by_key = {}
    sql_key_occ = {}
    existing_by_bkey = {}
    existing_full_rows = set()

    for r in sql_rows:
        rid = r[0]
        vals = tuple(_norm(v) for v in r[1:])
        row_map = dict(zip(common_cols, vals))
        bkey = tuple(row_map[k] for k in valid_key_cols)
        occ = sql_key_occ.get(bkey, 0)
        sql_key_occ[bkey] = occ + 1
        existing_by_key[(bkey, occ)] = (rid, vals)
        if bkey not in existing_by_bkey:
            existing_by_bkey[bkey] = (rid, vals)
        existing_full_rows.add(vals)

    records = df.to_dict('records')
    df_key_occ = {}
    new_rows = []
    update_batches = []
    skipped_count = 0

    for rec in records:
        vals = tuple(_norm(rec.get(c, '')) for c in common_cols)
        row_map = dict(zip(common_cols, vals))
        bkey = tuple(row_map[k] for k in valid_key_cols)
        occ = df_key_occ.get(bkey, 0)
        df_key_occ[bkey] = occ + 1

        composite_key = (bkey, occ)
        target_match = existing_by_key.get(composite_key)
        if not target_match and sql_key_occ.get(bkey, 0) == 1:
            target_match = existing_by_bkey.get(bkey)

        if target_match:
            rid, old_vals = target_match
            if vals == old_vals:
                skipped_count += 1
            else:
                update_batches.append(vals + (rid,))
        else:
            if vals in existing_full_rows:
                skipped_count += 1
            else:
                new_rows.append({c: row_map[c] for c in common_cols})
                existing_full_rows.add(vals)

    if update_batches:
        set_clause = ", ".join([f"[{c}] = ?" for c in common_cols])
        update_sql = f"UPDATE [{table_name}] SET {set_clause} WHERE rowid = ?"
        with conn:
            conn.executemany(update_sql, update_batches)

    if new_rows:
        df_new = pd.DataFrame(new_rows)
        df_new[common_cols].to_sql(table_name, conn, if_exists='append', index=False)

    new_count = len(new_rows)
    updated_count = len(update_batches)

    cursor.execute(f"SELECT count(*) FROM [{table_name}]")
    total_in_db = cursor.fetchone()[0]

    emit(f"    - Sheet nguồn có: {len(records):,} dòng.")
    emit(f"    - Giữ nguyên không đổi (tránh dán trùng): {skipped_count:,} dòng.")
    if updated_count > 0:
        emit(f"    - ĐÃ CẬP NHẬT THAY ĐỔI: {updated_count:,} dòng có ô bị thay đổi dữ liệu trong bảng [{table_name}].")
    emit(f"    - Đã dán tiếp thêm: {new_count:,} dòng MỚI vào bảng [{table_name}].")

    return total_in_db, new_count, updated_count, skipped_count


def sync_truck_tracker_table_smart(conn: sqlite3.Connection, df_truck: pd.DataFrame, emit) -> Tuple[int, int]:
    """
    Đồng bộ dữ liệu từ Google Sheet 'Daily truck tracker' vào SQLite:
    - Lấy toàn bộ data mới từ sheet nguồn.
    - Cột [Shipping ID] do người dùng nhập trên HTML/SQL thì KHÔNG ĐỤNG VÔ, giữ nguyên 100%.
    """
    cursor = conn.cursor()
    existing_shipping_ids = {}
    try:
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS _truck_shipping_ids_overrides (
                trip_key TEXT PRIMARY KEY,
                date_val TEXT,
                chuyen TEXT,
                so_xe TEXT,
                shipping_id TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("SELECT trip_key, date_val, chuyen, so_xe, shipping_id FROM _truck_shipping_ids_overrides WHERE shipping_id IS NOT NULL AND shipping_id != ''")
        for r in cursor.fetchall():
            ok, od, och, osx, osid = r
            if ok:
                existing_shipping_ids[ok.strip().lower()] = osid
            if od and och:
                existing_shipping_ids[(od.strip().lower(), och.strip().lower())] = osid
            if od and osx:
                existing_shipping_ids[(od.strip().lower(), osx.strip().lower())] = osid
            if osx:
                existing_shipping_ids[osx.strip().lower()] = osid

        cursor.execute("PRAGMA table_info([Daily truck tracker])")
        existing_cols = [c[1] for c in cursor.fetchall()]
        col_name = None
        for cand in ['Shipping ID', 'Shiping ID', 'shipping_id']:
            if cand in existing_cols:
                col_name = cand
                break
        if col_name:
            d_col = next((c for c in ['Ngày MTO', 'Ngay MTO', 'Dự kiến giao hàng', 'Ngày', 'Ngay'] if c in existing_cols), None)
            ch_col = next((c for c in ['Số chuyến', 'So chuyen', 'Chuyến', 'Chuyen'] if c in existing_cols), None)
            sx_col = next((c for c in ['Số xe', 'So xe', 'Biển số'] if c in existing_cols), None)
            sel_cols = [f"[{c}]" for c in [d_col, ch_col, sx_col] if c]
            if sel_cols:
                cursor.execute(f"SELECT {', '.join(sel_cols)}, [{col_name}] FROM [Daily truck tracker] WHERE [{col_name}] IS NOT NULL AND [{col_name}] != ''")
                for r in cursor.fetchall():
                    sid = str(r[-1] or '').strip()
                    if sid:
                        vals = [str(v or '').strip().lower() for v in r[:-1]]
                        if len(vals) == 3:
                            existing_shipping_ids[f"{vals[0]}_{vals[1]}_{vals[2]}"] = sid
                            existing_shipping_ids[(vals[0], vals[1])] = sid
                            existing_shipping_ids[(vals[0], vals[2])] = sid
                            if vals[2]:
                                existing_shipping_ids[vals[2]] = sid
    except Exception as e_map:
        emit(f"    [Lưu ý ánh xạ Shipping ID]: {e_map}")

    # Đảm bảo df_truck có cột 'Shipping ID'
    if 'Shipping ID' not in df_truck.columns:
        if 'Shiping ID' in df_truck.columns:
            df_truck.rename(columns={'Shiping ID': 'Shipping ID'}, inplace=True)
        else:
            df_truck['Shipping ID'] = ''

    # Gán lại Shipping ID đã lưu trước đó nếu có
    restored_count = 0
    for idx, row in df_truck.iterrows():
        cur_sid = str(row.get('Shipping ID', '') or '').strip()
        if not cur_sid:
            ngay = str(row.get('Ngày MTO', row.get('Dự kiến giao hàng', row.get('Ngày', row.get('Ngay', '')))) or '').strip().lower()
            chuyen = str(row.get('Số chuyến', row.get('Chuyến', row.get('Chuyen', ''))) or '').strip().lower()
            so_xe = str(row.get('Số xe', row.get('So xe', row.get('Biển số', ''))) or '').strip().lower()

            matched_sid = None
            if f"{ngay}_{chuyen}_{so_xe}" in existing_shipping_ids:
                matched_sid = existing_shipping_ids[f"{ngay}_{chuyen}_{so_xe}"]
            elif (ngay, chuyen) in existing_shipping_ids:
                matched_sid = existing_shipping_ids[(ngay, chuyen)]
            elif (ngay, so_xe) in existing_shipping_ids:
                matched_sid = existing_shipping_ids[(ngay, so_xe)]
            elif so_xe in existing_shipping_ids:
                matched_sid = existing_shipping_ids[so_xe]

            if matched_sid:
                df_truck.at[idx, 'Shipping ID'] = matched_sid
                restored_count += 1

    # Ghi đè bảng [Daily truck tracker] với data mới từ sheet và giữ nguyên Shipping ID
    with conn:
        conn.execute("DROP TABLE IF EXISTS [Daily truck tracker]")
    df_truck.to_sql('Daily truck tracker', conn, if_exists='replace', index=False)
    cursor.execute("SELECT count(*) FROM [Daily truck tracker]")
    total_in_db = cursor.fetchone()[0]
    return total_in_db, restored_count


# ==================== LOGIC ĐỒNG BỘ TOÀN BỘ 2 LINK GOOGLE SHEETS ====================
def sync_core_logic(log_callback=None):
    def emit(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    emit("=" * 60)
    emit("BẮT ĐẦU ĐỒNG BỘ TOÀN BỘ DỮ LIỆU TỪ GOOGLE SHEETS ➔ SQL")
    emit(f"Cơ sở dữ liệu: {DB_PATH}")
    emit("=" * 60)

    creds_path = get_credentials_path()
    if not os.path.exists(creds_path):
        raise FileNotFoundError(f"Không tìm thấy file xác thực: {creds_path}")

    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = Credentials.from_service_account_file(creds_path, scopes=scope)
    client = gspread.authorize(creds)

    conn = get_db_connection()
    init_db(conn)
    cursor = conn.cursor()
    t_start = time.time()

    # -------------------------------------------------------------
    # PHẦN 1: ĐỒNG BỘ LINK 1 (Put MTI, ONH, Cell, Put IB, RTS, Data IB)
    # -------------------------------------------------------------
    emit("\n--- [PHẦN 1/2] ĐỒNG BỘ LINK 1 (DATA VNDL) ---")
    try:
        sh1 = client.open_by_key(SPREADSHEET_ID_1)
        emit(f"Đã kết nối: '{sh1.title}'")
        all_ws_1 = sh1.worksheets()

        for sheet_title in TARGET_SHEETS_1:
            t0 = time.time()
            emit(f"\n--> Đang đọc sheet: '{sheet_title}'...")
            try:
                ws = None
                for w in all_ws_1:
                    if w.title == sheet_title or w.title.strip().lower() == sheet_title.strip().lower():
                        ws = w
                        break
                if ws is None:
                    try:
                        ws = sh1.worksheet(sheet_title)
                    except Exception:
                        ws = None

                if ws is None:
                    emit(f"    [THÔNG BÁO] Sheet '{sheet_title}' chưa có trên Google Sheet, bỏ qua.")
                    continue

                df = fetch_sheet_df(ws)
                row_count = len(df)
                col_count = len(df.columns)

                if row_count == 0:
                    emit(f"    [CẢNH BÁO] Sheet '{sheet_title}' không có dữ liệu để nạp.")
                    continue

                table_name = 'data IB' if sheet_title.strip().lower() == 'data ib' else sheet_title.strip()
                if sheet_title == 'Put MTI':
                    # Put MTI: Chỉ cập nhật thêm, không xóa data cũ, không dán trùng
                    dedup_cols = ['ASN ID', 'MT Order ID', 'SKU ID', 'Location', 'Putaway Time', 'Quantity']
                    total_in_db, new_count, skipped_count = sync_append_dedup_table(
                        conn, table_name, df, dedup_cols, emit
                    )

                elif sheet_title == 'Put IB':
                    # Put IB: Chỉ cập nhật thêm, không xóa data cũ, không dán trùng
                    dedup_cols = ['Inbound ID', 'BOX ID', 'SKU', 'Location', 'Time', 'Quantity On Rack']
                    total_in_db, new_count, skipped_count = sync_append_dedup_table(
                        conn, table_name, df, dedup_cols, emit
                    )

                elif table_name == 'RTS':
                    # RTS: Chỉ cập nhật thay đổi và dán data mới, không xóa data cũ, tránh dán trùng
                    key_cols = ['Order Number', 'Request ID', 'SKU']
                    total_in_db, new_count, updated_count, skipped_count = sync_upsert_dedup_table(
                        conn, 'RTS', df, key_cols, emit
                    )

                elif table_name == 'data IB':
                    # Data IB -> data IB: Chỉ cập nhật thay đổi và dán data mới, không xóa data cũ, tránh dán trùng
                    key_cols = ['Inbound ID', 'PO ID', 'SKU']
                    total_in_db, new_count, updated_count, skipped_count = sync_upsert_dedup_table(
                        conn, 'data IB', df, key_cols, emit
                    )

                else:
                    # ONH, Cell: Xóa toàn bộ và dán data mới
                    with conn:
                        conn.execute(f"DROP TABLE IF EXISTS [{table_name}]")
                    df.to_sql(table_name, conn, if_exists='replace', index=False)
                    cursor.execute(f"SELECT count(*) FROM [{table_name}]")
                    total_in_db = cursor.fetchone()[0]
                    emit(f"    - Đã xóa toàn bộ và dán {total_in_db:,} dòng mới vào bảng [{table_name}].")

                dur = round(time.time() - t0, 2)
                emit(f"    [THÀNH CÔNG] Sheet '{ws.title}' ➔ Bảng [{table_name}] | Tổng trong SQL: {total_in_db:,} dòng ({dur}s)")

                with conn:
                    conn.execute("""
                        INSERT INTO _sync_history (sheet_name, table_name, row_count, col_count, duration_seconds, status, message)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (ws.title, table_name, row_count, col_count, dur, 'SUCCESS', f'Tổng SQL: {total_in_db}'))

            except Exception as e:
                dur = round(time.time() - t0, 2)
                emit(f"    [LỖI] Sheet '{sheet_title}': {e}")
                with conn:
                    conn.execute("""
                        INSERT INTO _sync_history (sheet_name, table_name, row_count, col_count, duration_seconds, status, message)
                        VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, (sheet_title, sheet_title, 0, 0, dur, 'ERROR', str(e)))

    except Exception as e:
        emit(f"[LỖI KẾT NỐI LINK 1]: {e}")

    # -------------------------------------------------------------
    # PHẦN 2: ĐỒNG BỘ LINK 2 (VNDL File Nhập MTI -> data_mti)
    # -------------------------------------------------------------
    emit("\n--- [PHẦN 2/2] ĐỒNG BỘ LINK 2 ('File Nhập hàng' ➔ table 'data_mti') ---")
    try:
        t0 = time.time()
        sh2 = client.open_by_key(SPREADSHEET_ID_2)
        emit(f"Đã kết nối: '{sh2.title}'")

        # Tìm sheet 'FIle Nhập hàng'
        ws2 = None
        for w in sh2.worksheets():
            if w.title.strip().lower() == 'file nhập hàng':
                ws2 = w
                break
        if not ws2:
            ws2 = sh2.worksheet(SHEET_NAME_2)

        emit(f"--> Đang đọc sheet '{ws2.title}'...")
        all_vals = ws2.get_all_values()
        data_rows = [r for r in all_vals[1:] if any(c.strip() for c in r)]

        cols_mti = ['date', 'mtb', 'sku', 'qty', 'ship_order', 'box', 'kho_chuyen', 'asn', 'pallet']
        clean_rows = []
        for r in data_rows:
            d = [str(r[i]).strip() if i < len(r) else '' for i in range(9)]
            try:
                d[3] = int(float(d[3].replace(',', '').strip())) if d[3] else 0
            except Exception:
                d[3] = 0
            clean_rows.append(d)

        df_sheet = pd.DataFrame(clean_rows, columns=cols_mti)
        # Chỉ loại bỏ những dòng giống nhau toàn bộ 100% tất cả các cột
        df_sheet = df_sheet.drop_duplicates()
        total_valid = len(df_sheet)

        with conn:
            conn.execute("DROP TABLE IF EXISTS data_mti")
            conn.execute("""
                CREATE TABLE data_mti (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT,
                    mtb TEXT,
                    sku TEXT,
                    qty INTEGER,
                    ship_order TEXT,
                    box TEXT,
                    kho_chuyen TEXT,
                    asn TEXT,
                    pallet TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            conn.executemany("""
                INSERT INTO data_mti (date, mtb, sku, qty, ship_order, box, kho_chuyen, asn, pallet)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, clean_rows)
            conn.execute("CREATE INDEX IF NOT EXISTS idx_ship_order ON data_mti (ship_order)")
            conn.execute("CREATE INDEX IF NOT EXISTS idx_asn ON data_mti (asn)")

        total_data_mti = total_valid
        dur2 = round(time.time() - t0, 2)

        emit(f"    Sheet có {len(data_rows):,} dòng.")
        emit(f"    - Đã làm sạch & đồng bộ chuẩn: {total_valid:,} dòng vào bảng [data_mti] (chỉ loại bỏ dòng trùng lặp 100%)")
        emit(f"    [THÀNH CÔNG] Tổng hiện tại trong SQL [data_mti]: {total_data_mti:,} dòng ({dur2}s)")

        with conn:
            conn.execute("""
                INSERT INTO _sync_history (sheet_name, table_name, row_count, col_count, duration_seconds, status, message)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, ('File Nhập hàng', 'data_mti', total_data_mti, 9, dur2, 'SUCCESS', f'Đồng bộ chuẩn 100%: {total_data_mti} dòng'))

        # Tự động cập nhật bảng tổng hợp đối chiếu
        try:
            emit("\n--> Đang tự động cập nhật Bảng Tổng Hợp [Bao_Cao_Tong_Hop (MTI)]...")
            refresh_summary_table(conn)
            cursor.execute("SELECT count(*) FROM [Bao_Cao_Tong_Hop (MTI)]")
            cnt_sum = cursor.fetchone()[0]
            emit(f"    [THÀNH CÔNG] Bảng tổng hợp [Bao_Cao_Tong_Hop (MTI)] đã cập nhật: {cnt_sum:,} đơn!")
        except Exception as e_sum:
            emit(f"    [LƯU Ý CẬP NHẬT TỔNG HỢP]: {e_sum}")

    except Exception as e:
        emit(f"[LỖI LINK 2]: {e}")

    # -------------------------------------------------------------
    # PHẦN 3: ĐỒNG BỘ LINK 3 (Daily truck tracker ➔ SQL)
    # -------------------------------------------------------------
    emit("\n--- [PHẦN 3/3] ĐỒNG BỘ LINK 3 (Daily truck tracker ➔ SQL) ---")
    try:
        t0 = time.time()
        sh3 = client.open_by_key(SPREADSHEET_ID_3)
        emit(f"Đã kết nối: '{sh3.title}'")
        ws3 = sh3.worksheet(SHEET_NAME_3)
        df_truck = fetch_sheet_df(ws3)
        if not df_truck.empty:
            total_trk, restored_sid = sync_truck_tracker_table_smart(conn, df_truck, emit)
            dur3 = round(time.time() - t0, 2)
            emit(f"    [THÀNH CÔNG] Đã cập nhật {total_trk:,} chuyến xe vào bảng [Daily truck tracker] ({dur3}s)")
            if restored_sid > 0:
                emit(f"    - ĐÃ BẢO TOÀN NGUYÊN VẸN {restored_sid} mã [Shipping ID] người dùng đã nhập (không bị ghi đè)!")
            with conn:
                conn.execute("""
                    INSERT INTO _sync_history (sheet_name, table_name, row_count, col_count, duration_seconds, status, message)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (SHEET_NAME_3, 'Daily truck tracker', total_trk, len(df_truck.columns), dur3, 'SUCCESS', f'Tổng: {total_trk}, Giữ {restored_sid} Shipping ID'))
        else:
            emit(f"    [CẢNH BÁO] Sheet '{SHEET_NAME_3}' trống hoặc chưa có dữ liệu.")
    except Exception as e_truck:
        emit(f"    [LƯU Ý LINK 3]: {e_truck}")
        if '403' in str(e_truck) or 'permission' in str(e_truck).lower():
            try:
                import json
                with open(creds_path, 'r', encoding='utf-8') as cf:
                    sa_email = json.load(cf).get('client_email', '')
            except Exception:
                sa_email = 'nh-p-file-mti@rosy-cache-483805-c6.iam.gserviceaccount.com'
            emit(f"    ==> Vui lòng chia sẻ quyền Người xem (Viewer) trên Google Sheet cho Service Account:")
            emit(f"        {sa_email}")
        try:
            cursor.execute("SELECT count(*) FROM [Daily truck tracker]")
            cnt_trk = cursor.fetchone()[0]
            emit(f"    [SQL] Bảng [Daily truck tracker] hiện có sẵn {cnt_trk:,} dòng đang lưu trữ trên SQLite.")
        except Exception:
            pass

    # PHẦN 4: ĐỒNG BỘ LINK 4 (xe đến VNDL ➔ SQL)
    emit("\n--- [PHẦN 4/4] ĐỒNG BỘ LINK 4 (xe đến VNDL ➔ SQL) ---")
    try:
        t0_4 = time.time()
        sh4 = client.open_by_key(SPREADSHEET_ID_4)
        ws4 = sh4.worksheet(SHEET_NAME_4)
        vals4 = ws4.get_all_values()
        if vals4 and len(vals4) > 1:
            header_idx = 0
            if len(vals4) > 2 and any('biển số' in str(c).lower() or 'dấu thời gian' in str(c).lower() for c in vals4[1]):
                header_idx = 1
            raw_headers = vals4[header_idx]
            clean_cols = []
            for i, h in enumerate(raw_headers):
                h = str(h).strip()
                prefix = ''
                if header_idx == 1:
                    if 7 <= i <= 12:
                        prefix = 'IN '
                    elif 13 <= i <= 18:
                        prefix = 'OUT '
                col_name = (prefix + h).strip() if (prefix and h) else h
                if not col_name:
                    col_name = f'Col_{i+1}'
                if i == 19 or 'shipping' in h.lower():
                    col_name = 'Shipping ID'
                clean_cols.append(col_name)

            rows4 = vals4[header_idx + 1:]
            df_xeden = pd.DataFrame(rows4)
            df_xeden = df_xeden.iloc[:, :len(clean_cols)]
            df_xeden.columns = clean_cols
            if 'Shipping ID' not in df_xeden.columns:
                df_xeden['Shipping ID'] = ''

            # Bảo toàn dữ liệu đã nhập trên Dashboard lưu trong SQLite (Pallet, Type, Warehouse, Operater, Shipping ID, Leadtime)
            try:
                c_prev = conn.cursor()
                cols_prev = [c[1] for c in c_prev.execute(f'PRAGMA table_info([{TABLE_NAME_4}])').fetchall()]
                fields_to_preserve = [
                    'IN Type', 'IN Pallet', 'IN From WH', 'IN Operater', 'IN LDD',
                    'OUT Type', 'OUT Pallet', 'OUT To WH', 'OUT Operater', 'OUT LDD',
                    'Shipping ID', 'Leadtime'
                ]
                avail_preserve = [f for f in fields_to_preserve if f in cols_prev]
                if avail_preserve and 'Biển số xe' in cols_prev:
                    cols_str = ', '.join([f'[{f}]' for f in avail_preserve])
                    prev_rows = c_prev.execute(f"SELECT [Biển số xe], [Dấu thời gian], {cols_str} FROM [{TABLE_NAME_4}]").fetchall()
                    
                    prev_map = {}
                    for pr in prev_rows:
                        raw_plate = str(pr[0] or '')
                        clean_p = re.sub(r'[^a-zA-Z0-9]', '', raw_plate).lower()
                        if clean_p:
                            prev_map[clean_p] = {avail_preserve[idx]: pr[2 + idx] for idx in range(len(avail_preserve))}
                    
                    for col_name in avail_preserve:
                        if col_name not in df_xeden.columns:
                            df_xeden[col_name] = ''
                    
                    for idx, row in df_xeden.iterrows():
                        row_plate = re.sub(r'[^a-zA-Z0-9]', '', str(row.get('Biển số xe') or '')).lower()
                        matched_dict = None
                        if row_plate in prev_map:
                            matched_dict = prev_map[row_plate]
                        else:
                            for kp, vdict in prev_map.items():
                                if (kp in row_plate or row_plate in kp) and len(kp) >= 5 and len(row_plate) >= 5:
                                    matched_dict = vdict
                                    break
                        if matched_dict:
                            for col_name, prev_val in matched_dict.items():
                                if prev_val is not None and str(prev_val).strip() != '':
                                    cur_val = str(df_xeden.at[idx, col_name] or '').strip()
                                    if not cur_val or cur_val == '0':
                                        df_xeden.at[idx, col_name] = prev_val
            except Exception as e_pres:
                emit(f"    [Lưu ý bảo toàn xe đến]: {e_pres}")

            df_xeden.to_sql(TABLE_NAME_4, conn, if_exists='replace', index=False)
            dur4 = round(time.time() - t0_4, 2)
            emit(f"    [THÀNH CÔNG] Đã cập nhật {len(df_xeden):,} dòng vào bảng [{TABLE_NAME_4}] ({dur4}s)")
        else:
            emit(f"    [THÔNG BÁO] Sheet '{SHEET_NAME_4}' chưa có dữ liệu dòng.")
    except Exception as e_xeden:
        emit(f"    [LƯU Ý LINK 4]: {e_xeden}")
        if '403' in str(e_xeden) or 'permission' in str(e_xeden).lower():
            try:
                import json
                with open(creds_path, 'r', encoding='utf-8') as cf:
                    sa_email = json.load(cf).get('client_email', '')
            except Exception:
                sa_email = 'nh-p-file-mti@rosy-cache-483805-c6.iam.gserviceaccount.com'
            emit(f"    ==> Vui lòng chia sẻ quyền Người xem / Người chỉnh sửa trên Google Sheet cho Service Account:")
            emit(f"        {sa_email}")
        try:
            cursor.execute("SELECT count(*) FROM [xe đến VNDL]")
            cnt_xd = cursor.fetchone()[0]
            emit(f"    [SQL] Bảng [xe đến VNDL] hiện có sẵn {cnt_xd:,} dòng đang lưu trữ trên SQLite.")
        except Exception:
            pass

    # -------------------------------------------------------------
    # PHẦN 5: ĐỒNG BỘ LINK 5 ('Data Plan MTI' ➔ SQL)
    # -------------------------------------------------------------
    emit("\n--- [PHẦN 5/5] ĐỒNG BỘ LINK 5 ('Data Plan MTI' ➔ SQL) ---")
    try:
        t0_5 = time.time()
        sh_plan = client.open_by_key(SPREADSHEET_ID_3)
        ws_plan = sh_plan.worksheet('Data Plan MTI')
        vals_plan = ws_plan.get_all_values()
        if vals_plan and len(vals_plan) >= 5:
            date_row = vals_plan[3]
            months_dict = {
                'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04', 'may': '05', 'jun': '06',
                'jul': '07', 'aug': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
            }
            date_col_map = {}
            for c_idx in range(15, len(date_row)):
                d_raw = date_row[c_idx].strip()
                if not d_raw:
                    continue
                parts = d_raw.split('-')
                if len(parts) == 2:
                    day_str = parts[0].strip().zfill(2)
                    mon_str = parts[1].strip().lower()[:3]
                    mm = months_dict.get(mon_str, '11')
                    date_col_map[c_idx] = f"{day_str}/{mm}/2026"

            def _parse_plan_qty(val_raw):
                s = str(val_raw or '').strip().replace(' ', '')
                if not s or s in ('-', '  -  ', 'nan', 'None'):
                    return 0
                try:
                    if ',' in s:
                        parts = s.split(',')
                        if len(parts) == 2 and len(parts[1]) != 3:
                            return int(round(float(parts[0] + '.' + parts[1]) * 1000))
                        else:
                            return int(round(float(s.replace(',', ''))))
                    elif '.' in s:
                        parts = s.split('.')
                        if len(parts) == 2 and len(parts[1]) != 3:
                            return int(round(float(s) * 1000))
                        else:
                            return int(round(float(s.replace('.', ''))))
                    else:
                        val = float(s)
                        if 0 < val < 500:
                            return int(round(val * 1000))
                        return int(round(val))
                except Exception:
                    return 0

            parsed_records = []
            in_section = False
            for r_idx, r in enumerate(vals_plan):
                col_a = r[0].strip() if len(r) > 0 else ''
                col_d = r[3].strip() if len(r) > 3 else ''
                if 'item_mti_fbs mp_vnd*' in col_a.lower():
                    in_section = True
                if in_section:
                    clean_wh = col_d.replace('*', '').strip().upper()
                    if clean_wh.startswith('VN'):
                        for c_idx, dt_str in date_col_map.items():
                            if c_idx < len(r):
                                val_raw = r[c_idx].strip()
                                qty = _parse_plan_qty(val_raw)
                                if qty > 0:
                                    parsed_records.append({
                                        'date': dt_str,
                                        'warehouse': clean_wh,
                                        'raw_label': col_a,
                                        'qty': qty
                                    })
                    elif in_section and r_idx > 45:
                        break

            df_p = pd.DataFrame(parsed_records)
            df_p.to_sql('Data Plan MTI', conn, if_exists='replace', index=False)
            dur5 = round(time.time() - t0_5, 2)
            emit(f"    [THÀNH CÔNG] Đã lưu {len(df_p):,} bản ghi vào bảng [Data Plan MTI] ({dur5}s)")
        else:
            emit("    [LƯU Ý]: Sheet 'Data Plan MTI' chưa có đủ dữ liệu.")
    except Exception as e_p5:
        emit(f"    [LƯU Ý LINK 5]: {e_p5}")

    conn.close()

    # Tự động xuất Dashboard HTML độc lập
    try:
        if export_dashboard_html():
            emit("    [DASHBOARD HTML] Đã tự động cập nhật file index.html mới nhất!")
    except Exception as e_dash:
        emit(f"    [DASHBOARD LƯU Ý]: {e_dash}")

    total_dur = round(time.time() - t_start, 2)
    emit("\n" + "=" * 60)
    emit(f"HOÀN TẤT ĐỒNG BỘ TẤT CẢ CÁC LINK TRONG {total_dur}s")
    emit("=" * 60)


# ==================== CỬA SỔ TẠO SHEET / BẢNG MỚI TRÊN SQL ====================
class CreateEmptySheetDialog(tk.Toplevel):
    """
    Cửa sổ Tạo Sheet / Bảng Mới trên SQL:
    Đơn giản tạo một sheet/bảng trống mới để người dùng chuẩn bị nạp dữ liệu Excel/CSV vào.
    """
    def __init__(self, parent, on_success_callback=None):
        super().__init__(parent)
        self.title("➕ Tạo Sheet / Bảng Mới Trên SQL")
        self.geometry("540x330")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.on_success = on_success_callback

        self.table_name_var = tk.StringVar(value="")
        self.columns_var = tk.StringVar(value="Ship Order, ASN, Receive Date, Total SKU, Quantity, Putaway, Backlog, Status, Arrived")

        self.setup_ui()

    def setup_ui(self):
        header = tk.Frame(self, bg="#059669", padx=16, pady=12)
        header.pack(fill="x", side="top")
        tk.Label(
            header,
            text="➕ TẠO SHEET / BẢNG MỚI TRÊN SQL",
            font=("Segoe UI", 12, "bold"),
            fg="white",
            bg="#059669"
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Tạo thêm sheet mới trên SQL để chuẩn bị nạp dữ liệu vào",
            font=("Segoe UI", 9),
            fg="#d1fae5",
            bg="#059669"
        ).pack(anchor="w", pady=(2, 0))

        content = ttk.Frame(self, padding=16)
        content.pack(fill="both", expand=True)

        # 1. Tên Bảng
        ttk.Label(content, text="1. Tên Sheet / Bảng mới cần tạo trên SQL:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 4))
        ent_table = ttk.Entry(content, textvariable=self.table_name_var, font=("Segoe UI", 10))
        ent_table.pack(fill="x", pady=(0, 12))
        ent_table.focus_set()

        # 2. Danh sách cột
        ttk.Label(content, text="2. Danh sách cột (phân cách bằng dấu phẩy ','):", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(0, 4))
        ent_cols = ttk.Entry(content, textvariable=self.columns_var, font=("Segoe UI", 9))
        ent_cols.pack(fill="x", pady=(0, 8))

        # Ghi chú
        ttk.Label(
            content,
            text="💡 Bảng sẽ được tạo mới trên SQL để bạn có thể bấm '📤 Nạp Excel / CSV' tải dữ liệu vào ngay sau đó.",
            font=("Segoe UI", 8),
            foreground="#047857",
            wraplength=500,
            justify="left"
        ).pack(anchor="w", pady=(4, 10))

        # Buttons
        f_btn = ttk.Frame(self, padding=12)
        f_btn.pack(fill="x", side="bottom")

        btn_create = tk.Button(
            f_btn,
            text="✅ TẠO SHEET MỚI NGAY",
            font=("Segoe UI", 10, "bold"),
            bg="#16a34a",
            fg="white",
            activebackground="#15803d",
            activeforeground="white",
            padx=16,
            pady=7,
            relief="raised",
            bd=2,
            cursor="hand2",
            command=self.execute_create
        )
        btn_create.pack(side="right")

        btn_cancel = ttk.Button(f_btn, text="Hủy", command=self.destroy)
        btn_cancel.pack(side="right", padx=(0, 10))

    def execute_create(self):
        table_name = self.table_name_var.get().strip()
        if not table_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập tên Sheet / Bảng mới!", parent=self)
            return

        if re.search(r"[;'\"]", table_name):
            messagebox.showwarning("Cảnh báo", "Tên bảng không được chứa ký tự dấu nháy hoặc chấm phẩy!", parent=self)
            return

        try:
            conn = get_db_connection()
            cursor = conn.cursor()

            cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
            if cursor.fetchone()[0] > 0:
                overwrite = messagebox.askyesno(
                    "Bảng đã tồn tại",
                    f"Bảng [{table_name}] đã tồn tại trong SQL!\n\nBạn có muốn XÓA BẢNG CŨ VÀ TẠO LẠI MỚI không?",
                    parent=self
                )
                if not overwrite:
                    conn.close()
                    return

            raw_cols = [c.strip() for c in self.columns_var.get().split(',') if c.strip()]
            if not raw_cols:
                raw_cols = ["Ship Order", "ASN", "Receive Date", "Total SKU", "Quantity", "Putaway", "Backlog", "Status", "Arrived"]
            clean_cols = clean_headers(raw_cols)
            cursor.execute(f"DROP TABLE IF EXISTS [{table_name}]")
            cols_def = ", ".join([f'[{c}] TEXT' for c in clean_cols])
            create_sql = f"CREATE TABLE [{table_name}] ({cols_def})"
            cursor.execute(create_sql)
            conn.commit()
            conn.close()

            messagebox.showinfo(
                "Đã Tạo Thành Công",
                f"Đã tạo thành công Sheet / Bảng trống [{table_name}] trên SQL!\n\n"
                f"Bảng đã được chọn sẵn trên danh sách, bây giờ bạn có thể bấm '📤 Nạp Excel' để nạp dữ liệu vào.",
                parent=self
            )

            if self.on_success:
                self.on_success(table_name)

            self.destroy()

        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tạo bảng: {e}", parent=self)


# ==================== CỬA SỔ TẠO SHEET MỚI TỪ FILE EXCEL VÀO SQL ====================
class CreateSheetFromExcelDialog(tk.Toplevel):
    """
    Cửa sổ cho phép tạo một Sheet / Bảng MỚI trên SQL trực tiếp từ file Excel/CSV:
    - Người dùng chọn file Excel và Sheet mong muốn (hoặc file CSV)
    - Nhập tên Sheet / Bảng mới muốn tạo trên SQL (tự động gợi ý từ tên Sheet trong Excel)
    - Xem trước trực tiếp danh sách cột và các dòng dữ liệu mẫu
    - Nạp toàn bộ dữ liệu vào SQLite database với cấu trúc cột & dòng y hệt file Excel
    """
    def __init__(self, parent, on_success_callback=None):
        super().__init__(parent)
        self.title("➕ Tạo Sheet / Bảng Mới Trên SQL Từ File Excel")
        self.geometry("860x700")
        self.minsize(750, 560)
        self.on_success = on_success_callback

        self.file_path_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar()
        self.table_name_var = tk.StringVar()
        self.table_status_var = tk.StringVar(value="Vui lòng chọn file Excel hoặc CSV nguồn...")
        self.preview_df: Optional[pd.DataFrame] = None
        self.existing_tables: List[str] = []

        self.load_existing_tables()
        self.setup_ui()

    def load_existing_tables(self):
        if os.path.exists(DB_PATH):
            try:
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' AND substr(name, 1, 1) != '_'")
                self.existing_tables = [r[0] for r in c.fetchall()]
                conn.close()
            except Exception:
                self.existing_tables = []

    def setup_ui(self):
        # Header banner
        header = tk.Frame(self, bg="#0f766e", padx=16, pady=10)
        header.pack(fill="x", side="top")
        tk.Label(
            header,
            text="➕ TẠO SHEET / BẢNG MỚI TRÊN SQL TỪ FILE EXCEL",
            font=("Segoe UI", 12, "bold"),
            fg="white",
            bg="#0f766e"
        ).pack(anchor="w")
        tk.Label(
            header,
            text="Tự động đọc danh sách cột và dòng từ file Excel để tạo bảng mới trên SQL y hệt file gốc",
            font=("Segoe UI", 9),
            fg="#ccfbf1",
            bg="#0f766e"
        ).pack(anchor="w")

        # 1. Chọn file Excel
        f_file = ttk.LabelFrame(self, text=" 1. Chọn File Excel / CSV Nguồn ", padding=10)
        f_file.pack(fill="x", padx=14, pady=(10, 6))

        r1 = ttk.Frame(f_file)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="Đường dẫn file:").pack(side="left", padx=(0, 6))
        ent_path = ttk.Entry(r1, textvariable=self.file_path_var, font=("Segoe UI", 9))
        ent_path.pack(side="left", fill="x", expand=True, padx=(0, 8))
        btn_browse = ttk.Button(r1, text="📁 Chọn File Excel / CSV...", command=self.browse_file)
        btn_browse.pack(side="right")

        r2 = ttk.Frame(f_file)
        r2.pack(fill="x", pady=(6, 2))
        ttk.Label(r2, text="Chọn Sheet trong file Excel:").pack(side="left", padx=(0, 6))
        self.cbo_sheet = ttk.Combobox(r2, textvariable=self.sheet_name_var, state="readonly", width=28)
        self.cbo_sheet.pack(side="left", padx=(0, 10))
        self.cbo_sheet.bind("<<ComboboxSelected>>", lambda e: self.on_sheet_changed())

        # 2. Đặt tên Sheet/Bảng mới trên SQL
        f_name = ttk.LabelFrame(self, text=" 2. Đặt Tên Sheet / Bảng Mới Trên SQL ", padding=10)
        f_name.pack(fill="x", padx=14, pady=6)

        r_name = ttk.Frame(f_name)
        r_name.pack(fill="x", pady=2)
        ttk.Label(r_name, text="Tên Sheet / Bảng trên SQL:", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 8))
        self.ent_tbl_name = ttk.Entry(r_name, textvariable=self.table_name_var, font=("Segoe UI", 10, "bold"), width=30)
        self.ent_tbl_name.pack(side="left", padx=(0, 10))
        self.table_name_var.trace_add("write", lambda *args: self.check_table_name())

        self.lbl_table_status = ttk.Label(f_name, textvariable=self.table_status_var, font=("Segoe UI", 9))
        self.lbl_table_status.pack(anchor="w", pady=(4, 0))

        # 3. Xem trước cấu trúc cột và dòng
        f_prev = ttk.LabelFrame(self, text=" 3. Xem Trước Dữ Liệu (Cột & Các Dòng Đọc Từ Excel) ", padding=8)
        f_prev.pack(fill="both", expand=True, padx=14, pady=6)

        self.lbl_summary = ttk.Label(f_prev, text="Chưa chọn file...", font=("Segoe UI", 9, "italic"), foreground="#475569")
        self.lbl_summary.pack(anchor="w", pady=(0, 6))

        tree_frame = ttk.Frame(f_prev)
        tree_frame.pack(fill="both", expand=True)

        self.tree = ttk.Treeview(tree_frame, show="headings", selectmode="browse", height=7)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        # Bottom buttons
        f_btn = ttk.Frame(self, padding=12)
        f_btn.pack(fill="x", side="bottom")

        self.btn_run_create = tk.Button(
            f_btn,
            text="🚀 BẮT ĐẦU TẠO SHEET & NẠP VÀO SQL",
            font=("Segoe UI", 11, "bold"),
            bg="#16a34a",
            fg="white",
            activebackground="#15803d",
            activeforeground="white",
            padx=20,
            pady=8,
            relief="raised",
            bd=2,
            cursor="hand2",
            command=self.execute_create
        )
        self.btn_run_create.pack(side="right")

        btn_cancel = ttk.Button(f_btn, text="Đóng", command=self.destroy)
        btn_cancel.pack(side="right", padx=(0, 10))

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            parent=self,
            title="Chọn file Excel hoặc CSV để tạo Sheet trên SQL",
            filetypes=[
                ("File Excel & CSV", "*.xlsx;*.xls;*.csv"),
                ("Excel (*.xlsx)", "*.xlsx"),
                ("Excel 97-2003 (*.xls)", "*.xls"),
                ("CSV UTF-8 (*.csv)", "*.csv"),
                ("Tất cả file", "*.*")
            ]
        )
        if not file_path:
            return

        self.file_path_var.set(file_path)
        ext = os.path.splitext(file_path)[1].lower()

        if ext in ['.xlsx', '.xls']:
            try:
                xl = pd.ExcelFile(file_path)
                sheets = xl.sheet_names
                self.cbo_sheet['values'] = sheets
                if sheets:
                    self.sheet_name_var.set(sheets[0])
                    self.cbo_sheet.config(state="readonly")
                    self.table_name_var.set(sheets[0])
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc danh sách Sheet: {e}", parent=self)
                return
        else:
            self.cbo_sheet['values'] = ["(File CSV)"]
            self.sheet_name_var.set("(File CSV)")
            self.cbo_sheet.config(state="disabled")
            base_name = os.path.splitext(os.path.basename(file_path))[0]
            self.table_name_var.set(base_name)

        self.check_table_name()
        self.load_preview()

    def on_sheet_changed(self):
        sname = self.sheet_name_var.get()
        if sname and sname != "(File CSV)":
            self.table_name_var.set(sname)
        self.check_table_name()
        self.load_preview()

    def check_table_name(self):
        tname = self.table_name_var.get().strip()
        if not tname:
            self.table_status_var.set("⚠️ Vui lòng nhập tên Sheet / Bảng muốn tạo trên SQL.")
            self.lbl_table_status.config(foreground="#dc2626")
        elif tname in self.existing_tables:
            self.table_status_var.set(f"⚠️ Bảng [{tname}] ĐÃ TỒN TẠI trong SQL. Nếu tiếp tục, dữ liệu sẽ được ghi đè mới toàn bộ (hoặc bạn có thể đổi tên khác).")
            self.lbl_table_status.config(foreground="#d97706")
        else:
            self.table_status_var.set(f"✅ Tên hợp lệ: Sẽ tạo Sheet / Bảng mới [{tname}] trên SQL.")
            self.lbl_table_status.config(foreground="#16a34a")

    def load_preview(self):
        path = self.file_path_var.get()
        if not path or not os.path.exists(path):
            return

        ext = os.path.splitext(path)[1].lower()
        sheet = self.sheet_name_var.get()

        try:
            if ext in ['.xlsx', '.xls']:
                df = pd.read_excel(path, sheet_name=sheet, nrows=50, dtype=str)
            else:
                df = pd.read_csv(path, nrows=50, dtype=str)

            df.columns = clean_headers(list(df.columns))
            df = df.dropna(how='all')
            self.preview_df = df

            cols = list(df.columns)
            self.tree["columns"] = cols
            for col in cols:
                self.tree.heading(col, text=col)
                self.tree.column(col, width=130, minwidth=80)

            for item in self.tree.get_children():
                self.tree.delete(item)

            for _, row in df.head(10).iterrows():
                vals = [str(x) if pd.notna(x) else '' for x in row.values]
                self.tree.insert("", "end", values=vals)

            total_cols = len(cols)
            self.lbl_summary.config(
                text=f"📁 File: {os.path.basename(path)} | Sheet: {sheet} | Số cột: {total_cols} cột | Đang hiển thị 10 dòng đầu tiên",
                foreground="#0f172a"
            )

        except Exception as e:
            self.lbl_summary.config(text=f"Lỗi đọc file xem trước: {e}", foreground="#dc2626")
            self.preview_df = None

    def execute_create(self):
        path = self.file_path_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file Excel hoặc CSV nguồn!", parent=self)
            return

        table_name = self.table_name_var.get().strip()
        if not table_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập tên Sheet / Bảng mới cho SQL!", parent=self)
            return

        if table_name in self.existing_tables:
            ans = messagebox.askyesno(
                "Xác nhận ghi đè",
                f"Sheet/Bảng [{table_name}] đã tồn tại trong SQL.\nBạn có muốn GHI ĐÈ toàn bộ bảng này bằng dữ liệu mới từ file Excel không?",
                parent=self
            )
            if not ans:
                return

        ext = os.path.splitext(path)[1].lower()
        sheet = self.sheet_name_var.get()

        self.btn_run_create.config(state="disabled", text="⏳ ĐANG ĐỌC & NẠP VÀO SQL...")
        self.update_idletasks()

        t0 = time.time()
        try:
            if ext in ['.xlsx', '.xls']:
                df = pd.read_excel(path, sheet_name=sheet, dtype=str)
            else:
                df = pd.read_csv(path, dtype=str)

            df.columns = clean_headers(list(df.columns))
            df = df.dropna(how='all')
            df = df.fillna('')
            for c in df.columns:
                df[c] = df[c].astype(str).str.strip()

            total_rows = len(df)
            total_cols = len(df.columns)

            if total_rows == 0:
                messagebox.showwarning("Cảnh báo", "File Excel không có dòng dữ liệu nào để nạp!", parent=self)
                self.btn_run_create.config(state="normal", text="🚀 BẮT ĐẦU TẠO SHEET & NẠP VÀO SQL")
                return

            conn = get_db_connection()
            init_db(conn)

            # Tạo bảng mới và nạp toàn bộ dòng cột từ Excel
            df.to_sql(table_name, conn, if_exists='replace', index=False)

            dur = round(time.time() - t0, 2)

            with conn:
                conn.execute("""
                    INSERT INTO _sync_history (sheet_name, table_name, row_count, col_count, duration_seconds, status, message)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (f"Excel tạo Sheet: {os.path.basename(path)}", table_name, total_rows, total_cols, dur, 'SUCCESS', f'Tạo mới sheet [{table_name}]'))

            conn.close()

            messagebox.showinfo(
                "Tạo Sheet Thành Công! 🎉",
                f"Đã tạo thành công Sheet [{table_name}] trên SQL!\n\n"
                f"- Số cột: {total_cols} cột (theo đúng chuẩn file Excel)\n"
                f"- Số dòng nạp: {total_rows:,} dòng\n"
                f"- Thời gian hoàn tất: {dur}s\n\n"
                f"Dữ liệu đã được lưu an toàn vào SQLite Database.",
                parent=self
            )

            if self.on_success:
                self.on_success(table_name)

            self.destroy()

        except Exception as e:
            messagebox.showerror("Lỗi Nạp Dữ Liệu", f"Quá trình tạo sheet thất bại: {e}", parent=self)
            self.btn_run_create.config(state="normal", text="🚀 BẮT ĐẦU TẠO SHEET & NẠP VÀO SQL")


# ==================== CỬA SỔ NẠP DỮ LIỆU TỪ FILE EXCEL / CSV ====================
class ExcelImportDialog(tk.Toplevel):
    def __init__(self, parent, on_success_callback=None, initial_table=None):
        super().__init__(parent)
        self.title("📤 Nạp Dữ Liệu Từ File Excel / CSV Vào SQL")
        self.geometry("780x630")
        self.minsize(700, 520)
        self.on_success = on_success_callback

        self.file_path_var = tk.StringVar()
        self.sheet_name_var = tk.StringVar()
        default_tbl = initial_table if (initial_table and not initial_table.startswith('_')) else "Put MTI"
        self.table_name_var = tk.StringVar(value=default_tbl)
        self.mode_var = tk.StringVar(value="deduplicate")  # deduplicate, replace, append
        self.preview_df: Optional[pd.DataFrame] = None

        self.setup_ui()
        self.load_available_tables()
        self.on_table_selected()

    def setup_ui(self):
        # 1. Chọn file
        f_file = ttk.LabelFrame(self, text=" 1. Chọn File Excel / CSV Nguồn ", padding=10)
        f_file.pack(fill="x", padx=14, pady=(10, 6))

        r1 = ttk.Frame(f_file)
        r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="Đường dẫn file:").pack(side="left", padx=(0, 6))
        ent_path = ttk.Entry(r1, textvariable=self.file_path_var, font=("Segoe UI", 9))
        ent_path.pack(side="left", fill="x", expand=True, padx=(0, 8))
        btn_browse = ttk.Button(r1, text="📁 Chọn File...", command=self.browse_file)
        btn_browse.pack(side="right")

        r2 = ttk.Frame(f_file)
        r2.pack(fill="x", pady=(6, 2))
        ttk.Label(r2, text="Chọn Sheet (nếu là Excel):").pack(side="left", padx=(0, 6))
        self.cbo_sheet = ttk.Combobox(r2, textvariable=self.sheet_name_var, state="readonly", width=25)
        self.cbo_sheet.pack(side="left", padx=(0, 10))
        self.cbo_sheet.bind("<<ComboboxSelected>>", lambda e: self.load_preview())

        # 2. Cấu hình bảng đích & Chế độ nạp
        f_cfg = ttk.LabelFrame(self, text=" 2. Chọn Sheet / Bảng Đích Trong SQL ", padding=10)
        f_cfg.pack(fill="x", padx=14, pady=6)

        ttk.Label(f_cfg, text="👉 Chọn Sheet/Bảng bạn muốn nạp vào:", font=("Segoe UI", 10, "bold"), foreground="#1e3a8a").pack(anchor="w", pady=(0, 4))
        
        target_selector_frame = ttk.Frame(f_cfg)
        target_selector_frame.pack(fill="x", pady=(0, 6))

        for tname in ['Put MTI', 'Put IB', 'RTS', 'data IB', 'ONH', 'Cell', 'data_mti']:
            rb_target = ttk.Radiobutton(
                target_selector_frame,
                text=f"[{tname}]",
                variable=self.table_name_var,
                value=tname,
                command=self.on_table_selected
            )
            rb_target.pack(side="left", padx=(0, 10))

        r_cfg1 = ttk.Frame(f_cfg)
        r_cfg1.pack(fill="x", pady=(2, 6))
        ttk.Label(r_cfg1, text="Hoặc nhập tên bảng tùy ý:").pack(side="left", padx=(0, 6))
        self.cbo_table = ttk.Combobox(r_cfg1, textvariable=self.table_name_var, width=18, font=("Segoe UI", 9))
        self.cbo_table.pack(side="left", padx=(0, 15))
        self.cbo_table.bind("<<ComboboxSelected>>", lambda e: self.on_table_selected())

        # Gợi ý quy tắc nạp
        self.lbl_rule_hint = ttk.Label(
            f_cfg,
            text="Quy tắc: Dán tiếp CHỐNG TRÙNG. Dữ liệu đã có trong SQL sẽ giữ nguyên, TUYỆT ĐỐI KHÔNG DÁN TRÙNG.",
            font=("Segoe UI", 9, "bold"),
            foreground="#16a34a"
        )
        self.lbl_rule_hint.pack(anchor="w", pady=(0, 8))

        ttk.Label(f_cfg, text="Chế độ xử lý dữ liệu:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(4, 2))
        
        self.rb_dedup = ttk.Radiobutton(
            f_cfg,
            text="🛡️ Dán tiếp CHỐNG TRÙNG (Khuyên dùng): Dữ liệu đã có trong SQL sẽ giữ nguyên, TUYỆT ĐỐI KHÔNG DÁN TRÙNG!",
            variable=self.mode_var,
            value="deduplicate"
        )
        self.rb_dedup.pack(anchor="w", pady=2)

        self.rb_replace = ttk.Radiobutton(
            f_cfg,
            text="⚠️ Ghi đè toàn bộ (Replace): Xóa sạch bảng cũ và nạp mới toàn bộ (Khuyên dùng cho ONH, Cell)",
            variable=self.mode_var,
            value="replace"
        )
        self.rb_replace.pack(anchor="w", pady=2)

        self.rb_append = ttk.Radiobutton(
            f_cfg,
            text="➕ Dán nối tiếp toàn bộ (Append): Không kiểm tra trùng (chỉ dùng khi file chắc chắn là mới 100%)",
            variable=self.mode_var,
            value="append"
        )
        self.rb_append.pack(anchor="w", pady=2)

        # 3. Xem trước thông tin
        f_prev = ttk.LabelFrame(self, text=" 3. Thông Tin Dữ Liệu File Đọc Được ", padding=8)
        f_prev.pack(fill="both", expand=True, padx=14, pady=6)

        self.txt_preview = tk.Text(
            f_prev,
            height=6,
            bg="#f8fafc",
            fg="#0f172a",
            font=("Consolas", 9),
            padx=8,
            pady=6
        )
        self.txt_preview.pack(fill="both", expand=True)

        # Bottom buttons
        f_btn = ttk.Frame(self, padding=10)
        f_btn.pack(fill="x", side="bottom")

        self.btn_run_import = tk.Button(
            f_btn,
            text="🚀 BẮT ĐẦU NẠP VÀO SQL",
            font=("Segoe UI", 10, "bold"),
            bg="#16a34a",
            fg="white",
            activebackground="#15803d",
            activeforeground="white",
            padx=18,
            pady=6,
            relief="raised",
            bd=2,
            cursor="hand2",
            command=self.execute_import
        )
        self.btn_run_import.pack(side="right")

        btn_cancel = ttk.Button(f_btn, text="Đóng", command=self.destroy)
        btn_cancel.pack(side="right", padx=(0, 10))

    def on_table_selected(self):
        tname = self.table_name_var.get().strip()
        if tname == 'Put MTI':
            self.mode_var.set('deduplicate')
            self.lbl_rule_hint.config(text="Quy tắc [Put MTI]: Dán tiếp CHỐNG TRÙNG. Dữ liệu đã có trong SQL giữ nguyên, KHÔNG dán trùng.", foreground="#16a34a")
        elif tname == 'Put IB':
            self.mode_var.set('deduplicate')
            self.lbl_rule_hint.config(text="Quy tắc [Put IB]: Dán tiếp CHỐNG TRÙNG. Dữ liệu đã có trong SQL giữ nguyên, KHÔNG dán trùng.", foreground="#16a34a")
        elif tname == 'data_mti':
            self.mode_var.set('deduplicate')
            self.lbl_rule_hint.config(text="Quy tắc [data_mti]: Dán tiếp CHỐNG TRÙNG theo (Ship Order, MTB, SKU). Tự cập nhật ngày nếu ngày bị trống.", foreground="#16a34a")
        elif tname in ['RTS', 'data IB']:
            self.mode_var.set('deduplicate')
            self.lbl_rule_hint.config(text=f"Quy tắc [{tname}]: Chỉ cập nhật thay đổi & dán data mới, KHÔNG xóa data cũ, tránh dán trùng.", foreground="#16a34a")
        elif tname in ['ONH', 'Cell']:
            self.mode_var.set('replace')
            self.lbl_rule_hint.config(text=f"Quy tắc [{tname}]: Xóa và nạp đè toàn bộ dữ liệu mới nhất.", foreground="#d97706")
        else:
            self.lbl_rule_hint.config(text=f"Bảng [{tname}]: Mặc định chống trùng lặp, giữ nguyên dữ liệu đã có trong SQL.", foreground="#2563eb")

    def load_available_tables(self):
        default_tables = ['Put MTI', 'Put IB', 'RTS', 'data IB', 'ONH', 'Cell', 'data_mti', 'Over view VNDL', 'Tổng hợp ONH (VNDL)']
        if os.path.exists(DB_PATH):
            try:
                conn = sqlite3.connect(DB_PATH)
                c = conn.cursor()
                c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name NOT LIKE 'sqlite_%' AND substr(name, 1, 1) != '_'")
                found = [r[0] for r in c.fetchall()]
                conn.close()
                for t in found:
                    if t not in default_tables:
                        default_tables.append(t)
            except Exception:
                pass
        self.cbo_table['values'] = default_tables

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            parent=self,
            title="Chọn file Excel hoặc CSV",
            filetypes=[
                ("File Excel & CSV", "*.xlsx;*.xls;*.csv"),
                ("Excel (*.xlsx)", "*.xlsx"),
                ("Excel 97-2003 (*.xls)", "*.xls"),
                ("CSV UTF-8 (*.csv)", "*.csv"),
                ("Tất cả", "*.*")
            ]
        )
        if not file_path:
            return

        self.file_path_var.set(file_path)
        ext = os.path.splitext(file_path)[1].lower()

        if ext in ['.xlsx', '.xls']:
            try:
                xl = pd.ExcelFile(file_path)
                sheets = xl.sheet_names
                self.cbo_sheet['values'] = sheets
                if sheets:
                    chosen = sheets[0]
                    for s in sheets:
                        s_lower = s.lower()
                        if 'nhập hàng' in s_lower or 'mti' in s_lower:
                            chosen = s
                            break
                        elif s in TARGET_SHEETS_1 or s_lower in [x.lower() for x in TARGET_SHEETS_1]:
                            chosen = s
                            break
                    self.sheet_name_var.set(chosen)
                    self.cbo_sheet.config(state="readonly")
            except Exception as e:
                messagebox.showerror("Lỗi", f"Không thể đọc danh sách Sheet: {e}", parent=self)
                return
        else:
            self.cbo_sheet['values'] = ["(File CSV)"]
            self.sheet_name_var.set("(File CSV)")
            self.cbo_sheet.config(state="disabled")

        # Tự động đoán bảng đích
        sheet_selected = self.sheet_name_var.get().lower()
        fn = os.path.basename(file_path).lower()
        if 'nhập hàng' in sheet_selected or 'nhập hàng' in fn:
            self.table_name_var.set('data_mti')
        elif 'put mti' in sheet_selected or 'put mti' in fn:
            self.table_name_var.set('Put MTI')
        elif 'put ib' in sheet_selected or 'put ib' in fn:
            self.table_name_var.set('Put IB')
        elif 'onh' in sheet_selected or 'onh' in fn:
            self.table_name_var.set('ONH')
        elif 'cell' in sheet_selected or 'cell' in fn:
            self.table_name_var.set('Cell')

        self.on_table_selected()
        self.load_preview()

    def load_preview(self):
        path = self.file_path_var.get()
        if not path or not os.path.exists(path):
            return

        ext = os.path.splitext(path)[1].lower()
        sheet = self.sheet_name_var.get()

        try:
            if ext in ['.xlsx', '.xls']:
                df = pd.read_excel(path, sheet_name=sheet, dtype=str)
            else:
                df = pd.read_csv(path, dtype=str)

            df.columns = clean_headers(list(df.columns))
            df = df.dropna(how='all')
            self.preview_df = df

            date_info = "Không tìm thấy cột ngày rõ ràng."
            for col in ['Putaway Time', 'Time', 'Date', 'Ngày', 'Created Time', df.columns[0]]:
                dates = extract_dates_from_col(df, col)
                if dates:
                    date_info = f"Cột ngày '{col}': Phát hiện {len(dates)} ngày -> {sorted(list(dates))[:5]}..."
                    break

            info = (
                f"Tên File: {os.path.basename(path)}\n"
                f"Sheet: {sheet}\n"
                f"Tổng số dòng dữ liệu: {len(df):,} dòng | Tổng số cột: {len(df.columns)} cột\n"
                f"Danh sách cột: {list(df.columns)[:10]}\n"
                f"Nhận diện ngày: {date_info}\n"
            )
            self.txt_preview.delete("1.0", "end")
            self.txt_preview.insert("1.0", info)

        except Exception as e:
            self.txt_preview.delete("1.0", "end")
            self.txt_preview.insert("1.0", f"Lỗi đọc file: {e}")
            self.preview_df = None

    def execute_import(self):
        if self.preview_df is None or len(self.preview_df) == 0:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn file có dữ liệu trước khi nạp!", parent=self)
            return

        table_name = self.table_name_var.get().strip()
        if not table_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập hoặc chọn bảng SQL đích!", parent=self)
            return

        mode = self.mode_var.get()
        df = self.preview_df

        confirm = messagebox.askyesno(
            "Xác nhận nạp dữ liệu",
            f"Bạn có chắc chắn muốn nạp {len(df):,} dòng từ file vào bảng [{table_name}]?\n"
            f"Chế độ: {mode.upper()}",
            parent=self
        )
        if not confirm:
            return

        t0 = time.time()
        try:
            conn = get_db_connection()
            init_db(conn)
            cursor = conn.cursor()

            cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
            table_exists = cursor.fetchone()[0] > 0
            updated_date_count = 0
            inserted_count = 0
            skipped_count = 0

            # 1. Chế độ CHỐNG TRÙNG LẶP (Deduplicate)
            if mode == 'deduplicate' or mode == 'deduplicate_mti' or (table_name == 'data_mti' and mode != 'replace'):
                if table_name == 'data_mti':
                    cursor.execute("SELECT ship_order, mtb, sku, COALESCE(date, '') FROM data_mti")
                    existing_map = {}
                    for r in cursor.fetchall():
                        k = (str(r[0]).strip(), str(r[1]).strip(), str(r[2]).strip())
                        existing_map[k] = str(r[3]).strip()

                    new_records = []
                    update_date_records = []

                    for _, row in df.iterrows():
                        r = [str(x).strip() if pd.notna(x) else '' for x in row.values]
                        date_val = r[0] if len(r) > 0 else ''
                        mtb_val = r[1] if len(r) > 1 else ''
                        sku_val = r[2] if len(r) > 2 else ''
                        qty_str = r[3] if len(r) > 3 else '0'
                        try:
                            qty_val = int(qty_str)
                        except Exception:
                            qty_val = 0
                        ship_order_val = r[4] if len(r) > 4 else ''
                        box_val = r[5] if len(r) > 5 else ''
                        kho_chuyen_val = r[6] if len(r) > 6 else ''
                        asn_val = r[7] if len(r) > 7 else ''
                        pallet_val = r[8] if len(r) > 8 else ''

                        key = (ship_order_val, mtb_val, sku_val)
                        if key not in existing_map:
                            new_records.append((date_val, mtb_val, sku_val, qty_val, ship_order_val, box_val, kho_chuyen_val, asn_val, pallet_val))
                            existing_map[key] = date_val
                        else:
                            current_db_date = existing_map[key]
                            if not current_db_date and date_val:
                                update_date_records.append((date_val, ship_order_val, mtb_val, sku_val))
                                existing_map[key] = date_val

                    if update_date_records:
                        with conn:
                            conn.executemany("""
                                UPDATE data_mti 
                                SET date = ? 
                                WHERE ship_order = ? AND mtb = ? AND sku = ? AND (date IS NULL OR trim(date) = '')
                            """, update_date_records)
                        updated_date_count = len(update_date_records)

                    if new_records:
                        with conn:
                            conn.executemany("""
                                INSERT INTO data_mti (date, mtb, sku, qty, ship_order, box, kho_chuyen, asn, pallet)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                            """, new_records)

                    inserted_count = len(new_records)
                    skipped_count = len(df) - inserted_count

                elif table_name == 'Put MTI':
                    dedup_cols = ['ASN ID', 'MT Order ID', 'SKU ID', 'Location', 'Putaway Time', 'Quantity']
                    _, inserted_count, skipped_count = sync_append_dedup_table(
                        conn, table_name, df, dedup_cols, lambda m: None, view_name='put_mti'
                    )

                elif table_name == 'Put IB':
                    dedup_cols = ['Inbound ID', 'BOX ID', 'SKU', 'Location', 'Time', 'Quantity On Rack']
                    _, inserted_count, skipped_count = sync_append_dedup_table(
                        conn, table_name, df, dedup_cols, lambda m: None, view_name='put_ib'
                    )

                elif table_name == 'RTS':
                    key_cols = ['Order Number', 'Request ID', 'SKU']
                    _, inserted_count, updated_date_count, skipped_count = sync_upsert_dedup_table(
                        conn, 'RTS', df, key_cols, lambda m: None
                    )

                elif table_name == 'data IB':
                    key_cols = ['Inbound ID', 'PO ID', 'SKU']
                    _, inserted_count, updated_date_count, skipped_count = sync_upsert_dedup_table(
                        conn, 'data IB', df, key_cols, lambda m: None
                    )

                else:
                    # Bảng tùy biến khác
                    cursor.execute(f"PRAGMA table_info([{table_name}])")
                    sql_cols = [r[1] for r in cursor.fetchall()]
                    if not sql_cols:
                        df.to_sql(table_name, conn, if_exists='replace', index=False)
                        inserted_count = len(df)
                        skipped_count = 0
                    else:
                        common_cols = [c for c in df.columns if c in sql_cols]
                        _, inserted_count, skipped_count = sync_append_dedup_table(
                            conn, table_name, df, common_cols, lambda m: None
                        )

            elif mode == 'append':
                df.to_sql(table_name, conn, if_exists='append' if table_exists else 'replace', index=False)
                inserted_count = len(df)
                skipped_count = 0

            elif mode == 'replace':
                df.to_sql(table_name, conn, if_exists='replace', index=False)
                inserted_count = len(df)
                skipped_count = 0

            cursor.execute(f"SELECT count(*) FROM [{table_name}]")
            total_after = cursor.fetchone()[0]
            dur = round(time.time() - t0, 2)

            with conn:
                conn.execute("""
                    INSERT INTO _sync_history (sheet_name, table_name, row_count, col_count, duration_seconds, status, message)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (f"File Excel: {os.path.basename(self.file_path_var.get())}", table_name, inserted_count, len(df.columns), dur, 'SUCCESS', f'Chế độ: {mode}'))

            conn.close()

            msg_res = f"ĐÃ NẠP DỮ LIỆU VÀO BẢNG [{table_name}] THÀNH CÔNG!\n\n"
            if skipped_count > 0:
                msg_res += f"🛡️ Đã có trong SQL: {skipped_count:,} dòng (GIỮ NGUYÊN, KHÔNG DÁN TRÙNG)\n"
            if updated_date_count > 0:
                msg_res += f"📅 Đã cập nhật bổ sung ngày: {updated_date_count:,} dòng\n"
            msg_res += f"➕ Số dòng mới nạp thêm: {inserted_count:,} dòng\n"
            msg_res += f"📊 Tổng cộng dữ liệu trong SQL hiện có: {total_after:,} dòng\n"
            msg_res += f"⏱️ Thời gian xử lý: {dur}s"

            messagebox.showinfo("Nạp Thành Công!", msg_res, parent=self)

            if self.on_success:
                try:
                    self.on_success(table_name)
                except TypeError:
                    self.on_success()

            self.destroy()

        except Exception as e:
            messagebox.showerror("Lỗi Nạp Dữ Liệu", f"Quá trình nạp thất bại: {e}", parent=self)


# ==================== CỬA SỔ XEM DỮ LIỆU SQL ====================
class SQLDataViewer(tk.Toplevel):
    def __init__(self, parent, initial_table=None):
        super().__init__(parent)
        self.title("Xem Dữ Liệu SQL - Data MTI")
        self.geometry("1180x680")
        self.minsize(850, 520)

        self.current_table = tk.StringVar(value=initial_table or "")
        self.search_var = tk.StringVar()
        self.limit_var = tk.StringVar(value="500")

        self.setup_ui()
        self.load_tables(prefer_table=initial_table)

    def setup_ui(self):
        # DÒNG 1: CHỌN BẢNG, TÌM KIẾM, SỐ DÒNG & TRẠNG THÁI
        row1 = ttk.Frame(self, padding=(10, 8, 10, 4))
        row1.pack(fill="x", side="top")

        ttk.Label(row1, text="Chọn Bảng:", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 6))
        self.cbo_tables = ttk.Combobox(row1, textvariable=self.current_table, state="readonly", width=22, font=("Segoe UI", 10))
        self.cbo_tables.pack(side="left", padx=(0, 14))
        self.cbo_tables.bind("<<ComboboxSelected>>", lambda e: self.load_data())

        ttk.Label(row1, text="Tìm kiếm:").pack(side="left", padx=(0, 6))
        ent_search = ttk.Entry(row1, textvariable=self.search_var, width=22, font=("Segoe UI", 10))
        ent_search.pack(side="left", padx=(0, 6))
        ent_search.bind("<Return>", lambda e: self.load_data())

        btn_search = ttk.Button(row1, text="🔍 Tìm", command=self.load_data)
        btn_search.pack(side="left", padx=(0, 4))

        btn_clear_search = ttk.Button(row1, text="✖ Xóa Tìm", command=self.clear_search)
        btn_clear_search.pack(side="left", padx=(0, 14))

        ttk.Label(row1, text="Số dòng:").pack(side="left", padx=(0, 6))
        cbo_limit = ttk.Combobox(row1, textvariable=self.limit_var, values=["100", "500", "1000", "5000", "Tất cả"], width=8, state="readonly")
        cbo_limit.pack(side="left", padx=(0, 10))
        cbo_limit.bind("<<ComboboxSelected>>", lambda e: self.load_data())

        self.lbl_sheet_rows = tk.Label(
            row1,
            text="📊 SQL: 0 dòng",
            font=("Segoe UI", 9, "bold"),
            fg="#2563eb",
            bg="#dbeafe",
            padx=8,
            pady=3,
            relief="solid",
            bd=1
        )
        self.lbl_sheet_rows.pack(side="left", padx=(0, 10))

        btn_refresh = ttk.Button(row1, text="🔄 Tải Lại", command=self.load_data)
        btn_refresh.pack(side="right")

        # DÒNG 2: CÁC NÚT THAO TÁC NGHIỆP VỤ (CẮT XUỐNG DÒNG THỨ 2 GỌN GÀNG)
        row2 = ttk.Frame(self, padding=(10, 2, 10, 8))
        row2.pack(fill="x", side="top")

        btn_aggregate = tk.Button(
            row2,
            text="⚡ TỔNG HỢP DATA MTI",
            font=("Segoe UI", 9, "bold"),
            bg="#059669",
            fg="white",
            activebackground="#047857",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self.quick_aggregate_mti
        )
        btn_aggregate.pack(side="left", padx=(0, 6))

        btn_push_data_onh = tk.Button(
            row2,
            text="📤 ĐẨY ONH -> SHEET DATA",
            font=("Segoe UI", 9, "bold"),
            bg="#2563eb",
            fg="white",
            activebackground="#1d4ed8",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=8,
            pady=3,
            cursor="hand2",
            command=self.push_onh_to_data_sheet_thread
        )
        btn_push_data_onh.pack(side="left", padx=(0, 6))

        btn_create_sheet = tk.Button(
            row2,
            text="➕ Tạo Sheet Mới",
            font=("Segoe UI", 9, "bold"),
            bg="#10b981",
            fg="white",
            activebackground="#059669",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=8,
            pady=3,
            cursor="hand2",
            command=self.open_create_sheet_dialog
        )
        btn_create_sheet.pack(side="left", padx=(0, 6))

        btn_rename_table = tk.Button(
            row2,
            text="✏️ Đổi Tên Bảng",
            font=("Segoe UI", 9, "bold"),
            bg="#f59e0b",
            fg="white",
            activebackground="#d97706",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=8,
            pady=3,
            cursor="hand2",
            command=self.rename_current_table
        )
        btn_rename_table.pack(side="left", padx=(0, 6))

        btn_import_excel = ttk.Button(row2, text="📤 Nạp Excel", command=self.open_import_dialog)
        btn_import_excel.pack(side="left", padx=(0, 6))

        # NÚT XÓA DÒNG BÔI ĐEN
        self.btn_del_selected = tk.Button(
            row2,
            text="🗑️ XÓA DÒNG",
            font=("Segoe UI", 9, "bold"),
            bg="#ef4444",
            fg="white",
            activebackground="#dc2626",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=8,
            pady=3,
            cursor="hand2",
            command=self.delete_selected_rows
        )
        self.btn_del_selected.pack(side="left", padx=(0, 6))

        # NÚT XÓA TOÀN BỘ DATA SHEET
        self.btn_del_all = tk.Button(
            row2,
            text="💣 XÓA DATA",
            font=("Segoe UI", 9, "bold"),
            bg="#b91c1c",
            fg="white",
            activebackground="#991b1b",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=8,
            pady=3,
            cursor="hand2",
            command=self.clear_entire_table
        )
        self.btn_del_all.pack(side="left", padx=(0, 6))

        # NÚT XÓA HẲN BẢNG KHỎI SQL
        self.btn_drop_table = tk.Button(
            row2,
            text="❌ XÓA BẢNG",
            font=("Segoe UI", 9, "bold"),
            bg="#7f1d1d",
            fg="white",
            activebackground="#450a0a",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=8,
            pady=3,
            cursor="hand2",
            command=self.drop_current_table
        )
        self.btn_drop_table.pack(side="left", padx=(0, 6))

        btn_export = ttk.Button(row2, text="📥 Xuất CSV", command=self.export_csv)
        btn_export.pack(side="left", padx=(0, 6))

        self.lbl_status = ttk.Label(self, text="Đang tải...", padding=6, font=("Segoe UI", 9))
        self.lbl_status.pack(fill="x", side="bottom")

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.tree = ttk.Treeview(tree_frame, show="headings", selectmode="extended")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        # Phím tắt Delete, chuột phải và nhấp đúp để đổi trạng thái Arrived (Yes/No)
        self.tree.bind("<Delete>", self.delete_selected_rows)
        self.tree.bind("<Button-3>", self.show_context_menu)
        self.tree.bind("<Double-1>", self.on_double_click_row)

    def on_double_click_row(self, event=None):
        """Nhấp đúp vào dòng trong Bao_Cao_Tong_Hop (MTI) hoặc (IB) để bật/tắt Arrived (Yes <-> No)."""
        table = self._get_clean_table_name()
        if table not in ('Bao_Cao_Tong_Hop (MTI)', 'Bao_Cao_Tong_Hop (IB)', 'Bao_Cao_Tong_Hop', 'Tổng Hợp Data MTI'):
            return
        iid = self.tree.identify_row(event.y) if event else None
        if not iid:
            sel = self.tree.selection()
            iid = sel[0] if sel else None
        if not iid:
            return
        vals = self.tree.item(iid, "values")
        if not vals or len(vals) < 9:
            return
        # Với MTI vals[0] là Ship Order; Với IB vals[1] là ASN ID hoặc vals[0] là PO ID
        order_key = str(vals[1]).strip() if table == 'Bao_Cao_Tong_Hop (IB)' and len(vals) > 1 else str(vals[0]).strip()
        curr_arrived = str(vals[8]).strip()
        new_arrived = "No" if curr_arrived.lower() == "yes" else "Yes"
        self.set_selected_arrived(new_arrived, target_orders=[order_key])

    def set_selected_arrived(self, arrived_val, target_orders=None):
        table = self._get_clean_table_name()
        if table not in ('Bao_Cao_Tong_Hop (MTI)', 'Bao_Cao_Tong_Hop (IB)', 'Bao_Cao_Tong_Hop', 'Tổng Hợp Data MTI'):
            messagebox.showinfo("Thông báo", "Chức năng tích chọn Arrived (Yes/No) áp dụng cho bảng [Bao_Cao_Tong_Hop (MTI)] và [Bao_Cao_Tong_Hop (IB)].", parent=self)
            return
        if not target_orders:
            selected_iids = self.tree.selection()
            if not selected_iids:
                messagebox.showwarning("Chưa chọn dòng", "Vui lòng chọn ít nhất 1 dòng để đổi Arrived (Yes/No)!", parent=self)
                return
            target_orders = []
            for iid in selected_iids:
                vals = self.tree.item(iid, "values")
                if vals:
                    okey = str(vals[1]).strip() if table == 'Bao_Cao_Tong_Hop (IB)' and len(vals) > 1 else str(vals[0]).strip()
                    target_orders.append(okey)
        try:
            if table == 'Bao_Cao_Tong_Hop (IB)':
                mod_ib = _get_compute_ib_module()
                if mod_ib:
                    for okey in target_orders:
                        mod_ib.set_ib_order_arrived(okey, arrived_val, DB_PATH)
            else:
                mod = _get_compute_module()
                for so in target_orders:
                    mod.set_order_arrived(so, arrived_val, DB_PATH)
            export_dashboard_html()
            self.load_data()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể cập nhật Arrived: {e}", parent=self)

    def quick_aggregate_mti(self):
        try:
            conn = get_db_connection()
            refresh_summary_table(conn)
            conn.close()
            export_dashboard_html()
            messagebox.showinfo(
                "Tổng Hợp Thành Công! ⚡",
                "Đã tự động tính toán toàn bộ:\n"
                "1. [Bao_Cao_Tong_Hop (MTI)]: Đối chiếu [data_mti] và [Put MTI]\n"
                "2. [Bao_Cao_Tong_Hop (IB)]: Đối chiếu [data IB] và [Put IB]\n\n"
                "Dữ liệu bảng [Bao_Cao_Tong_Hop (MTI)] đang được hiển thị ngay bên dưới.",
                parent=self
            )
            self.load_tables(prefer_table='Bao_Cao_Tong_Hop (MTI)')
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể tổng hợp: {e}", parent=self)

    def push_onh_to_data_sheet_thread(self):
        def worker():
            try:
                mod_onh = _get_compute_onh_module()
                if mod_onh and hasattr(mod_onh, 'push_onh_to_data_sheet'):
                    mod_onh.push_onh_to_data_sheet(db_path=DB_PATH)
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Đẩy Sheet Data Thành Công! 🚀",
                        "Đã xóa sạch dữ liệu cũ và nạp mới toàn bộ bảng [Tổng hợp ONH] sang sheet 'Data' trên Google Sheets!\n\n"
                        "Link: https://docs.google.com/spreadsheets/d/1ONBR6oeqRCGGT-vqbiNtOjkbQXQWpNec6i3BLOTSaSE",
                        parent=self
                    ))
                else:
                    self.root.after(0, lambda: messagebox.showerror("Lỗi", "Không tìm thấy module compute_summary_onh!", parent=self))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror(
                    "Lỗi Đẩy Google Sheet",
                    f"Không thể đẩy sang sheet 'Data': {e}\n\n"
                    "Lưu ý: Vui lòng chia sẻ quyền Chỉnh sửa (Editor) cho Service Account:\n"
                    "nh-p-file-mti@rosy-cache-483805-c6.iam.gserviceaccount.com",
                    parent=self
                ))
        threading.Thread(target=worker, daemon=True).start()

    def open_create_sheet_dialog(self):
        CreateEmptySheetDialog(self, on_success_callback=self.on_create_sheet_success)

    def on_create_sheet_success(self, new_table_name=None):
        self.load_tables()
        if new_table_name:
            self.current_table.set(new_table_name)
            self.cbo_tables.set(new_table_name)
        self.load_data()

    def open_import_dialog(self):
        curr = self.cbo_tables.get().strip() or self.current_table.get().strip()
        ExcelImportDialog(self, on_success_callback=self.on_import_success, initial_table=curr)

    def _get_clean_table_name(self) -> str:
        raw = self.cbo_tables.get().strip() or self.current_table.get().strip()
        if ' (' in raw and raw.endswith(' dòng)'):
            return raw.rsplit(' (', 1)[0].strip()
        return raw

    def on_import_success(self, imported_table=None):
        self.load_tables(prefer_table=imported_table)
        self.load_data()

    def load_tables(self, prefer_table=None):
        if not os.path.exists(DB_PATH):
            messagebox.showerror("Lỗi", f"Chưa tìm thấy file SQLite tại:\n{DB_PATH}", parent=self)
            return

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("SELECT name FROM sqlite_master WHERE type IN ('table', 'view') AND name NOT LIKE 'sqlite_%' AND substr(name, 1, 1) != '_' ORDER BY name")
        tables = [r[0] for r in cursor.fetchall()]

        table_counts = {}
        for t in tables:
            try:
                cursor.execute(f"SELECT count(*) FROM [{t}]")
                table_counts[t] = cursor.fetchone()[0]
            except Exception:
                table_counts[t] = 0
        conn.close()

        self.table_counts = table_counts
        display_tables = [f"{t} ({table_counts.get(t, 0):,} dòng)" for t in tables]
        self.cbo_tables['values'] = display_tables

        def _find_display(t_name):
            for dt in display_tables:
                if dt.startswith(t_name + " (") or dt == t_name:
                    return dt
            return t_name

        clean_curr = self._get_clean_table_name()
        if prefer_table and prefer_table in tables:
            self.current_table.set(_find_display(prefer_table))
            self.cbo_tables.set(_find_display(prefer_table))
        elif 'Bao_Cao_Tong_Hop (MTI)' in tables and not clean_curr:
            self.current_table.set(_find_display('Bao_Cao_Tong_Hop (MTI)'))
            self.cbo_tables.set(_find_display('Bao_Cao_Tong_Hop (MTI)'))
        elif 'Bao_Cao_Tong_Hop' in tables and not clean_curr:
            self.current_table.set(_find_display('Bao_Cao_Tong_Hop'))
            self.cbo_tables.set(_find_display('Bao_Cao_Tong_Hop'))
        elif 'Bao_Cao_Tong_Hop (IB)' in tables and not clean_curr:
            self.current_table.set(_find_display('Bao_Cao_Tong_Hop (IB)'))
            self.cbo_tables.set(_find_display('Bao_Cao_Tong_Hop (IB)'))
        elif 'Put MTI' in tables and not clean_curr:
            self.current_table.set(_find_display('Put MTI'))
            self.cbo_tables.set(_find_display('Put MTI'))
        elif tables and (clean_curr not in tables):
            self.current_table.set(_find_display(tables[0]))
            self.cbo_tables.set(_find_display(tables[0]))

        if self.current_table.get():
            self.load_data()

    def rename_current_table(self):
        curr = self._get_clean_table_name()
        if not curr:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn bảng cần đổi tên!", parent=self)
            return

        from tkinter import simpledialog
        new_name = simpledialog.askstring("Đổi Tên Bảng SQL", f"Nhập tên mới cho bảng [{curr}]:", initialvalue=curr, parent=self)
        if not new_name or not new_name.strip():
            return
        new_name = new_name.strip()
        if new_name == curr:
            return

        if any(ch in new_name for ch in (';', "'", '"')):
            messagebox.showwarning("Cảnh báo", "Tên bảng không được chứa ký tự đặc biệt như dấu nháy hoặc chấm phẩy!", parent=self)
            return

        try:
            conn = get_db_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT count(*) FROM sqlite_master WHERE type IN ('table', 'view') AND name=?", (new_name,))
            if cursor.fetchone()[0] > 0:
                overwrite = messagebox.askyesno(
                    "Bảng đã tồn tại",
                    f"Bảng [{new_name}] đã tồn tại trong SQL!\nBạn có muốn ghi đè bảng [{curr}] sang tên [{new_name}] không?",
                    parent=self
                )
                if not overwrite:
                    conn.close()
                    return
                cursor.execute(f'DROP TABLE IF EXISTS [{new_name}]')
                cursor.execute(f'DROP VIEW IF EXISTS [{new_name}]')

            try:
                cursor.execute(f'ALTER TABLE [{curr}] RENAME TO [{new_name}]')
            except Exception:
                cursor.execute(f'CREATE TABLE [{new_name}] AS SELECT * FROM [{curr}]')
                cursor.execute(f'DROP TABLE IF EXISTS [{curr}]')

            conn.commit()
            conn.close()

            messagebox.showinfo("Thành công", f"Đã đổi tên bảng từ [{curr}] sang [{new_name}] thành công!", parent=self)
            self.load_tables(prefer_table=new_name)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể đổi tên bảng: {e}", parent=self)

    def drop_current_table(self):
        table_name = self._get_clean_table_name()
        if not table_name:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn bảng cần xóa khỏi SQL!", parent=self)
            return

        core_tables = ['Put MTI', 'Put IB', 'ONH', 'Cell', 'data_mti']
        if table_name in core_tables:
            confirm = messagebox.askyesno(
                "Cảnh báo xóa bảng nguồn",
                f"Bảng [{table_name}] là bảng đồng bộ dữ liệu chính từ Google Sheet!\n\n"
                f"Bạn có THỰC SỰ muốn XÓA BỎ HOÀN TOÀN bảng [{table_name}] khỏi SQLite database không?\n"
                f"(Lưu ý: Bảng này và toàn bộ dữ liệu của nó sẽ bị xóa vĩnh viễn)",
                parent=self,
                icon='warning'
            )
        else:
            confirm = messagebox.askyesno(
                "Xác nhận xóa bảng",
                f"Bạn có chắc chắn muốn XÓA BỎ BẢNG [{table_name}] khỏi SQLite database?\n"
                f"(Bảng này sẽ bị gỡ bỏ vĩnh viễn khỏi danh sách)",
                parent=self
            )

        if not confirm:
            return

        try:
            conn = sqlite3.connect(DB_PATH)
            c = conn.cursor()
            c.execute(f"DROP TABLE IF EXISTS [{table_name}]")
            c.execute(f"DROP VIEW IF EXISTS [{table_name}]")
            conn.commit()
            conn.close()

            messagebox.showinfo("Đã xóa", f"Đã xóa thành công bảng [{table_name}] khỏi SQLite database!", parent=self)
            self.current_table.set("")
            self.load_tables()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Không thể xóa bảng: {e}", parent=self)

    def clear_search(self):
        self.search_var.set("")
        self.load_data()

    def load_data(self):
        table = self._get_clean_table_name()
        if not table or not os.path.exists(DB_PATH):
            return

        self.lbl_status.config(text="⏳ Đang tải và tìm kiếm dữ liệu...", foreground="#2563eb")
        self.update_idletasks()

        conn = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute(f"PRAGMA table_info([{table}])")
        columns = [col[1] for col in cursor.fetchall()]

        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, minwidth=70)

        for item in self.tree.get_children():
            self.tree.delete(item)

        search = self.search_var.get().strip()
        query = f"SELECT rowid, * FROM [{table}]"
        params = []

        if search:
            where_clauses = [f"[{col}] LIKE ?" for col in columns]
            query += " WHERE " + " OR ".join(where_clauses)
            params = [f"%{search}%"] * len(columns)

        limit_val = self.limit_var.get()
        if limit_val != "Tất cả":
            query += f" LIMIT {int(limit_val)}"

        cursor.execute(query, params)
        rows = cursor.fetchall()

        cursor.execute(f"SELECT count(*) FROM [{table}]")
        total_rows = cursor.fetchone()[0]
        conn.close()

        if hasattr(self, 'lbl_sheet_rows'):
            self.lbl_sheet_rows.config(text=f"📊 SQL: {total_rows:,} dòng")

        for row in rows:
            rowid = row[0]
            vals = row[1:]
            self.tree.insert("", "end", iid=str(rowid), values=vals)

        if search:
            if rows:
                self.lbl_status.config(
                    text=f"✅ Tìm thấy {len(rows):,} dòng khớp với từ khóa '{search}' trong bảng [{table}] (Tổng SQL: {total_rows:,} dòng)",
                    foreground="#16a34a"
                )
            else:
                self.lbl_status.config(
                    text=f"❌ Không tìm thấy dòng nào khớp với '{search}' trong bảng [{table}]! (Gợi ý: Kiểm tra lại tên bảng hoặc thử từ khóa ngắn hơn)",
                    foreground="#dc2626"
                )
        else:
            self.lbl_status.config(
                text=f"Bảng: [{table}] | Đang hiển thị: {len(rows):,} dòng / Tổng cộng trong SQL: {total_rows:,} dòng  (Mẹo: Bôi đen dòng rồi bấm 'Delete' hoặc chuột phải để xóa)",
                foreground="#334155"
            )

    def delete_selected_rows(self, event=None):
        """Bôi đen xóa các dòng đang chọn trong SQL."""
        selected_iids = self.tree.selection()
        if not selected_iids:
            messagebox.showwarning(
                "Chưa chọn dòng",
                "Vui lòng bôi đen (chọn) các dòng muốn xóa trong bảng trước!\n(Bạn có thể giữ phím Shift hoặc Ctrl để chọn nhiều dòng cùng lúc).",
                parent=self
            )
            return

        table = self._get_clean_table_name()
        count = len(selected_iids)
        confirm = messagebox.askyesno(
            "Xác nhận xóa dữ liệu",
            f"Bạn có chắc chắn muốn xóa {count:,} dòng đang chọn khỏi bảng [{table}] không?\n\nThao tác này sẽ xóa trực tiếp dữ liệu trong SQL!",
            parent=self,
            icon="warning"
        )
        if not confirm:
            return

        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            chunk_size = 400
            iid_list = list(selected_iids)
            for i in range(0, len(iid_list), chunk_size):
                chunk = iid_list[i:i + chunk_size]
                placeholders = ",".join(["?"] * len(chunk))
                cursor.execute(f"DELETE FROM [{table}] WHERE rowid IN ({placeholders})", chunk)
            conn.commit()
            conn.close()

            for iid in selected_iids:
                if self.tree.exists(iid):
                    self.tree.delete(iid)

            messagebox.showinfo("Thành công", f"Đã xóa thành công {count:,} dòng khỏi bảng [{table}]!", parent=self)
            self.load_data()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Xóa dòng thất bại: {e}", parent=self)

    def clear_entire_table(self):
        """Click xóa toàn bộ data ở sheet/bảng đã chọn."""
        table = self._get_clean_table_name()
        if not table:
            return

        confirm = messagebox.askyesno(
            "CẢNH BÁO: XÓA TOÀN BỘ DATA SHEET",
            f"⚠️ BẠN CÓ CHẮC CHẮN MUỐN XÓA TOÀN BỘ DATA CỦA SHEET/BẢNG [{table}] KHÔNG?\n\n- Toàn bộ dữ liệu trong bảng sẽ bị xóa sạch khỏi SQL.\n- Cấu trúc các cột vẫn được giữ nguyên để đón nhận dữ liệu mới.",
            parent=self,
            icon="warning"
        )
        if not confirm:
            return

        try:
            conn = sqlite3.connect(DB_PATH)
            cursor = conn.cursor()
            cursor.execute(f"DELETE FROM [{table}]")
            conn.commit()
            cursor.execute("VACUUM")
            conn.close()

            messagebox.showinfo("Thành công", f"Đã xóa sạch toàn bộ dữ liệu của bảng [{table}] thành công!", parent=self)
            self.load_data()
        except Exception as e:
            messagebox.showerror("Lỗi", f"Xóa toàn bộ data thất bại: {e}", parent=self)

    def show_context_menu(self, event):
        """Hiển thị menu chuột phải khi click vào dòng."""
        iid = self.tree.identify_row(event.y)
        if iid:
            if iid not in self.tree.selection():
                self.tree.selection_set(iid)

        menu = tk.Menu(self, tearoff=0)
        if self.current_table.get() in ('Bao_Cao_Tong_Hop', 'Tổng Hợp Data MTI'):
            menu.add_command(label="✅ Đánh dấu Arrived = Yes (Đã đến - Arriving)", command=lambda: self.set_selected_arrived("Yes"))
            menu.add_command(label="⏳ Đánh dấu Arrived = No (Chưa đến - Pending)", command=lambda: self.set_selected_arrived("No"))
            menu.add_separator()
        menu.add_command(label="🗑️ Xóa các dòng đang bôi đen (Delete)", command=self.delete_selected_rows)
        menu.add_separator()
        menu.add_command(label="💣 Xóa toàn bộ data sheet này", command=self.clear_entire_table)
        menu.add_separator()
        menu.add_command(label="🔄 Làm mới dữ liệu", command=self.load_data)
        menu.post(event.x_root, event.y_root)

    def export_csv(self):
        table = self.current_table.get()
        if not table:
            return
        file_path = filedialog.asksaveasfilename(
            parent=self,
            defaultextension=".csv",
            filetypes=[("CSV UTF-8", "*.csv"), ("All Files", "*.*")],
            initialfile=f"{table}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        if not file_path:
            return
        try:
            conn = sqlite3.connect(DB_PATH)
            df = pd.read_sql_query(f"SELECT * FROM [{table}]", conn)
            conn.close()
            df.to_csv(file_path, index=False, encoding='utf-8-sig')
            messagebox.showinfo("Thành công", f"Đã xuất {len(df):,} dòng ra file:\n{file_path}", parent=self)
        except Exception as e:
            messagebox.showerror("Lỗi", f"Xuất file thất bại: {e}", parent=self)


# ==================== GIAO DIỆN CHÍNH CỦA APP ====================
class MainSyncApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Auto Sync Data Google Sheets ➔ SQL (Tự Động & Real-time)")
        self.root.geometry("860x600")
        self.root.minsize(760, 520)

        self.log_queue = queue.Queue()
        self.is_syncing = False
        self.auto_sync_enabled = True
        self.seconds_remaining = INTERVAL_SECONDS
        self.last_sync_time_str = "Chưa đồng bộ"

        self.last_mod_1 = None
        self.last_mod_2 = None

        self.setup_ui()
        self.setup_styles()

        self.root.after(200, self.process_log_queue)
        self.root.after(1000, self.countdown_timer)

        # Chạy đồng bộ lần đầu khi khởi động app
        self.trigger_sync_thread()

        # Bật luồng theo dõi thay đổi Real-time trên Google Sheets
        self.start_realtime_watcher()

        # Tự động chạy máy chủ Realtime Dashboard (localhost:8500) ngầm
        self.start_dashboard_server()

    def start_dashboard_server(self):
        def _server():
            try:
                mod = None
                for dash_py in [
                    r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\dashboard.py',
                    r'c:\Users\srvn0514\Desktop\mới\dashboard.py',
                    r'C:\Users\srvn0514\Desktop\SQL\dashboard.py',
                    os.path.join(BUNDLE_DIR, 'dashboard.py'),
                    os.path.join(EXE_DIR, 'dashboard.py')
                ]:
                    if os.path.exists(dash_py):
                        import importlib.util
                        spec = importlib.util.spec_from_file_location("dash_module", dash_py)
                        mod = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(mod)
                        break
                if not mod:
                    try:
                        import dashboard as mod
                    except Exception:
                        pass
                if mod and hasattr(mod, 'run_server'):
                    mod.run_server(port=8500, open_browser=False)
            except Exception as e:
                print(f"Lỗi khởi động dashboard server ngầm: {e}")
        t = threading.Thread(target=_server, daemon=True)
        t.start()

    def setup_styles(self):
        style = ttk.Style()
        try:
            style.theme_use('clam')
        except Exception:
            pass

    def setup_ui(self):
        header_frame = tk.Frame(self.root, bg="#1e293b", padx=16, pady=12)
        header_frame.pack(fill="x", side="top")

        lbl_title = tk.Label(
            header_frame,
            text="HỆ THỐNG ĐỒNG BỘ DỮ LIỆU GOOGLE SHEETS ➔ SQL",
            font=("Segoe UI", 13, "bold"),
            fg="#f8fafc",
            bg="#1e293b"
        )
        lbl_title.pack(anchor="w")

        lbl_sub = tk.Label(
            header_frame,
            text="⚡ Tự động cập nhật ngay khi sheet nguồn đổi hoặc định kỳ 40' | Put MTI, Put IB, data_mti dán tiếp không lặp | Hỗ trợ xóa data SQL",
            font=("Segoe UI", 9),
            fg="#94a3b8",
            bg="#1e293b"
        )
        lbl_sub.pack(anchor="w", pady=(2, 0))

        # Control panel
        control_frame = ttk.LabelFrame(self.root, text=" Điều Khiển & Trạng Thái ", padding=12)
        control_frame.pack(fill="x", padx=14, pady=10)

        # Hàng 1: Trạng thái & đếm ngược & Realtime Watcher
        status_row = ttk.Frame(control_frame)
        status_row.pack(fill="x", pady=(0, 10))

        ttk.Label(status_row, text="Trạng thái tự động:", font=("Segoe UI", 10, "bold")).pack(side="left")
        self.lbl_auto_status = ttk.Label(status_row, text="Đang BẬT", foreground="#16a34a", font=("Segoe UI", 10, "bold"))
        self.lbl_auto_status.pack(side="left", padx=(6, 16))

        ttk.Label(status_row, text="Đồng bộ định kỳ sau:", font=("Segoe UI", 10)).pack(side="left")
        self.lbl_countdown = ttk.Label(status_row, text="40:00", foreground="#2563eb", font=("Segoe UI", 11, "bold"))
        self.lbl_countdown.pack(side="left", padx=(6, 16))

        ttk.Label(status_row, text="| ⚡ Watcher:", font=("Segoe UI", 10, "bold")).pack(side="left", padx=(0, 4))
        self.lbl_watcher_status = ttk.Label(status_row, text="🟢 Theo dõi sheet nguồn (Cập nhật ngay khi có sửa đổi)", foreground="#059669", font=("Segoe UI", 9, "bold"))
        self.lbl_watcher_status.pack(side="left")

        # Hàng 2: Các nút bấm chính
        btn_row = ttk.Frame(control_frame)
        btn_row.pack(fill="x")

        self.btn_sync_now = ttk.Button(btn_row, text="⚡ ĐỒNG BỘ NGAY", command=self.trigger_sync_thread)
        self.btn_sync_now.pack(side="left", padx=(0, 6))

        self.btn_create_sheet_excel = tk.Button(
            btn_row,
            text="➕ TẠO SHEET TỪ EXCEL",
            font=("Segoe UI", 9, "bold"),
            bg="#10b981",
            fg="white",
            activebackground="#059669",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=8,
            pady=3,
            cursor="hand2",
            command=self.open_create_sheet_from_excel
        )
        self.btn_create_sheet_excel.pack(side="left", padx=(0, 6))

        self.btn_import_excel = ttk.Button(btn_row, text="📤 Nạp Excel Cũ", command=self.open_excel_import)
        self.btn_import_excel.pack(side="left", padx=(0, 6))

        self.btn_toggle_auto = ttk.Button(btn_row, text="⏸️ Tạm Dừng", command=self.toggle_auto_sync)
        self.btn_toggle_auto.pack(side="left", padx=(0, 6))

        self.btn_open_html_dashboard = tk.Button(
            btn_row,
            text="📊 DASHBOARD",
            font=("Segoe UI", 9, "bold"),
            bg="#6366f1",
            fg="white",
            activebackground="#4f46e5",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=8,
            pady=3,
            cursor="hand2",
            command=self.open_html_dashboard
        )
        self.btn_open_html_dashboard.pack(side="left", padx=(0, 6))

        # NÚT XÓA DATA SQL TRỰC TIẾP TRÊN TRANG CHỦ
        self.btn_del_sql = tk.Button(
            btn_row,
            text="🗑️ XÓA DATA SQL",
            font=("Segoe UI", 9, "bold"),
            bg="#ef4444",
            fg="white",
            activebackground="#dc2626",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self.open_data_viewer
        )
        self.btn_del_sql.pack(side="right", padx=(6, 0))

        self.btn_view_data = tk.Button(
            btn_row,
            text="👁️ XEM DỮ LIỆU SQL",
            font=("Segoe UI", 9, "bold"),
            bg="#0284c7",
            fg="white",
            activebackground="#0369a1",
            activeforeground="white",
            relief="raised",
            bd=2,
            padx=10,
            pady=3,
            cursor="hand2",
            command=self.open_data_viewer
        )
        self.btn_view_data.pack(side="right")

        # Log Frame
        log_frame = ttk.LabelFrame(self.root, text=" Nhật Ký Hoạt Động (Logs) ", padding=8)
        log_frame.pack(fill="both", expand=True, padx=14, pady=(0, 10))

        self.txt_log = tk.Text(
            log_frame,
            wrap="word",
            bg="#0f172a",
            fg="#e2e8f0",
            insertbackground="white",
            font=("Consolas", 9),
            padx=8,
            pady=8
        )
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=log_scroll.set)

        log_scroll.pack(side="right", fill="y")
        self.txt_log.pack(fill="both", expand=True)

        # Footer info
        footer_frame = ttk.Frame(self.root, padding=6)
        footer_frame.pack(fill="x", side="bottom")

        self.lbl_last_sync = ttk.Label(footer_frame, text=f"Lần đồng bộ gần nhất: {self.last_sync_time_str}", font=("Segoe UI", 9))
        self.lbl_last_sync.pack(side="left")

        lbl_db_path = ttk.Label(footer_frame, text=f"File SQL: {os.path.basename(DB_PATH)}", font=("Segoe UI", 9), foreground="#64748b")
        lbl_db_path.pack(side="right")

    def log(self, message: str):
        self.log_queue.put(message)

    def process_log_queue(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get_nowait()
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            line = f"[{ts}] {msg}\n"
            self.txt_log.insert("end", line)
            self.txt_log.see("end")

            try:
                with open(LOG_FILE, 'a', encoding='utf-8') as f:
                    f.write(line)
            except Exception:
                pass

        self.root.after(200, self.process_log_queue)

    def countdown_timer(self):
        if self.auto_sync_enabled and not self.is_syncing:
            self.seconds_remaining -= 1
            if self.seconds_remaining <= 0:
                self.seconds_remaining = INTERVAL_SECONDS
                self.trigger_sync_thread()
            
            mins = self.seconds_remaining // 60
            secs = self.seconds_remaining % 60
            self.lbl_countdown.config(text=f"{mins:02d}:{secs:02d}")

        self.root.after(1000, self.countdown_timer)

    def toggle_auto_sync(self):
        self.auto_sync_enabled = not self.auto_sync_enabled
        if self.auto_sync_enabled:
            self.lbl_auto_status.config(text="Đang BẬT", foreground="#16a34a")
            self.btn_toggle_auto.config(text="⏸️ Tạm Dừng")
            self.lbl_watcher_status.config(text="🟢 Theo dõi sheet nguồn (Cập nhật ngay khi có sửa đổi)", foreground="#059669")
            self.log("Đã BẬT chế độ tự động đồng bộ & Real-time Watcher.")
        else:
            self.lbl_auto_status.config(text="Đã TẠM DỪNG", foreground="#dc2626")
            self.btn_toggle_auto.config(text="▶️ Bật Lại")
            self.lbl_countdown.config(text="--:--")
            self.lbl_watcher_status.config(text="⏸️ Watcher đã tạm dừng", foreground="#dc2626")
            self.log("Đã TẠM DỪNG tự động đồng bộ & Watcher.")

    def trigger_sync_thread(self):
        if self.is_syncing:
            return
        self.is_syncing = True
        self.btn_sync_now.config(state="disabled")

        def worker():
            try:
                sync_core_logic(log_callback=self.log)
                now_str = datetime.datetime.now().strftime("%H:%M:%S ngày %d/%m/%Y")
                self.last_sync_time_str = now_str
                self.root.after(0, lambda: self.lbl_last_sync.config(text=f"Lần đồng bộ gần nhất: {now_str}"))
            except Exception as e:
                self.log(f"[LỖI ĐỒNG BỘ]: {e}")
            finally:
                self.is_syncing = False
                self.seconds_remaining = INTERVAL_SECONDS
                self.root.after(0, lambda: self.btn_sync_now.config(state="normal"))

        thread = threading.Thread(target=worker, daemon=True)
        thread.start()

    def start_realtime_watcher(self):
        """Luồng ngầm theo dõi thay đổi thời gian thực trên Google Sheets qua modifiedTime."""
        def watcher_loop():
            creds_path = get_credentials_path()
            if not os.path.exists(creds_path):
                return

            session = None
            try:
                scope = ['https://www.googleapis.com/auth/drive.readonly']
                creds = Credentials.from_service_account_file(creds_path, scopes=scope)
                session = google.auth.transport.requests.AuthorizedSession(creds)
            except Exception as e:
                self.log(f"[CẢNH BÁO WATCHER] Không thể khởi tạo session Drive API: {e}")
                return

            def get_mod(sid):
                try:
                    url = f"https://www.googleapis.com/drive/v3/files/{sid}?fields=modifiedTime,name"
                    r = session.get(url, timeout=8)
                    if r.status_code == 200:
                        return r.json().get('modifiedTime'), r.json().get('name', sid)
                except Exception:
                    pass
                return None, sid

            # Lấy thời điểm modified ban đầu
            m1, name1 = get_mod(SPREADSHEET_ID_1)
            m2, name2 = get_mod(SPREADSHEET_ID_2)
            self.last_mod_1 = m1
            self.last_mod_2 = m2

            while True:
                time.sleep(20)  # Kiểm tra mỗi 20 giây một lần (nhẹ, nhanh)
                if not self.auto_sync_enabled or self.is_syncing:
                    continue

                curr1, _ = get_mod(SPREADSHEET_ID_1)
                curr2, _ = get_mod(SPREADSHEET_ID_2)

                changed_sheets = []
                if curr1 and self.last_mod_1 and curr1 != self.last_mod_1:
                    changed_sheets.append(f"Link 1: '{name1}'")
                    self.last_mod_1 = curr1
                elif curr1 and not self.last_mod_1:
                    self.last_mod_1 = curr1

                if curr2 and self.last_mod_2 and curr2 != self.last_mod_2:
                    changed_sheets.append(f"Link 2: '{name2}'")
                    self.last_mod_2 = curr2
                elif curr2 and not self.last_mod_2:
                    self.last_mod_2 = curr2

                if changed_sheets:
                    reason = " & ".join(changed_sheets)
                    self.log(f"\n⚡ [REAL-TIME WATCHER] Phát hiện dữ liệu tại {reason} vừa được chỉnh sửa!")
                    self.log("   ➔ Kích hoạt TỰ ĐỘNG ĐỒNG BỘ NGAY LẬP TỨC (không cần chờ 40 phút)...")
                    self.root.after(0, self.trigger_sync_thread)

        thread = threading.Thread(target=watcher_loop, daemon=True)
        thread.start()

    def open_data_viewer(self):
        SQLDataViewer(self.root)

    def open_html_dashboard(self):
        # Luôn làm mới dữ liệu ra HTML trước khi mở
        export_dashboard_html()
        candidates = [
            r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\index.html',
            os.path.join(EXE_DIR, 'index.html'),
            r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\index.html',
            r'C:\Users\srvn0514\Desktop\SQL\index.html',
            os.path.join(BUNDLE_DIR, 'index.html'),
        ]
        html_file = None
        for p in candidates:
            if os.path.exists(p):
                html_file = p
                break

        if html_file:
            import webbrowser
            webbrowser.open(f"file:///{os.path.abspath(html_file).replace(os.sep, '/')}")
        else:
            messagebox.showinfo("Thông báo", "File Dashboard HTML (index.html) chưa được tạo hoặc chưa tìm thấy.", parent=self.root)

    def open_create_sheet_from_excel(self):
        CreateEmptySheetDialog(self.root, on_success_callback=self.on_create_sheet_success_main)

    def on_create_sheet_success_main(self, new_table_name=None):
        self.log(f"Đã tạo thành công Sheet / Bảng [{new_table_name}] trên SQL!")
        # Mở ngay cửa sổ xem dữ liệu và hiển thị bảng mới tạo
        SQLDataViewer(self.root, initial_table=new_table_name)

    def open_excel_import(self):
        ExcelImportDialog(self.root, on_success_callback=lambda: self.log("Đã nạp dữ liệu từ file Excel vào SQL thành công."))


if __name__ == '__main__':
    root = tk.Tk()
    app = MainSyncApp(root)
    root.mainloop()
