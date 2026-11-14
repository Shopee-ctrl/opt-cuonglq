# -*- coding: utf-8 -*-
"""
INBOUND, MTI & RTS BACKLOG DASHBOARD & REALTIME SERVER
- Tự động đọc dữ liệu từ SQLite (Data MTI.db)
- Hỗ trợ đầy đủ 3 Phân Hệ qua menu BUCKET:
  1. MTI: data_mti đối chiếu Put MTI -> Bao_Cao_Tong_Hop (MTI)
  2. Inbound: data IB đối chiếu Put IB theo ASN ID == Inbound ID -> Bao_Cao_Tong_Hop (IB)
  3. RTS: Bảng [RTS] tổng hợp theo Order Number, sum Total SKU, QTY, và bảng chi tiết từng dòng
  4. ALL: Toàn bộ đơn hàng tổng hợp
- Cột ARRIVED có 3 lựa chọn: No (Chưa đến), Yes (Đã đến), Done (Hoàn thành)
- Cột Status của RTS tự động cập nhật từ Sheet nhưng có thể chỉnh sửa trực tiếp trên Web và lưu vào SQL!
"""

import os
import sys
import json
import time
import socket
import sqlite3
import webbrowser
try:
    from http.server import ThreadingHTTPServer as HTTPServer, SimpleHTTPRequestHandler
except ImportError:
    from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import re
import pandas as pd

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)
for extra in [r'C:\Users\srvn0514\Desktop\mới', r'C:\Users\srvn0514\Desktop\SQL', r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL']:
    if os.path.exists(extra) and extra not in sys.path:
        sys.path.append(extra)

DB_PATH = r'C:\Users\srvn0514\Desktop\SQL\Data MTI.db'
HTML_FILE = os.path.join(BASE_DIR, 'index.html')
RTS_HTML_FILE = os.path.join(BASE_DIR, 'rts.html')
OVERVIEW_HTML_FILE = os.path.join(BASE_DIR, 'overview_vndl.html')
TRUCK_HTML_FILE = os.path.join(BASE_DIR, 'truck_tracker.html')
PORT = 8500


def get_db_connection():
    if not os.path.exists(DB_PATH):
        raise FileNotFoundError(f"Không tìm thấy file Database SQLite tại: {DB_PATH}")
    return sqlite3.connect(DB_PATH)


def _get_compute_module():
    import importlib.util
    for p in [
        os.path.join(BASE_DIR, 'compute_summary.py'),
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\compute_summary.py',
        r'c:\Users\srvn0514\Desktop\mới\compute_summary.py',
        r'C:\Users\srvn0514\Desktop\SQL\compute_summary.py',
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
        os.path.join(BASE_DIR, 'compute_summary_ib.py'),
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\compute_summary_ib.py',
        r'c:\Users\srvn0514\Desktop\mới\compute_summary_ib.py',
        r'C:\Users\srvn0514\Desktop\SQL\compute_summary_ib.py',
    ]:
        if os.path.exists(p):
            spec = importlib.util.spec_from_file_location("compute_summary_ib_disk", p)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    import compute_summary_ib
    return compute_summary_ib


def _get_compute_rts_module():
    import importlib.util
    for p in [
        os.path.join(BASE_DIR, 'compute_summary_rts.py'),
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\compute_summary_rts.py',
        r'c:\Users\srvn0514\Desktop\mới\compute_summary_rts.py',
        r'C:\Users\srvn0514\Desktop\SQL\compute_summary_rts.py',
    ]:
        if os.path.exists(p):
            spec = importlib.util.spec_from_file_location("compute_summary_rts_disk", p)
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            return mod
    import compute_summary_rts
    return compute_summary_rts


def _get_compute_overview_module():
    import importlib.util
    for p in [
        os.path.join(BASE_DIR, 'compute_overview_vndl.py'),
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\compute_overview_vndl.py',
        r'c:\Users\srvn0514\Desktop\mới\compute_overview_vndl.py',
        r'C:\Users\srvn0514\Desktop\SQL\compute_overview_vndl.py',
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


def compute_overview_summary_data():
    """Thuật toán tổng hợp Capacity cho Over view VNDL"""
    try:
        mod = _get_compute_overview_module()
        if mod and hasattr(mod, 'compute_and_save_overview_vndl'):
            return mod.compute_and_save_overview_vndl(DB_PATH)
    except Exception as e:
        print(f"Lỗi compute_overview_summary_data: {e}")
    return {'items': [], 'total': {}, 'chart_data': []}


def compute_inbound_summary():
    """Thuật toán đối chiếu MTI giữa data_mti và Put MTI"""
    try:
        mod = _get_compute_module()
        return mod.compute_and_save_summary(DB_PATH)
    except Exception as e:
        print(f"Lỗi compute_inbound_summary: {e}")
        conn = get_db_connection()
        try:
            return pd.read_sql_query("SELECT * FROM [Bao_Cao_Tong_Hop (MTI)]", conn)
        except Exception:
            try:
                return pd.read_sql_query("SELECT * FROM Bao_Cao_Tong_Hop", conn)
            except Exception:
                return pd.DataFrame()
        finally:
            conn.close()


def compute_inbound_summary_ib():
    """Thuật toán đối chiếu Inbound giữa data IB và Put IB"""
    try:
        mod = _get_compute_ib_module()
        return mod.compute_and_save_summary_ib(DB_PATH)
    except Exception as e:
        print(f"Lỗi compute_inbound_summary_ib: {e}")
        conn = get_db_connection()
        try:
            return pd.read_sql_query("SELECT * FROM [Bao_Cao_Tong_Hop (IB)]", conn)
        except Exception:
            return pd.DataFrame()
        finally:
            conn.close()


def compute_inbound_summary_rts():
    """Thuật toán tổng hợp RTS"""
    try:
        mod = _get_compute_rts_module()
        return mod.compute_and_save_summary_rts(DB_PATH)
    except Exception as e:
        print(f"Lỗi compute_inbound_summary_rts: {e}")
        conn = get_db_connection()
        try:
            df_o = pd.read_sql_query("SELECT * FROM [Bao_Cao_Tong_Hop (RTS)]", conn)
            df_i = pd.read_sql_query("SELECT * FROM [RTS]", conn)
            return df_o, df_i
        except Exception:
            return pd.DataFrame(), pd.DataFrame()
        finally:
            conn.close()


def calc_kpis(items):
    total_orders = len(items)
    total_qty = sum(int(r.get('Quantity', 0) or 0) for r in items)
    total_put = sum(int(r.get('Putaway', 0) or 0) for r in items)
    total_backlog = sum(int(r.get('Backlog', 0) or 0) for r in items)
    count_pending = sum(1 for r in items if r.get('Status') == 'Pending')
    count_arriving = sum(1 for r in items if r.get('Status') == 'Arriving')
    count_in_prog = sum(1 for r in items if r.get('Status') == 'In Progress')
    count_done = sum(1 for r in items if r.get('Status') == 'Done')
    rate = round((total_put / total_qty * 100), 1) if total_qty > 0 else 0.0

    return {
        "total_orders": total_orders,
        "total_quantity": total_qty,
        "total_putaway": total_put,
        "total_backlog": total_backlog,
        "count_pending": count_pending,
        "count_arriving": count_arriving,
        "count_in_progress": count_in_prog,
        "count_done": count_done,
        "completion_rate": rate
    }


def calc_rts_kpis(orders, items):
    total_orders = len(orders)
    total_items = len(items)
    total_qty = sum(int(r.get('QTY', 0) or 0) for r in items)
    distinct_skus = len(set(r.get('SKU', '') for r in items if r.get('SKU')))

    count_alloc = sum(1 for r in items if r.get('Status') == 'Allocated')
    count_outbound = sum(1 for r in items if r.get('Status') == 'Outbound')
    count_pend_pick = sum(1 for r in items if r.get('Status') == 'Pending Pick')
    count_done = sum(1 for r in items if r.get('Status') == 'Done')

    outbound_qty = sum(int(r.get('QTY', 0) or 0) for r in items if r.get('Status') in ('Outbound', 'Done'))
    rate = round((outbound_qty / total_qty * 100), 1) if total_qty > 0 else 0.0

    return {
        "total_orders": total_orders,
        "total_items": total_items,
        "distinct_skus": distinct_skus,
        "total_qty": total_qty,
        "outbound_qty": outbound_qty,
        "count_allocated": count_alloc,
        "count_outbound": count_outbound,
        "count_pending_pick": count_pend_pick,
        "count_done": count_done,
        "completion_rate": rate
    }


def query_actual_putaway_records():
    """Tổng hợp năng suất Putaway theo Operator và ngày từ [Put MTI] và [Put IB]"""
    conn = get_db_connection()
    c = conn.cursor()
    mti_puts = []
    c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='Put MTI'")
    if c.fetchone()[0] > 0:
        try:
            df = pd.read_sql_query("SELECT [Putaway Time], [Quantity], [Operator] FROM [Put MTI] WHERE [Putaway Time] IS NOT NULL", conn)
            if not df.empty:
                df['qty'] = pd.to_numeric(df['Quantity'], errors='coerce').fillna(0).astype(int)
                df['date'] = df['Putaway Time'].astype(str).str.split(' ').str[0].str.strip()
                df['op'] = df['Operator'].fillna('Unknown').astype(str).str.strip()
                def _std_date(d):
                    if '-' in d:
                        p = d.split('-')
                        if len(p) == 3 and len(p[0]) == 4:
                            return f"{p[2].zfill(2)}/{p[1].zfill(2)}/{p[0]}"
                    return d
                df['date'] = df['date'].apply(_std_date)
                grp = df.groupby(['date', 'op']).agg({'qty': 'sum', 'Operator': 'count'}).rename(columns={'Operator': 'cnt'}).reset_index()
                mti_puts = grp.to_dict('records')
        except Exception as e_m:
            print(f"Lỗi query_actual_putaway_records MTI: {e_m}")

    ib_puts = []
    c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='Put IB'")
    if c.fetchone()[0] > 0:
        try:
            df = pd.read_sql_query("SELECT [Time], [Quantity On Rack], [Operator] FROM [Put IB] WHERE [Time] IS NOT NULL", conn)
            if not df.empty:
                df['qty'] = pd.to_numeric(df['Quantity On Rack'], errors='coerce').fillna(0).astype(int)
                df['date'] = df['Time'].astype(str).str.split(' ').str[0].str.strip()
                df['op'] = df['Operator'].fillna('Unknown').astype(str).str.strip()
                def _std_date(d):
                    if '-' in d:
                        p = d.split('-')
                        if len(p) == 3 and len(p[0]) == 4:
                            return f"{p[2].zfill(2)}/{p[1].zfill(2)}/{p[0]}"
                    return d
                df['date'] = df['date'].apply(_std_date)
                grp = df.groupby(['date', 'op']).agg({'qty': 'sum', 'Operator': 'count'}).rename(columns={'Operator': 'cnt'}).reset_index()
                ib_puts = grp.to_dict('records')
        except Exception as e_i:
            print(f"Lỗi query_actual_putaway_records IB: {e_i}")

    conn.close()
    return {'mti': mti_puts, 'ib': ib_puts}


def query_truck_tracker_records():
    """Lấy dữ liệu thông tin xe đến từ bảng [Daily truck tracker] trong SQLite"""
    conn = get_db_connection()
    c = conn.cursor()
    truck_data = {
        "columns": [],
        "items": [],
        "total": 0,
        "service_account": "nh-p-file-mti@rosy-cache-483805-c6.iam.gserviceaccount.com"
    }
    try:
        c.execute("""
            CREATE TABLE IF NOT EXISTS _truck_shipping_ids_overrides (
                trip_key TEXT PRIMARY KEY,
                date_val TEXT,
                chuyen TEXT,
                so_xe TEXT,
                shipping_id TEXT,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        conn.commit()

        c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='Daily truck tracker'")
        if c.fetchone()[0] > 0:
            df = pd.read_sql_query("SELECT rowid as _rowid, * FROM [Daily truck tracker]", conn)
            if not df.empty:
                c.execute("SELECT trip_key, date_val, chuyen, so_xe, shipping_id FROM _truck_shipping_ids_overrides WHERE shipping_id IS NOT NULL AND shipping_id != ''")
                overrides_map = {}
                so_xe_map = {}
                for ok, od, och, osx, osid in c.fetchall():
                    if ok:
                        overrides_map[ok.strip().lower()] = osid
                    if osx:
                        so_xe_map[osx.strip().lower()] = osid

                sid_col = 'Shipping ID' if 'Shipping ID' in df.columns else ('Shiping ID' if 'Shiping ID' in df.columns else None)
                if not sid_col:
                    df['Shipping ID'] = ''
                    sid_col = 'Shipping ID'

                date_col = next((col for col in ['Ngày MTO', 'Dự kiến giao hàng', 'Ngày', 'Ngay'] if col in df.columns), None)
                chuyen_col = next((col for col in ['Số chuyến', 'So chuyen', 'Chuyến', 'Chuyen'] if col in df.columns), None)
                so_xe_col = next((col for col in ['Số xe', 'So xe', 'Biển số'] if col in df.columns), None)

                for idx, row in df.iterrows():
                    cur_sid = str(row.get(sid_col, '') or '').strip()
                    if not cur_sid:
                        d_val = str(row.get(date_col, '') or '').strip() if date_col else ''
                        ch_val = str(row.get(chuyen_col, '') or '').strip() if chuyen_col else ''
                        sx_val = str(row.get(so_xe_col, '') or '').strip() if so_xe_col else ''
                        t_key = f"{d_val}_{ch_val}_{sx_val}".strip().lower()
                        if t_key in overrides_map:
                            df.at[idx, sid_col] = overrides_map[t_key]
                        elif sx_val and sx_val.lower() in so_xe_map:
                            df.at[idx, sid_col] = so_xe_map[sx_val.lower()]

                truck_data["columns"] = list(df.columns)
                truck_data["items"] = df.fillna('').to_dict(orient='records')
                truck_data["total"] = len(df)

        c.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='xe đến VNDL'")
        if c.fetchone()[0] > 0:
            df_xeden = pd.read_sql_query("SELECT rowid as _rowid, * FROM [xe đến VNDL]", conn)
            truck_data["xe_den_vndl"] = df_xeden.fillna('').to_dict(orient='records')
        else:
            truck_data["xe_den_vndl"] = []
    except Exception as e_trk:
        print(f"Lỗi query_truck_tracker_records: {e_trk}")
    conn.close()
    return truck_data


def query_data_mti_records():
    """Lấy dữ liệu từ bảng data_mti trong SQLite để hỗ trợ tra cứu nhanh MTP, MTB, SKU"""
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT date, pallet, mtb, sku, qty, kho_chuyen, ship_order, asn FROM data_mti ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()
        return [list(r) for r in rows]
    except Exception as e:
        print(f"Lỗi query data_mti: {e}")
        return []



def query_plan_mti():
    """Query [Data Plan MTI] from SQLite for plan volume per date/warehouse."""
    try:
        conn = get_db_connection()
        df = pd.read_sql_query("SELECT date, warehouse, qty FROM [Data Plan MTI]", conn)
        conn.close()
        records = []
        for _, row in df.iterrows():
            records.append({
                'date': str(row['date'] or ''),
                'warehouse': str(row['warehouse'] or ''),
                'qty': int(row['qty'] or 0)
            })
        return records
    except Exception as e:
        print(f"[plan_mti] Error: {e}")
        return []

def query_inbound_backlog():
    """Lấy dữ liệu tổng hợp và KPI cho Dashboard cả MTI, Inbound và RTS"""
    df_mti = compute_inbound_summary()
    items_mti = df_mti.to_dict('records') if not df_mti.empty else []
    for r in items_mti:
        r['_bucket'] = 'MTI'

    df_ib = compute_inbound_summary_ib()
    items_ib = df_ib.to_dict('records') if not df_ib.empty else []
    for r in items_ib:
        r['_bucket'] = 'Inbound'

    df_rts_o, df_rts_i = compute_inbound_summary_rts()
    rts_orders = df_rts_o.to_dict('records') if not df_rts_o.empty else []
    rts_items = df_rts_i.to_dict('records') if not df_rts_i.empty else []
    for r in rts_items:
        r['_bucket'] = 'RTS'

    overview_data = compute_overview_summary_data()
    actual_puts = query_actual_putaway_records()
    truck_tracker = query_truck_tracker_records()

    return {
        "updated_at": time.strftime("%H:%M:%S %d/%m/%Y"),
        "mti": {
            "kpis": calc_kpis(items_mti),
            "items": items_mti
        },
        "ib": {
            "kpis": calc_kpis(items_ib),
            "items": items_ib
        },
        "rts": {
            "kpis": calc_rts_kpis(rts_orders, rts_items),
            "orders": rts_orders,
            "items": rts_items
        },
        "overview_vndl": overview_data,
        "actual_putaway": actual_puts,
        "truck_tracker": truck_tracker,
        "data_mti": query_data_mti_records(),
        "plan_mti": query_plan_mti(),
        "kpis": calc_kpis(items_mti),
        "items": items_mti
    }


def generate_html_content(data=None, initial_bucket='MTI'):
    """Tạo mã nguồn HTML hoàn chỉnh cho Dashboard với hỗ trợ MTI, Inbound, RTS và ALL"""
    if data is None:
        try:
            data = query_inbound_backlog()
        except Exception:
            data = {
                "updated_at": time.strftime("%H:%M:%S %d/%m/%Y"),
                "mti": {"kpis": calc_kpis([]), "items": []},
                "ib": {"kpis": calc_kpis([]), "items": []},
                "rts": {"kpis": calc_rts_kpis([], []), "orders": [], "items": []},
                "kpis": calc_kpis([]),
                "items": []
            }

    data_json = json.dumps(data, ensure_ascii=False)

    search_placeholder = "Tìm mã đơn Ship Order, ASN..."
    if initial_bucket == 'RTS':
        search_placeholder = "Tìm Order Number, Request ID, Nhà cung cấp (Supplier)..."
    elif initial_bucket == 'Inbound':
        search_placeholder = "Tìm PO ID, ASN ID..."
    elif initial_bucket in ('Overview', 'Over view VNDL'):
        search_placeholder = "Tìm Zone, Loại kệ (RACK / PALLET)..."

    html = f"""<!DOCTYPE html>
<html lang="vi">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Inbound, MTI & RTS Backlog Dashboard</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel="stylesheet">
  <style>
    :root {{
      --bg-base: #0a0f1d;
      --bg-surface: #111827;
      --bg-surface-elevated: #1e293b;
      --bg-card: #151f32;
      --border-color: #243247;
      --border-subtle: #1e293b;
      --text-main: #f8fafc;
      --text-muted: #94a3b8;
      --text-dim: #64748b;
      --accent-blue: #38bdf8;
      --accent-green: #22c55e;
      --accent-rose: #f43f5e;
      --accent-amber: #f59e0b;
      --accent-purple: #a855f7;
    }}
    * {{
      margin: 0;
      padding: 0;
      box-sizing: border-box;
      font-family: 'Plus Jakarta Sans', -apple-system, sans-serif;
    }}
    html, body {{
      height: 100vh;
      overflow: hidden;
      background-color: var(--bg-base);
      color: var(--text-main);
    }}
    body {{
      display: flex;
      flex-direction: column;
      padding: 10px 16px 10px 16px;
      gap: 8px;
    }}
    .web-portal-nav {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      background-color: var(--bg-surface);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      padding: 8px 14px;
      flex-shrink: 0;
    }}
    .portal-left {{
      display: flex;
      align-items: center;
      gap: 12px;
    }}
    .header-title {{
      display: flex;
      align-items: center;
      gap: 10px;
    }}
    .header-title h1 {{
      font-size: 16px;
      font-weight: 800;
      color: #fff;
      letter-spacing: -0.3px;
    }}
    .badge-mti {{
      background: linear-gradient(135deg, #2563eb, #38bdf8);
      color: #fff;
      font-size: 10px;
      font-weight: 800;
      padding: 3px 8px;
      border-radius: 5px;
      text-transform: uppercase;
    }}
    .header-actions {{
      display: flex;
      align-items: center;
      gap: 10px;
    }}
    .live-badge {{
      display: flex;
      align-items: center;
      gap: 6px;
      font-size: 11px;
      color: var(--accent-green);
      font-weight: 700;
      background: rgba(34, 197, 94, 0.1);
      border: 1px solid rgba(34, 197, 94, 0.3);
      padding: 4px 10px;
      border-radius: 6px;
    }}
    .live-dot {{
      width: 7px;
      height: 7px;
      border-radius: 50%;
      background-color: var(--accent-green);
      animation: pulse 1.5s infinite;
    }}
    @keyframes pulse {{
      0% {{ transform: scale(0.95); opacity: 0.8; }}
      50% {{ transform: scale(1.3); opacity: 1; }}
      100% {{ transform: scale(0.95); opacity: 0.8; }}
    }}
    .btn-action {{
      background-color: #2563eb;
      color: white;
      border: none;
      padding: 5px 12px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
      display: flex;
      align-items: center;
      gap: 6px;
      transition: all 0.2s;
    }}
    .btn-action:hover {{
      background-color: #1d4ed8;
    }}
    .top-filter-bar {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 10px;
      background-color: var(--bg-surface);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      padding: 6px 12px;
      flex-shrink: 0;
    }}
    .filter-group-left {{
      display: flex;
      align-items: center;
      gap: 12px;
    }}
    .filter-control {{
      display: flex;
      align-items: center;
      gap: 6px;
    }}
    .filter-label {{
      font-size: 11px;
      font-weight: 800;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}
    .filter-select {{
      background-color: var(--bg-card);
      border: 1px solid #0284c7;
      color: #fff;
      font-size: 12px;
      font-weight: 700;
      padding: 4px 10px;
      border-radius: 6px;
      outline: none;
      cursor: pointer;
    }}
    .filter-date-input {{
      background-color: var(--bg-card);
      border: 1px solid #0284c7;
      color: #fff;
      font-size: 12px;
      font-weight: 700;
      padding: 4px 8px;
      border-radius: 6px;
      outline: none;
      cursor: pointer;
    }}
    .filter-date-input::-webkit-calendar-picker-indicator {{
      filter: invert(1);
      cursor: pointer;
    }}
    .btn-reset-date {{
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid var(--border-color);
      color: var(--text-muted);
      font-size: 11px;
      font-weight: 700;
      padding: 5px 10px;
      border-radius: 6px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 4px;
      transition: all 0.2s;
    }}
    .btn-reset-date:hover {{
      background: rgba(255, 255, 255, 0.12);
      color: #fff;
      border-color: #0284c7;
    }}
    .cal-trigger-wrapper {{
      position: relative;
    }}
    .cal-trigger-btn {{
      background-color: var(--bg-card);
      border: 1px solid #0284c7;
      color: #fff;
      font-size: 12px;
      font-weight: 700;
      padding: 4px 12px;
      border-radius: 6px;
      outline: none;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      gap: 6px;
    }}
    .cal-trigger-btn.open {{
      border-color: #38bdf8;
      box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.25);
    }}
    .active-filter-pill {{
      font-size: 11px;
      font-weight: 700;
      color: #38bdf8;
      background: rgba(56, 189, 248, 0.1);
      border: 1px solid rgba(56, 189, 248, 0.3);
      padding: 3px 10px;
      border-radius: 20px;
    }}
    /* POPUP LỊCH 2 THÁNG */
    .dual-cal-popover {{
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      z-index: 99999;
      background: #ffffff;
      color: #0f172a;
      border-radius: 12px;
      box-shadow: 0 20px 45px rgba(0, 0, 0, 0.55), 0 0 0 1px rgba(0, 0, 0, 0.1);
      width: 610px;
      display: none;
      flex-direction: column;
      overflow: hidden;
      animation: calSlideDown 0.18s ease-out;
    }}
    .dual-cal-popover.show {{
      display: flex;
    }}
    @keyframes calSlideDown {{
      from {{ opacity: 0; transform: translateY(-8px); }}
      to {{ opacity: 1; transform: translateY(0); }}
    }}
    .cal-quick-months {{
      display: flex;
      align-items: center;
      gap: 6px;
      padding: 8px 14px;
      background-color: #f8fafc;
      border-bottom: 1px solid #e2e8f0;
      overflow-x: auto;
    }}
    .cal-month-chip {{
      background-color: #ffffff;
      color: #475569;
      border: 1px solid #cbd5e1;
      border-radius: 6px;
      padding: 4px 10px;
      font-size: 11px;
      font-weight: 700;
      cursor: pointer;
      white-space: nowrap;
      transition: all 0.15s;
    }}
    .cal-month-chip:hover {{
      border-color: #ee4d2d;
      color: #ee4d2d;
    }}
    .cal-month-chip.active {{
      background-color: #ee4d2d;
      color: #ffffff;
      border-color: #ee4d2d;
    }}
    .cal-panels {{
      display: grid;
      grid-template-columns: 1fr 1fr;
      border-bottom: 1px solid #e2e8f0;
    }}
    .cal-panel {{
      padding: 12px 16px;
    }}
    .cal-panel:first-child {{
      border-right: 1px solid #e2e8f0;
    }}
    .cal-panel-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 10px;
      height: 28px;
    }}
    .cal-nav-group {{
      display: flex;
      gap: 4px;
      min-width: 48px;
    }}
    .cal-nav-btn {{
      background: none;
      border: none;
      color: #64748b;
      font-size: 15px;
      font-weight: 700;
      width: 24px;
      height: 24px;
      border-radius: 4px;
      cursor: pointer;
      display: inline-flex;
      align-items: center;
      justify-content: center;
    }}
    .cal-nav-btn:hover {{
      background-color: #f1f5f9;
      color: #0f172a;
    }}
    .cal-month-title {{
      font-size: 14px;
      font-weight: 700;
      color: #0f172a;
      cursor: pointer;
      padding: 3px 8px;
      border-radius: 5px;
      transition: background 0.15s;
    }}
    .cal-month-title:hover {{
      background-color: #fff1f0;
      color: #ee4d2d;
    }}
    .cal-weekdays {{
      display: grid;
      grid-template-columns: repeat(7, 1fr);
      text-align: center;
      font-size: 12px;
      font-weight: 600;
      color: #475569;
      margin-bottom: 6px;
    }}
    .cal-days-grid {{
      display: grid;
      grid-template-columns: repeat(7, 1fr);
      row-gap: 4px;
    }}
    .cal-day {{
      height: 34px;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      font-size: 13px;
      font-weight: 500;
      color: #1e293b;
      cursor: pointer;
      border-radius: 4px;
      position: relative;
      transition: all 0.12s;
    }}
    .cal-day:hover:not(.other-month) {{
      background-color: #fff1f0;
      color: #ee4d2d;
    }}
    .cal-day.other-month {{
      color: #94a3b8;
      cursor: default;
    }}
    .cal-day.is-today {{
      color: #ee4d2d;
      font-weight: 800;
    }}
    .cal-day.in-range {{
      background-color: #ffebe6;
      color: #ee4d2d;
      border-radius: 0;
    }}
    .cal-day.selected-start, .cal-day.selected-end {{
      background-color: #ee4d2d !important;
      color: #ffffff !important;
      font-weight: 700;
      border-radius: 4px;
    }}
    .cal-day-dot {{
      width: 4px;
      height: 4px;
      border-radius: 50%;
      background-color: #22c55e;
      position: absolute;
      bottom: 3px;
    }}
    .cal-day.selected-start .cal-day-dot, .cal-day.selected-end .cal-day-dot {{
      background-color: #ffffff;
    }}
    .cal-footer {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 10px 16px;
      background-color: #ffffff;
    }}
    .cal-footer-info {{
      font-size: 12px;
      color: #475569;
      font-weight: 600;
    }}
    .cal-footer-actions {{
      display: flex;
      align-items: center;
      gap: 8px;
    }}
    .btn-cal-clear {{
      background-color: #f1f5f9;
      color: #475569;
      border: 1px solid #cbd5e1;
      padding: 6px 12px;
      border-radius: 5px;
      font-size: 12px;
      font-weight: 600;
      cursor: pointer;
    }}
    .btn-cal-clear:hover {{
      background-color: #e2e8f0;
      color: #0f172a;
    }}
    .btn-cal-confirm {{
      background-color: #ee4d2d;
      color: #ffffff;
      border: none;
      padding: 6px 16px;
      border-radius: 5px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
      transition: background 0.15s;
    }}
    .btn-cal-confirm:hover {{
      background-color: #d73211;
    }}
    /* KPI GRID */
    .kpi-grid {{
      display: grid;
      grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
      gap: 10px;
      flex-shrink: 0;
    }}
    @media (max-width: 1400px) {{
      .kpi-grid {{
        grid-template-columns: repeat(3, 1fr);
      }}
    }}
    @media (max-width: 768px) {{
      .kpi-grid {{
        grid-template-columns: repeat(2, 1fr);
      }}
    }}
    .kpi-card {{
      background-color: var(--bg-surface);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      padding: 10px 14px;
      position: relative;
      overflow: hidden;
      transition: transform 0.15s, border-color 0.15s;
    }}
    .kpi-card::before {{
      content: '';
      position: absolute;
      top: 0;
      left: 0;
      width: 100%;
      height: 3px;
    }}
    .kpi-card.blue::before {{ background: var(--accent-blue); }}
    .kpi-card.purple::before {{ background: var(--accent-purple); }}
    .kpi-card.green::before {{ background: var(--accent-green); }}
    .kpi-card.rose::before {{ background: var(--accent-rose); }}
    .kpi-card.cyan::before {{ background: #06b6d4; }}
    .kpi-card.indigo::before {{ background: #6366f1; }}
    .kpi-card.indigo:hover {{
      border-color: #818cf8;
      transform: translateY(-2px);
    }}
    .kpi-card.cyan:hover {{
      border-color: #38bdf8;
      transform: translateY(-2px);
    }}
    .kpi-card.amber::before {{ background: #f59e0b; }}
    .kpi-card.amber:hover {{
      border-color: #f59e0b;
      transform: translateY(-2px);
    }}
    .kpi-title {{
      font-size: 11px;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.4px;
    }}
    .kpi-val {{
      font-size: 22px;
      font-weight: 800;
      color: #fff;
      margin: 4px 0 2px 0;
      font-family: 'JetBrains Mono', monospace;
    }}
    .kpi-sub {{
      font-size: 11px;
      color: var(--text-dim);
    }}
    /* PIPELINE */
    .pipeline-card {{
      background-color: var(--bg-surface);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      padding: 10px 14px;
      flex-shrink: 0;
    }}
    .pipeline-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      margin-bottom: 8px;
    }}
    .pipeline-header h3 {{
      font-size: 12px;
      font-weight: 800;
      color: #fff;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }}
    .pipeline-bar-wrapper {{
      height: 10px;
      background-color: #0f172a;
      border-radius: 6px;
      overflow: hidden;
      display: flex;
      margin-bottom: 8px;
    }}
    .pipeline-segment {{
      height: 100%;
      transition: width 0.3s;
    }}
    .seg-done {{ background-color: var(--accent-green); }}
    .seg-prog {{ background-color: var(--accent-amber); }}
    .seg-arriving {{ background-color: var(--accent-blue); }}
    .seg-pend {{ background-color: var(--accent-rose); }}
    .pipeline-stats {{
      display: grid;
      grid-template-columns: repeat(5, 1fr);
      gap: 6px;
    }}
    .stat-item {{
      background-color: var(--bg-card);
      border: 1px solid var(--border-subtle);
      border-radius: 6px;
      padding: 6px 10px;
      cursor: pointer;
      transition: all 0.2s;
    }}
    .stat-item:hover {{
      border-color: #0284c7;
      background-color: #1e293b;
    }}
    .stat-lbl {{
      font-size: 10px;
      font-weight: 700;
      color: var(--text-muted);
      text-transform: uppercase;
    }}
    .stat-val {{
      font-size: 16px;
      font-weight: 800;
      margin-top: 2px;
      font-family: 'JetBrains Mono', monospace;
    }}
    /* MAIN TABLE CARD */
    .main-table-card {{
      background-color: var(--bg-surface);
      border: 1px solid var(--border-color);
      border-radius: 10px;
      padding: 10px 14px;
      flex: 1;
      min-height: 0;
      display: flex;
      flex-direction: column;
    }}
    .table-toolbar {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      gap: 12px;
      margin-bottom: 8px;
      flex-shrink: 0;
    }}
    .search-box {{
      position: relative;
      flex: 1;
      max-width: 380px;
    }}
    .search-input {{
      width: 100%;
      background-color: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: 6px;
      padding: 6px 12px 6px 32px;
      color: #fff;
      font-size: 12px;
      outline: none;
    }}
    .search-input:focus {{
      border-color: #0284c7;
    }}
    .search-icon {{
      position: absolute;
      left: 10px;
      top: 50%;
      transform: translateY(-50%);
      font-size: 12px;
      color: var(--text-dim);
    }}
    .filter-tabs {{
      display: flex;
      gap: 6px;
      flex-wrap: wrap;
    }}
    .filter-tab {{
      background-color: var(--bg-card);
      color: var(--text-muted);
      border: 1px solid var(--border-subtle);
      border-radius: 6px;
      padding: 5px 10px;
      font-size: 11px;
      font-weight: 700;
      cursor: pointer;
      transition: all 0.2s;
    }}
    .filter-tab:hover {{
      color: #fff;
      background-color: #1e293b;
    }}
    .filter-tab.active {{
      background-color: #0284c7;
      color: #fff;
      border-color: #0284c7;
    }}
    /* SCROLLABLE TABLE */
    .table-scroll-container {{
      flex: 1;
      overflow-y: auto;
      overflow-x: auto;
      border: 1px solid var(--border-subtle);
      border-radius: 6px;
      background-color: var(--bg-card);
    }}
    .data-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
      text-align: left;
    }}
    .data-table thead {{
      position: sticky;
      top: 0;
      z-index: 10;
      background-color: #0f172a;
    }}
    .data-table th {{
      padding: 9px 12px;
      font-size: 10px;
      font-weight: 800;
      color: var(--text-muted);
      text-transform: uppercase;
      letter-spacing: 0.6px;
      border-bottom: 2px solid var(--border-color);
      white-space: nowrap;
    }}
    .data-table td {{
      padding: 8px 12px;
      border-bottom: 1px solid rgba(255, 255, 255, 0.05);
      white-space: nowrap;
      color: #cbd5e1;
    }}
    .data-table tr:hover td {{
      background-color: rgba(255, 255, 255, 0.03);
    }}
    .order-code {{
      font-family: 'JetBrains Mono', monospace;
      font-weight: 700;
      color: #fff;
    }}
    .status-badge {{
      display: inline-block;
      padding: 3px 8px;
      border-radius: 4px;
      font-size: 10px;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0.4px;
    }}
    .badge-pending {{
      background: rgba(244, 63, 94, 0.15);
      color: #fda4af;
      border: 1px solid rgba(244, 63, 94, 0.3);
    }}
    .badge-arriving {{
      background: rgba(56, 189, 248, 0.15);
      color: #7dd3fc;
      border: 1px solid rgba(56, 189, 248, 0.3);
    }}
    .badge-inprog {{
      background: rgba(245, 158, 11, 0.15);
      color: #fcd34d;
      border: 1px solid rgba(245, 158, 11, 0.3);
    }}
    .badge-done {{
      background: rgba(34, 197, 94, 0.15);
      color: #86efac;
      border: 1px solid rgba(34, 197, 94, 0.3);
    }}
    .badge-allocated {{
      background: rgba(168, 85, 247, 0.15);
      color: #d8b4fe;
      border: 1px solid rgba(168, 85, 247, 0.3);
    }}
    .badge-outbound {{
      background: rgba(34, 197, 94, 0.15);
      color: #86efac;
      border: 1px solid rgba(34, 197, 94, 0.3);
    }}
    .badge-pendingpick {{
      background: rgba(245, 158, 11, 0.15);
      color: #fcd34d;
      border: 1px solid rgba(245, 158, 11, 0.3);
    }}
    .arrived-select {{
      background-color: var(--bg-surface);
      border-radius: 6px;
      padding: 5px 12px;
      font-size: 12px;
      font-weight: 800;
      letter-spacing: 0.3px;
      outline: none;
      cursor: pointer;
      transition: all 0.2s ease;
    }}
    .arrived-select:hover {{
      filter: brightness(1.2);
    }}
    .arrived-select.yes {{
      border: 1.5px solid #16a34a;
      background: rgba(22, 163, 74, 0.25);
      color: #86efac;
    }}
    .arrived-select.no {{
      border: 1.5px solid #e11d48;
      background: rgba(225, 29, 72, 0.25);
      color: #fda4af;
    }}
    .arrived-select.done {{
      border: 1.5px solid #059669;
      background: rgba(5, 150, 105, 0.3);
      color: #34d399;
    }}
    /* RTS Status styles: sáng, rõ nét theo yêu cầu */
    .arrived-select.outbound {{
      border: 2px solid #22c55e;
      background: rgba(34, 197, 94, 0.28);
      color: #4ade80;
      font-weight: 800;
      text-shadow: 0 0 10px rgba(34, 197, 94, 0.5);
      box-shadow: 0 0 10px rgba(34, 197, 94, 0.25);
    }}
    .arrived-select.pendingpick {{
      border: 2px solid #ef4444;
      background: rgba(239, 68, 68, 0.28);
      color: #f87171;
      font-weight: 800;
      text-shadow: 0 0 10px rgba(239, 68, 68, 0.5);
      box-shadow: 0 0 10px rgba(239, 68, 68, 0.25);
    }}
    .arrived-select.allocated {{
      border: 2px solid #eab308;
      background: rgba(234, 179, 8, 0.28);
      color: #fde047;
      font-weight: 800;
      text-shadow: 0 0 10px rgba(234, 179, 8, 0.5);
      box-shadow: 0 0 10px rgba(234, 179, 8, 0.25);
    }}
    .arrived-select option {{
      background-color: #0f172a;
      color: #f8fafc;
      font-weight: 700;
      padding: 8px 12px;
    }}
    .arrived-select option[value="Allocated"] {{
      color: #fde047;
      background-color: #1e293b;
    }}
    .arrived-select option[value="Outbound"] {{
      color: #4ade80;
      background-color: #1e293b;
    }}
    .arrived-select option[value="Pending Pick"] {{
      color: #f87171;
      background-color: #1e293b;
    }}
    /* WAREHOUSE TRANSFER BADGES & COMPACT DROPDOWN */
    .whs-combobox {{
      position: relative;
      display: inline-block;
      width: 95px;
    }}
    .rts-whs-input {{
      width: 100%;
      height: 26px;
      border-radius: 6px;
      padding: 2px 6px;
      font-size: 11px;
      font-weight: 800;
      letter-spacing: 0.5px;
      text-align: center;
      outline: none;
      cursor: pointer;
      transition: all 0.2s ease;
      background: rgba(255, 255, 255, 0.05);
      border: 1px solid rgba(255, 255, 255, 0.15);
      color: #94a3b8;
    }}
    .rts-whs-input:hover {{
      filter: brightness(1.2);
    }}
    .rts-whs-input:focus {{
      border-color: #38bdf8;
      box-shadow: 0 0 0 2px rgba(56, 189, 248, 0.25);
    }}
    .rts-whs-input.whs-vnnl {{
      background: rgba(56, 189, 248, 0.18) !important;
      border: 1.5px solid #0284c7 !important;
      color: #38bdf8 !important;
    }}
    .rts-whs-input.whs-vna {{
      background: rgba(74, 222, 128, 0.18) !important;
      border: 1.5px solid #16a34a !important;
      color: #4ade80 !important;
    }}
    .rts-whs-input.whs-vnvl {{
      background: rgba(192, 132, 252, 0.18) !important;
      border: 1.5px solid #9333ea !important;
      color: #c084fc !important;
    }}
    .rts-whs-input.whs-vng {{
      background: rgba(251, 191, 36, 0.18) !important;
      border: 1.5px solid #d97706 !important;
      color: #fbbf24 !important;
    }}
    .rts-whs-input.whs-vnsl {{
      background: rgba(251, 113, 133, 0.18) !important;
      border: 1.5px solid #e11d48 !important;
      color: #fb7185 !important;
    }}
    .whs-dropdown-menu {{
      display: none;
      position: fixed;
      z-index: 99999;
      background: #0f172a;
      border: 1px solid rgba(255, 255, 255, 0.18);
      border-radius: 8px;
      box-shadow: 0 10px 25px rgba(0, 0, 0, 0.7);
      padding: 4px;
      width: 95px;
      backdrop-filter: blur(10px);
    }}
    .whs-opt-item {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      height: 22px;
      padding: 2px 8px;
      margin-bottom: 2px;
      border-radius: 4px;
      font-size: 11px;
      font-weight: 800;
      cursor: pointer;
      transition: all 0.15s ease;
      letter-spacing: 0.4px;
    }}
    .whs-opt-item:last-child {{
      margin-bottom: 0;
    }}
    .whs-opt-vnnl {{
      color: #38bdf8;
      background: rgba(56, 189, 248, 0.12);
      border: 1px solid rgba(56, 189, 248, 0.3);
    }}
    .whs-opt-vnnl:hover {{
      background: rgba(56, 189, 248, 0.3);
      transform: translateX(2px);
    }}
    .whs-opt-vna {{
      color: #4ade80;
      background: rgba(74, 222, 128, 0.12);
      border: 1px solid rgba(74, 222, 128, 0.3);
    }}
    .whs-opt-vna:hover {{
      background: rgba(74, 222, 128, 0.3);
      transform: translateX(2px);
    }}
    .whs-opt-vnvl {{
      color: #c084fc;
      background: rgba(192, 132, 252, 0.12);
      border: 1px solid rgba(192, 132, 252, 0.3);
    }}
    .whs-opt-vnvl:hover {{
      background: rgba(192, 132, 252, 0.3);
      transform: translateX(2px);
    }}
    .whs-opt-vng {{
      color: #fbbf24;
      background: rgba(251, 191, 36, 0.12);
      border: 1px solid rgba(251, 191, 36, 0.3);
    }}
    .whs-opt-vng:hover {{
      background: rgba(251, 191, 36, 0.3);
      transform: translateX(2px);
    }}
    .whs-opt-vnsl {{
      color: #fb7185;
      background: rgba(251, 113, 133, 0.12);
      border: 1px solid rgba(251, 113, 133, 0.3);
    }}
    .whs-opt-vnsl:hover {{
      background: rgba(251, 113, 133, 0.3);
      transform: translateX(2px);
    }}
    .whs-opt-clear {{
      color: #94a3b8;
      background: rgba(255, 255, 255, 0.04);
      border: 1px dashed rgba(255, 255, 255, 0.15);
      font-size: 10px;
      font-weight: 600;
      justify-content: center;
      margin-top: 3px;
    }}
    .whs-opt-clear:hover {{
      background: rgba(239, 68, 68, 0.15);
      color: #f87171;
      border-color: rgba(239, 68, 68, 0.3);
    }}
    .backlog-num.has-backlog {{
      color: var(--accent-rose);
      font-weight: 700;
    }}
    .backlog-num.zero {{
      color: var(--accent-green);
    }}
    .toast-msg {{
      position: fixed;
      bottom: 24px;
      right: 24px;
      background-color: #059669;
      color: white;
      padding: 10px 18px;
      border-radius: 8px;
      font-size: 12px;
      font-weight: 700;
      box-shadow: 0 10px 25px rgba(0,0,0,0.4);
      opacity: 0;
      transform: translateY(20px);
      transition: all 0.3s ease;
      pointer-events: none;
      z-index: 999999;
    }}
    .toast-msg.show {{
      opacity: 1;
      transform: translateY(0);
    }}
    .filter-date-btn {{
      display: flex;
      align-items: center;
      gap: 6px;
      background-color: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 6px;
      color: var(--text-color);
      padding: 6px 12px;
      font-size: 13px;
      font-weight: 600;
      cursor: pointer;
      user-select: none;
      transition: all 0.2s ease;
      min-width: 130px;
      justify-content: space-between;
    }}
    .filter-date-btn:hover {{
      border-color: var(--accent-blue);
      background: rgba(255, 255, 255, 0.08);
    }}
    .multi-date-popover {{
      display: none;
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      z-index: 99999;
      background: #0f172a;
      border: 1px solid rgba(255, 255, 255, 0.15);
      border-radius: 10px;
      box-shadow: 0 15px 35px rgba(0, 0, 0, 0.6);
      width: 260px;
      padding: 12px;
      user-select: none;
    }}
    .multi-date-popover.show {{
      display: block;
    }}
    .cal-header {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 8px;
    }}
    .cal-header-title {{
      display: flex;
      align-items: center;
      gap: 4px;
      font-size: 13px;
      font-weight: 700;
      color: #fff;
    }}
    .cal-nav-btn {{
      background: rgba(255, 255, 255, 0.06);
      border: 1px solid rgba(255, 255, 255, 0.12);
      border-radius: 4px;
      color: #cbd5e1;
      width: 24px;
      height: 24px;
      display: flex;
      align-items: center;
      justify-content: center;
      cursor: pointer;
      font-size: 10px;
      transition: background 0.15s ease;
    }}
    .cal-nav-btn:hover {{
      background: rgba(255, 255, 255, 0.18);
      color: #fff;
    }}
    .cal-grid-days {{
      display: grid;
      grid-template-columns: repeat(7, 1fr);
      text-align: center;
      font-size: 11px;
      font-weight: 700;
      color: #94a3b8;
      margin-bottom: 6px;
      gap: 2px;
    }}
    .cal-grid-dates, .cal-grid-cells {{
      display: grid;
      grid-template-columns: repeat(7, 1fr);
      gap: 3px;
      margin-bottom: 8px;
    }}
    .cal-date-cell, .cal-day-cell {{
      height: 28px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 12px;
      font-weight: 600;
      border-radius: 4px;
      cursor: pointer;
      color: #f1f5f9;
      background: transparent;
      transition: all 0.15s ease;
      user-select: none;
    }}
    .cal-date-cell:hover, .cal-day-cell:hover {{
      background: rgba(255, 255, 255, 0.12);
    }}
    .cal-date-cell.other-month, .cal-day-cell.other-month {{
      color: #475569;
      pointer-events: none;
    }}
    .cal-date-cell.has-data {{
      position: relative;
    }}
    .cal-date-cell.has-data::after {{
      content: '';
      position: absolute;
      bottom: 2px;
      width: 4px;
      height: 4px;
      border-radius: 50%;
      background: #38bdf8;
    }}
    .cal-date-cell.today, .cal-day-cell.today {{
      border: 1px solid #38bdf8;
      color: #38bdf8;
      font-weight: 800;
    }}
    .cal-date-cell.selected, .cal-day-cell.selected {{
      background: #0284c7 !important;
      color: #ffffff !important;
      font-weight: 800;
    }}
    .cal-date-cell.in-range {{
      background: rgba(2, 132, 199, 0.45) !important;
      color: #ffffff !important;
      font-weight: 600;
    }}
    .cal-date-cell.range-start {{
      background: #0284c7 !important;
      color: #ffffff !important;
      font-weight: 800;
      border-radius: 6px 0 0 6px;
    }}
    .cal-date-cell.range-end {{
      background: #0284c7 !important;
      color: #ffffff !important;
      font-weight: 800;
      border-radius: 0 6px 6px 0;
    }}
    .cal-footer {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding-top: 8px;
      border-top: 1px solid rgba(255, 255, 255, 0.08);
      font-size: 11px;
    }}
    .cal-action-btn, .cal-footer-btn {{
      background: none;
      border: none;
      color: #38bdf8;
      font-size: 11px;
      font-weight: 700;
      cursor: pointer;
      padding: 3px 8px;
      border-radius: 4px;
      transition: background 0.15s ease;
    }}
    .cal-action-btn:hover, .cal-footer-btn:hover {{
      background: rgba(56, 189, 248, 0.15);
    }}
  </style>
</head>
<body>
  <div class="web-portal-nav">
    <div class="portal-left">
      <div class="header-title">
        <span class="badge-mti" id="headerBadge">MTI</span>
        <h1 id="headerTitleText">Inbound MTI Backlog Dashboard</h1>
      </div>
    </div>
    <div class="header-actions">
      <div class="live-badge">
        <span class="live-dot"></span>
        <span id="live-status">ĐỒNG BỘ TRỰC TIẾP VỚI SQL</span>
      </div>
      <div style="font-size: 11px; color: var(--text-dim);" id="last-update">Cập nhật: {data.get('updated_at', '')}</div>
      <button class="btn-action" onclick="fetchLiveSQLData(true)">⚡ Làm Mới Ngay</button>
      <button class="btn-action" style="background-color: #059669;" onclick="exportCSV()">📥 Xuất CSV</button>
    </div>
  </div>

  <div class="top-filter-bar">
    <div class="filter-group-left">
      <div class="filter-control">
        <span class="filter-label">Bucket</span>
        <select id="bucketFilter" class="filter-select" onchange="onBucketChange(this.value)">
          <option value="MTI" {'selected' if initial_bucket == 'MTI' else ''}>MTI</option>
          <option value="Inbound" {'selected' if initial_bucket == 'Inbound' else ''}>Inbound</option>
          <option value="RTS" {'selected' if initial_bucket == 'RTS' else ''}>RTS</option>
          <option value="Overview" {'selected' if initial_bucket in ('Overview', 'Over view VNDL') else ''}>Over view VNDL</option>
        </select>
      </div>

      <div class="filter-control">
        <span class="filter-label">📅 Month</span>
        <select id="monthFilter" class="filter-select" onchange="onMonthChange(this.value)">
          <option value="ALL">All Months</option>
        </select>
      </div>

      <div class="filter-control" style="position: relative;">
        <span class="filter-label">📅 CHỌN LỊCH</span>
        <div id="multiDatePickerTrigger" class="filter-date-btn" onclick="toggleMultiDatePicker(event)" title="Click để chọn một hoặc nhiều ngày cùng lúc">
          <span id="multiDateBtnText">Tất cả ngày</span>
          <span style="font-size: 9px; opacity: 0.7;">▼</span>
        </div>

        <!-- POPUP LỊCH ĐA CHỌN NGÀY -->
        <div id="multiDatePopover" class="multi-date-popover" onclick="event.stopPropagation()">
          <div class="cal-header">
            <div class="cal-header-title">
              <span id="calMonthTitle" style="font-size: 13px; font-weight: 700; color: #fff;">November 2026</span>
              <span style="font-size: 9px; margin-left: 3px; opacity: 0.7;">▼</span>
            </div>
            <div style="display: flex; gap: 4px;">
              <button type="button" class="cal-nav-btn" onclick="navCalendarMonth(-1)" title="Tháng trước">▲</button>
              <button type="button" class="cal-nav-btn" onclick="navCalendarMonth(1)" title="Tháng sau">▼</button>
            </div>
          </div>
          <div class="cal-grid-days">
            <div>Su</div><div>Mo</div><div>Tu</div><div>We</div><div>Th</div><div>Fr</div><div>Sa</div>
          </div>
          <div class="cal-grid-dates" id="calGridDates"></div>
          <div class="cal-footer">
            <button type="button" class="cal-action-btn" onclick="clearSelectedDates()">Clear</button>
            <button type="button" class="cal-action-btn" onclick="selectTodayDate()">Today</button>
            <button type="button" class="cal-action-btn" style="color: #4ade80; font-weight: 800;" onclick="closeMultiDatePicker()">Xong</button>
          </div>
        </div>
      </div>

      <button type="button" class="btn-reset-date" onclick="resetDateMonthFilters()">🔄 All</button>
    </div>

    <div class="active-filter-pill" id="activeFilterLabel">📅 Tất cả ngày & tháng</div>
  </div>

  <div class="kpi-grid">
    <div class="kpi-card blue">
      <div class="kpi-title" id="kpi-orders-title">Tổng Ship Order</div>
      <div class="kpi-val" id="kpi-orders">-</div>
      <div class="kpi-sub" id="kpi-orders-sub">Tổng số đơn nhập</div>
    </div>
    <!-- PLAN VOLUME MTI CARD -->
    <div class="kpi-card indigo" id="kpi-card-plan-volume" onclick="openPlanVolumeModal()" style="cursor: pointer;" title="Click để xem bảng chi tiết số lượng Item/SKU theo Kho chuyển">
      <div class="kpi-title" id="kpi-plan-title">Plan Volume MTI</div>
      <div class="kpi-val" id="kpi-plan-volume" style="color: #818cf8;">-</div>
      <div class="kpi-sub" id="kpi-plan-sub">Kế hoạch từ data MTI • Click xem</div>
    </div>
    <div class="kpi-card purple" id="kpi-card-actual-volume" onclick="openActualVolumeModal()" style="cursor: pointer;" title="Click để xem chi tiết Actual Volume theo Kho chuyển">
      <div class="kpi-title" id="kpi-qty-title">Actual Volume MTI</div>
      <div class="kpi-val" id="kpi-qty">-</div>
      <div class="kpi-sub" id="kpi-qty-sub">Thực tế nhập theo SO</div>
    </div>
    <!-- GAP VOLUME MTI CARD -->
    <div class="kpi-card orange" id="kpi-card-gap-volume" onclick="openGapVolumeModal()" style="cursor: pointer; display: none;" title="Click để xem Gap giữa Plan và Actual theo từng Kho">
      <div class="kpi-title" id="kpi-gap-title">Gap Volume MTI</div>
      <div class="kpi-val" id="kpi-gap-volume" style="color: #fb923c;">-</div>
      <div class="kpi-sub" id="kpi-gap-sub">Plan vs Actual • Click xem chi tiết</div>
      <div id="kpi-gap-pct" style="margin-top: 6px; padding-top: 4px; border-top: 1px dashed rgba(255,255,255,0.14); font-size: 11px; font-weight: 700; color: #fb923c; text-align: center;">-% Gap</div>
    </div>
    <div class="kpi-card green">
      <div class="kpi-title" id="kpi-putaway-title">Đã Lên Kệ (Putaway)</div>
      <div class="kpi-val" id="kpi-putaway">-</div>
      <div class="kpi-sub" id="kpi-putaway-sub">Thực tế lên kệ</div>
    </div>
    <div class="kpi-card rose">
      <div class="kpi-title" id="kpi-backlog-title">Số Lượng Tồn Đọng (Backlog)</div>
      <div class="kpi-val" id="kpi-backlog">-</div>
      <div class="kpi-sub" id="kpi-backlog-sub">Backlog = Quantity - Putaway</div>
      <div id="kpi-backlog-breakdown" style="margin-top: 8px; padding-top: 6px; border-top: 1px solid rgba(255,255,255,0.1); font-size: 11px; display: flex; flex-direction: column; gap: 3px;">
        <div style="display: flex; justify-content: space-between; align-items: center;">
          <span style="color: #fb7185; font-weight: 700;">⏳ Pending (Xe chưa đến):</span>
          <span style="color: #fff; font-weight: 800;" id="kpi-backlog-pending">- QTY</span>
        </div>
        <div style="display: flex; justify-content: space-between; align-items: center;">
          <span style="color: #38bdf8; font-weight: 700;">🚚 Đã đến kho (Đang xử lý):</span>
          <span style="color: #fff; font-weight: 800;" id="kpi-backlog-arriving">- QTY</span>
        </div>
      </div>
    </div>
    <div class="kpi-card cyan" id="kpi-card-actual-put" onclick="openActualPutawayModal()" style="cursor: pointer;" title="Click để xem bảng năng suất theo Operator">
      <div class="kpi-title" id="kpi-actual-title">Actual Putaway</div>
      <div class="kpi-val" id="kpi-actual-putaway" style="color: #38bdf8;">-</div>
      <div class="kpi-sub" id="kpi-actual-sub">Theo ngày Putaway Time</div>
    </div>
    <div class="kpi-card amber" id="kpi-card-truck-tracker" onclick="openTruckTrackerModal()" style="cursor: pointer;" title="Click để tra cứu thông tin xe đến (Daily Truck Tracker)">
      <div class="kpi-title" id="kpi-truck-title">Daily truck tracker</div>
      <div class="kpi-val" id="kpi-truck-count" style="color: #f59e0b;">🚚 Tra cứu xe</div>
      <div class="kpi-sub" id="kpi-truck-sub">Thông tin xe đến</div>
      <div id="kpi-truck-arrival-status" style="margin-top: 6px; padding-top: 4px; border-top: 1px dashed rgba(255,255,255,0.14); font-size: 11px; font-weight: 700; display: flex; align-items: center; justify-content: center; gap: 6px;">
        <span id="kpi-truck-arrived-badge" style="color: #4ade80; cursor: pointer; text-decoration: underline;" onclick="event.stopPropagation(); openTruckTrackerModal('arrived')" title="Xem danh sách xe đã đến">0 xe đã đến</span>
        <span style="color: rgba(255,255,255,0.3);">/</span>
        <span id="kpi-truck-not-arrived-badge" style="color: #fb7185; cursor: pointer; text-decoration: underline;" onclick="event.stopPropagation(); openTruckTrackerModal('not_arrived')" title="Xem danh sách xe chưa đến">0 xe chưa đến</span>
      </div>
    </div>
  </div>

  <div class="pipeline-card" id="pipelineCard">
    <div class="pipeline-header">
      <h3 id="pipelineTitle">Tiến Độ Xử Lý Toàn Bộ Đơn Hàng (Processing Pipeline)</h3>
      <div style="font-size: 13px; font-weight: 700; color: var(--accent-green);" id="pipeline-rate">0% Hoàn Thành</div>
    </div>
    <div class="pipeline-bar-wrapper">
      <div id="bar-done" class="pipeline-segment seg-done" style="width: 0%" title="Done"></div>
      <div id="bar-prog" class="pipeline-segment seg-prog" style="width: 0%" title="In Progress"></div>
      <div id="bar-arriving" class="pipeline-segment seg-arriving" style="width: 0%" title="Arriving"></div>
      <div id="bar-pend" class="pipeline-segment seg-pend" style="width: 0%" title="Pending"></div>
    </div>
    <div class="pipeline-stats" id="pipelineStats">
      <div class="stat-item" onclick="handlePipelineStatClick(1)" title="Lọc tất cả">
        <div class="stat-lbl" id="pipe-lbl-1">1. TỔNG ASN</div>
        <div class="stat-val" style="color: var(--accent-purple);" id="stat-total-asn">-</div>
      </div>
      <div class="stat-item" onclick="handlePipelineStatClick(2)" title="Lọc trạng thái này">
        <div class="stat-lbl" id="pipe-lbl-2">2. PENDING (XE CHƯA ĐẾN)</div>
        <div class="stat-val" style="color: var(--accent-rose);" id="stat-pending">-</div>
      </div>
      <div class="stat-item" onclick="handlePipelineStatClick(3)" title="Lọc trạng thái này">
        <div class="stat-lbl" id="pipe-lbl-3">3. ARRIVING (ĐÃ ĐẾN)</div>
        <div class="stat-val" style="color: var(--accent-blue);" id="stat-arriving">-</div>
      </div>
      <div class="stat-item" onclick="handlePipelineStatClick(4)" title="Lọc trạng thái này">
        <div class="stat-lbl" id="pipe-lbl-4">4. ĐANG LÀM (IN PROGRESS)</div>
        <div class="stat-val" style="color: var(--accent-amber);" id="stat-in-prog">-</div>
      </div>
      <div class="stat-item" onclick="handlePipelineStatClick(5)" title="Lọc trạng thái này">
        <div class="stat-lbl" id="pipe-lbl-5">5. DONE</div>
        <div class="stat-val" style="color: var(--accent-green);" id="stat-done">-</div>
      </div>
    </div>
  </div>

  <div class="main-table-card">
    <div class="table-toolbar">
      <div class="search-box">
        <span class="search-icon">🔍</span>
        <input type="text" id="searchInput" class="search-input" placeholder="{search_placeholder}" oninput="handleSearch()">
      </div>
      <div class="search-box" id="dataMtiSearchBox" style="display: {'flex' if initial_bucket == 'MTI' else 'none'}; border-color: rgba(56, 189, 248, 0.45); background: rgba(56, 189, 248, 0.06); width: 260px;">
        <span class="search-icon" style="font-size: 14px;">📦</span>
        <input type="text" id="dataMtiSearchInput" class="search-input" placeholder="Tìm MTP, MTB, SKU (data_mti)..." oninput="handleDataMtiSearch(this.value)" style="color: #38bdf8;">
        <button id="clearDataMtiBtn" onclick="clearDataMtiSearch()" style="display: none; background: none; border: none; color: #94a3b8; cursor: pointer; padding: 0 8px; font-size: 13px;" title="Xóa tìm kiếm">✕</button>
      </div>
      <div class="filter-tabs" id="filterTabs">
        <!-- Render động theo Bucket -->
      </div>
    </div>

    <div id="overviewWidgetsContainer" style="display: none; padding: 20px; background: rgba(15, 23, 42, 0.4); border-bottom: 1px solid var(--border-color);"></div>

    <div class="table-scroll-container">
      <table class="data-table">
        <thead id="tableHead">
          <tr id="tableHeadRow">
            <!-- Render động theo Bucket -->
          </tr>
        </thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>
  </div>

  <!-- MODAL ACTUAL PUTAWAY OPERATOR STATS -->
  <div id="actualPutawayModal" onclick="closeActualPutawayModal(event)" style="display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.75); backdrop-filter: blur(4px); z-index: 9999; align-items: center; justify-content: center; padding: 20px;">
    <div onclick="event.stopPropagation()" style="background: #181d28; border: 1px solid rgba(255,255,255,0.12); border-radius: 14px; width: 100%; max-width: 520px; max-height: 85vh; display: flex; flex-direction: column; box-shadow: 0 20px 50px rgba(0,0,0,0.8);">
      <div style="padding: 16px 20px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center;">
        <div>
          <div style="font-size: 15px; font-weight: 800; color: #f1f5f9; display: flex; align-items: center; gap: 8px;">
            <span>🟦</span> Kho VNDL &bull; 🚚 Actual Putaway
          </div>
          <div style="font-size: 12px; color: var(--text-dim); margin-top: 2px;" id="modal-putaway-subtitle">
            Thống kê năng suất theo Operator
          </div>
        </div>
        <button onclick="closeActualPutawayModal()" style="background: none; border: none; font-size: 24px; color: var(--text-dim); cursor: pointer; padding: 0 8px; line-height: 1;">&times;</button>
      </div>
      <div style="padding: 12px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); display: flex; gap: 10px;">
        <input type="text" id="modalOperatorSearch" placeholder="Tìm theo tên/email operator..." oninput="renderModalOperators()" style="flex: 1; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.1); border-radius: 6px; padding: 8px 12px; color: #fff; font-size: 13px; outline: none;">
      </div>
      <div style="padding: 8px 20px; font-size: 11px; font-weight: 700; color: #a78bfa; display: flex; justify-content: space-between; border-bottom: 1px solid rgba(255,255,255,0.08); text-transform: uppercase;">
        <span>👤 OPERATOR</span>
        <span>QTY &nbsp;|&nbsp; COUNT</span>
      </div>
      <div id="modalOperatorList" style="flex: 1; overflow-y: auto; padding: 10px 20px; font-family: 'JetBrains Mono', monospace; font-size: 13px; display: flex; flex-direction: column; gap: 8px;">
      </div>
      <div style="padding: 12px 20px; border-top: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; font-size: 12px; background: rgba(255,255,255,0.02);">
        <span style="color: var(--text-dim);" id="modalTotalOps">0 operators</span>
        <span style="color: #4ade80; font-weight: 800;" id="modalTotalQty">Total: 0 QTY</span>
      </div>
    </div>
  </div>



  <!-- MODAL ACTUAL VOLUME MTI (CHI TIẾT THEO KHO CHUYỂN) -->
  <div id="actualVolumeModal" onclick="closeActualVolumeModal(event)" style="display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.75); backdrop-filter: blur(4px); z-index: 9999; align-items: center; justify-content: center; padding: 20px;">
    <div onclick="event.stopPropagation()" style="background: #181d28; border: 1px solid rgba(255,255,255,0.14); border-radius: 14px; width: 100%; max-width: 680px; max-height: 88vh; display: flex; flex-direction: column; box-shadow: 0 25px 60px rgba(0,0,0,0.9);">
      <div style="padding: 16px 20px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; background: rgba(255,255,255,0.02);">
        <div>
          <div style="font-size: 15px; font-weight: 800; color: #c084fc; display: flex; align-items: center; gap: 8px;">
            <span>📦</span> ACTUAL VOLUME MTI &bull; Thực Tế Nhập Theo Kho Chuyển
          </div>
          <div style="font-size: 12px; color: var(--text-dim); margin-top: 2px;" id="modal-actual-volume-subtitle">
            Dữ liệu thực tế từ SO data MTI
          </div>
        </div>
        <button onclick="closeActualVolumeModal()" style="background: none; border: none; font-size: 24px; color: var(--text-dim); cursor: pointer; padding: 0 8px; line-height: 1;">&times;</button>
      </div>
      <div style="padding: 10px 20px; font-size: 11px; font-weight: 700; color: #d8b4fe; display: grid; grid-template-columns: 50px 1fr 110px 120px 80px; border-bottom: 1px solid rgba(255,255,255,0.08); text-transform: uppercase; background: rgba(192, 132, 252, 0.05);">
        <span>STT</span>
        <span>🏭 KHO CHUYỂN</span>
        <span style="text-align: right;">📦 SỐ SKU</span>
        <span style="text-align: right;">📊 SỐ LƯỢNG (QTY)</span>
        <span style="text-align: right;">TỶ TRỌNG</span>
      </div>
      <div id="modalActualVolumeList" style="flex: 1; overflow-y: auto; padding: 12px 20px; font-size: 13px; display: flex; flex-direction: column; gap: 8px;">
        <!-- Render động theo kho chuyển -->
      </div>
      <div style="padding: 14px 20px; border-top: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; font-size: 13px; background: rgba(255,255,255,0.02);">
        <span style="color: var(--text-dim);" id="modalActualTotalWh">0 kho chuyển</span>
        <div style="display: flex; gap: 16px;">
          <span style="color: #38bdf8; font-weight: 700;" id="modalActualTotalSku">0 SKU</span>
          <span style="color: #4ade80; font-weight: 800; font-family: 'JetBrains Mono', monospace;" id="modalActualTotalQty">Tổng: 0 QTY</span>
        </div>
      </div>
    </div>
  </div>

  <!-- MODAL GAP VOLUME MTI (PLAN VS ACTUAL) -->
  <div id="gapVolumeModal" onclick="closeGapVolumeModal(event)" style="display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.75); backdrop-filter: blur(4px); z-index: 9999; align-items: center; justify-content: center; padding: 20px;">
    <div style="background: #1e293b; border: 1px solid rgba(255,255,255,0.12); border-radius: 16px; width: 100%; max-width: 680px; max-height: 85vh; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 20px 60px rgba(0,0,0,0.6);">
      <div style="padding: 18px 20px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; flex-shrink: 0;">
        <div>
          <div style="font-size: 15px; font-weight: 800; color: #fb923c; display: flex; align-items: center; gap: 8px;">
            <span>📊</span> GAP VOLUME MTI &bull; Plan vs Actual
          </div>
          <div id="modal-gap-volume-subtitle" style="font-size: 12px; color: var(--text-dim); margin-top: 2px;">So sánh Plan và Actual theo Kho chuyển</div>
        </div>
        <button onclick="closeGapVolumeModal()" style="background: none; border: none; font-size: 24px; color: var(--text-dim); cursor: pointer; padding: 0 8px; line-height: 1;">&times;</button>
      </div>
      <div style="padding: 12px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); display: flex; gap: 16px; flex-shrink: 0; flex-wrap: wrap;">
        <div style="display: flex; gap: 6px; align-items: center; font-size: 12px;">
          <span style="color: var(--text-dim);">📦 Tổng Kho:</span>
          <strong id="modalGapTotalWh" style="color: #fb923c;">-</strong>
        </div>
        <div style="display: flex; gap: 6px; align-items: center; font-size: 12px;">
          <span style="color: var(--text-dim);">📋 Plan:</span>
          <strong id="modalGapTotalPlan" style="color: #818cf8;">-</strong>
        </div>
        <div style="display: flex; gap: 6px; align-items: center; font-size: 12px;">
          <span style="color: var(--text-dim);">✅ Actual:</span>
          <strong id="modalGapTotalActual" style="color: #4ade80;">-</strong>
        </div>
        <div style="display: flex; gap: 6px; align-items: center; font-size: 12px;">
          <span style="color: var(--text-dim);">⚡ Gap:</span>
          <strong id="modalGapTotalGap" style="color: #fb923c;">-</strong>
        </div>
      </div>
      <div style="padding: 8px 20px; border-bottom: 1px solid rgba(255,255,255,0.06); flex-shrink: 0;">
        <div style="display: grid; grid-template-columns: 40px 1fr 110px 110px 110px 80px; align-items: center; padding: 6px 10px; font-size: 11px; font-weight: 700; color: var(--text-dim); text-transform: uppercase; letter-spacing: 0.5px;">
          <span>STT</span>
          <span>🏢 KHO CHUYỂN</span>
          <span style="text-align: right;">Plan</span>
          <span style="text-align: right;">Actual</span>
          <span style="text-align: right;">Gap</span>
          <span style="text-align: right;">% Gap</span>
        </div>
      </div>
      <div id="modalGapVolumeList" style="flex: 1; overflow-y: auto; padding: 12px 20px; font-size: 13px; display: flex; flex-direction: column; gap: 8px;">
        <div style="text-align: center; padding: 30px; color: var(--text-dim);">Đang tải...</div>
      </div>
    </div>
  </div>

  <!-- MODAL PLAN VOLUME MTI (KẾ HOẠCH THEO NGÀY & KHO CHUYỂN) -->
  <div id="planVolumeModal" onclick="closePlanVolumeModal(event)" style="display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.75); backdrop-filter: blur(4px); z-index: 9999; align-items: center; justify-content: center; padding: 20px;">
    <div onclick="event.stopPropagation()" style="background: #181d28; border: 1px solid rgba(255,255,255,0.14); border-radius: 14px; width: 100%; max-width: 820px; max-height: 88vh; display: flex; flex-direction: column; box-shadow: 0 25px 60px rgba(0,0,0,0.9); overflow: hidden; font-family: 'JetBrains Mono', 'Segoe UI', monospace;">
      
      <!-- TOP HEADER -->
      <div style="padding: 14px 20px 10px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: flex-start; background: rgba(255,255,255,0.02);">
        <div>
          <div style="font-size: 15px; font-weight: 800; color: #818cf8; display: flex; align-items: center; gap: 8px;">
            <span>📦</span> PLAN VOLUME MTI &bull; KẾ HOẠCH THEO NGÀY &amp; KHO CHUYỂN
          </div>
          <div style="font-size: 12px; color: var(--text-dim); margin-top: 3px; display: flex; align-items: center; gap: 8px;" id="modal-plan-volume-subtitle">
            Kế hoạch từ sheet Data Plan MTI
          </div>
        </div>
        <button onclick="closePlanVolumeModal()" style="background: none; border: none; font-size: 24px; color: var(--text-dim); cursor: pointer; padding: 0 8px; line-height: 1;">&times;</button>
      </div>

      <!-- TOOLBAR: TABS & FILTER -->
      <div style="padding: 10px 20px; background: rgba(0,0,0,0.3); border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; flex-wrap: wrap; gap: 10px; justify-content: space-between; align-items: center;">
        
        <!-- VIEW TABS -->
        <div style="display: inline-flex; gap: 4px; background: rgba(255,255,255,0.04); padding: 3px; border-radius: 8px; border: 1px solid rgba(255,255,255,0.1);">
          <button type="button" id="planTabWh" onclick="setPlanModalViewMode('wh')" style="font-size: 11px; font-weight: 800; padding: 5px 12px; border-radius: 6px; cursor: pointer; border: none; background: #6366f1; color: #fff;">
            🏢 Theo Kho
          </button>
          <button type="button" id="planTabDaily" onclick="setPlanModalViewMode('daily')" style="font-size: 11px; font-weight: 800; padding: 5px 12px; border-radius: 6px; cursor: pointer; border: none; background: transparent; color: #94a3b8;">
            📅 Theo Từng Ngày
          </button>
          <button type="button" id="planTabMonthly" onclick="setPlanModalViewMode('monthly')" style="font-size: 11px; font-weight: 800; padding: 5px 12px; border-radius: 6px; cursor: pointer; border: none; background: transparent; color: #94a3b8;">
            🗓️ Theo Tháng
          </button>
        </div>

        <!-- QUICK FILTERS -->
        <div style="display: flex; align-items: center; gap: 8px; font-size: 12px; flex-wrap: wrap;">
          <select id="planModalMonthSelect" onchange="onPlanModalFilterChange()" style="background: #0f172a; border: 1px solid rgba(255,255,255,0.2); border-radius: 6px; color: #cbd5e1; padding: 4px 8px; font-size: 11px; outline: none; cursor: pointer;">
            <option value="ALL">🗓️ Tất cả tháng</option>
            <option value="10/2026">Tháng 10/2026</option>
            <option value="11/2026">Tháng 11/2026</option>
            <option value="12/2026">Tháng 12/2026</option>
          </select>

          <select id="planModalDateSelect" onchange="onPlanModalFilterChange()" style="background: #0f172a; border: 1px solid rgba(255,255,255,0.2); border-radius: 6px; color: #cbd5e1; padding: 4px 8px; font-size: 11px; outline: none; cursor: pointer; max-width: 155px;">
            <option value="ALL">📅 Tất cả ngày</option>
          </select>

          <button type="button" onclick="setPlanModalToday()" style="background: rgba(34, 197, 94, 0.15); border: 1px solid #4ade80; color: #4ade80; font-size: 11px; font-weight: 700; padding: 4px 9px; border-radius: 6px; cursor: pointer;" title="Xem kế hoạch hôm nay">
            ⚡ Hôm nay
          </button>

          <button type="button" onclick="resetPlanModalFilter()" style="background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.15); color: #cbd5e1; font-size: 11px; font-weight: 700; padding: 4px 9px; border-radius: 6px; cursor: pointer;" title="Xem tất cả">
            🔄 Tất cả
          </button>
        </div>
      </div>

      <!-- TABLE HEADER -->
      <div id="modalPlanTableHeader" style="padding: 10px 20px; font-size: 11px; font-weight: 700; color: #a5b4fc; display: grid; grid-template-columns: 50px 1fr 140px 120px 80px; border-bottom: 1px solid rgba(255,255,255,0.08); text-transform: uppercase; background: rgba(99, 102, 241, 0.05);">
        <span>STT</span>
        <span id="colHeaderName">🏭 KHO CHUYỂN</span>
        <span id="colHeaderExtra" style="text-align: right;">📦 CHI TIẾT</span>
        <span style="text-align: right;">📊 SỐ LƯỢNG (QTY)</span>
        <span style="text-align: right;">TỶ TRỌNG</span>
      </div>

      <!-- LIST ITEMS -->
      <div id="modalPlanVolumeList" style="flex: 1; overflow-y: auto; padding: 12px 20px; font-size: 13px; display: flex; flex-direction: column; gap: 8px;">
        <!-- Render động theo chế độ xem -->
      </div>

      <!-- FOOTER -->
      <div style="padding: 12px 20px; border-top: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; font-size: 13px; background: rgba(255,255,255,0.02);">
        <span style="color: var(--text-dim);" id="modalPlanTotalWh">0 kho chuyển</span>
        <div style="display: flex; gap: 16px;">
          <span style="color: #38bdf8; font-weight: 700;" id="modalPlanTotalSku">0 SKU</span>
          <span style="color: #4ade80; font-weight: 800; font-family: 'JetBrains Mono', monospace;" id="modalPlanTotalQty">Tổng: 0 QTY</span>
        </div>
      </div>
    </div>
  </div>

  <!-- MODAL DAILY TRUCK TRACKER (THÔNG TIN XE ĐẾN) -->
  <div id="truckTrackerModal" onclick="closeTruckTrackerModal(event)" style="display: none; position: fixed; inset: 0; background: rgba(0,0,0,0.75); backdrop-filter: blur(4px); z-index: 9999; align-items: center; justify-content: center; padding: 20px;">
    <div onclick="event.stopPropagation()" style="background: #181d28; border: 1px solid rgba(255,255,255,0.14); border-radius: 14px; width: 100%; max-width: 780px; max-height: 88vh; display: flex; flex-direction: column; box-shadow: 0 25px 60px rgba(0,0,0,0.9); font-family: 'JetBrains Mono', 'Segoe UI', monospace; overflow: hidden;">
      
      <!-- TOP TITLE BAR -->
      <div style="padding: 16px 20px 12px; border-bottom: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: flex-start; background: rgba(255,255,255,0.02);">
        <div>
          <div style="font-size: 14px; font-weight: 800; color: #38bdf8; letter-spacing: 0.5px; display: flex; align-items: center; gap: 4px;">
            <span>🟦🟦🟦</span> KHO VNDL <span>🟦🟦🟦</span>
          </div>
          <div style="font-size: 15px; font-weight: 800; color: #f8fafc; margin-top: 4px; display: flex; align-items: center; gap: 6px;">
            <span>🚚</span> DAILY TRUCK TRACKER &bull; XE ĐẾN VNDL
          </div>
        </div>
        <div style="display: flex; align-items: center; gap: 8px;">
          <button onclick="copyTruckTrackerText()" title="Sao chép toàn bộ nội dung" style="background: rgba(255,255,255,0.08); border: 1px solid rgba(255,255,255,0.15); border-radius: 6px; color: #cbd5e1; padding: 4px 10px; cursor: pointer; font-size: 12px; display: flex; align-items: center; gap: 4px;">
            📋 Copy
          </button>
          <button onclick="closeTruckTrackerModal()" style="background: none; border: none; font-size: 24px; color: var(--text-dim); cursor: pointer; padding: 0 4px; line-height: 1;">&times;</button>
        </div>
      </div>

      <!-- DATE & TOTAL STRIP -->
      <div style="padding: 10px 20px; border-bottom: 1px solid rgba(255,255,255,0.08); background: rgba(0,0,0,0.25); font-size: 13px; line-height: 1.6; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
        <div style="display: flex; align-items: center; gap: 8px; color: #f1f5f9; font-weight: 700; flex-wrap: wrap;">
          <span>📌</span> <span>NGÀY:</span> <span id="truckDateDisplay" style="color: #38bdf8;">-</span>
          <div style="display: inline-flex; align-items: center; gap: 4px; margin-left: 6px;">
            <button type="button" onclick="changeTruckModalDate(-1)" style="background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.15); color: #cbd5e1; font-size: 11px; padding: 2px 7px; border-radius: 4px; cursor: pointer;" title="Lùi 1 ngày">◀ Hôm trước</button>
            <button type="button" onclick="setTruckModalDateToday()" style="background: rgba(34,197,94,0.15); border: 1px solid #4ade80; color: #4ade80; font-size: 11px; font-weight: 700; padding: 2px 7px; border-radius: 4px; cursor: pointer;" title="Về hôm nay">⚡ Hôm nay</button>
            <button type="button" onclick="changeTruckModalDate(1)" style="background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.15); color: #cbd5e1; font-size: 11px; padding: 2px 7px; border-radius: 4px; cursor: pointer;" title="Tiến 1 ngày">Hôm sau ▶</button>
          </div>
          <span style="color: rgba(255,255,255,0.3); margin: 0 4px;">|</span>
          <span>📦</span> <span>TỔNG CHUYẾN:</span> <span id="truckTotalDisplay" style="color: #f59e0b;">0</span>
        </div>
        <div id="truckArrivedSummary" style="font-size: 12px; font-weight: 700;">
          <!-- Tóm tắt Đã đến / Chưa đến -->
        </div>
      </div>

      <!-- TABS LỌC XE: TẤT CẢ / XE ĐÃ ĐẾN / XE CHƯA ĐẾN / NHẬP XE ĐẾN -->
      <div id="truckModalTabs" style="display: flex; gap: 6px; padding: 8px 20px; border-bottom: 1px solid rgba(255,255,255,0.08); background: rgba(255,255,255,0.02); overflow-x: auto;">
        <!-- Render động qua renderTruckTrackerModal -->
      </div>

      <!-- TRIPS CONTENT LIST -->
      <div id="truckTripsList" style="flex: 1; overflow-y: auto; padding: 16px 20px; font-size: 13px; line-height: 1.8; color: #e2e8f0; display: flex; flex-direction: column; gap: 14px;">
        <!-- Render động từng chuyến xe -->
      </div>

      <!-- FORM NHẬP XE ĐẾN VNDL -->
      <div id="truckManualForm" style="display: none; padding: 16px 20px; overflow-y: auto; flex: 1;">
        <div style="font-size: 14px; font-weight: 800; color: #38bdf8; margin-bottom: 12px; display: flex; align-items: center; gap: 6px;">
          <span>📝</span> Ghi Nhận Xe Đến VNDL (Lưu về Google Sheet &amp; SQLite)
        </div>

        <!-- IN / OUT TOGGLE -->
        <div style="display: flex; gap: 8px; margin-bottom: 12px;">
          <button type="button" id="btn_in_toggle" onclick="setXeDenMode('IN')"
            style="flex:1; background: #16a34a; border: 2px solid #4ade80; border-radius: 8px; color: #fff; padding: 8px; font-weight: 800; font-size: 13px; cursor: pointer;">
            ✅ IN (Nhập kho)
          </button>
          <button type="button" id="btn_out_toggle" onclick="setXeDenMode('OUT')"
            style="flex:1; background: rgba(255,255,255,0.05); border: 1px solid rgba(255,255,255,0.2); border-radius: 8px; color: #94a3b8; padding: 8px; font-weight: 700; font-size: 13px; cursor: pointer;">
            📤 OUT (Xuất kho)
          </button>
          <input type="hidden" id="manual_inout_mode" value="IN">
        </div>

        <div style="display: flex; flex-direction: column; gap: 10px; font-size: 12px;">
          <!-- ROW 1: Biển số xe -->
          <div>
            <label style="color: #94a3b8; display: block; margin-bottom: 4px;">Biển số xe (*):</label>
            <input type="text" id="manual_so_xe" placeholder="VD: 29H-815.28" style="width: 100%; background: #0f172a; border: 1px solid rgba(255,255,255,0.18); border-radius: 6px; color: #fff; padding: 8px 10px; font-size: 13px; outline: none; box-sizing: border-box;">
          </div>
          <!-- ROW 2: SĐT + Tên tài xế -->
          <div style="display: flex; gap: 10px;">
            <div style="flex: 1;">
              <label style="color: #94a3b8; display: block; margin-bottom: 4px;">Số ĐT tài xế:</label>
              <input type="text" id="manual_sdt" placeholder="VD: 0987654321" style="width: 100%; background: #0f172a; border: 1px solid rgba(255,255,255,0.18); border-radius: 6px; color: #fff; padding: 8px 10px; font-size: 13px; outline: none; box-sizing: border-box;">
            </div>
            <div style="flex: 1;">
              <label style="color: #94a3b8; display: block; margin-bottom: 4px;">Tên tài xế:</label>
              <input type="text" id="manual_tai_xe" placeholder="Nguyễn Văn A" style="width: 100%; background: #0f172a; border: 1px solid rgba(255,255,255,0.18); border-radius: 6px; color: #fff; padding: 8px 10px; font-size: 13px; outline: none; box-sizing: border-box;">
            </div>
          </div>
          <!-- ROW 3: Pallet + Type -->
          <div style="display: flex; gap: 10px;">
            <div style="flex: 1;">
              <label style="color: #4ade80; display: block; margin-bottom: 4px; font-weight: 700;">Số Pallet (*):</label>
              <input type="number" id="manual_pallet" placeholder="VD: 10" min="0" style="width: 100%; background: #0f172a; border: 1px solid #16a34a; border-radius: 6px; color: #4ade80; font-weight: 700; padding: 8px 10px; font-size: 14px; outline: none; box-sizing: border-box;">
            </div>
            <div style="flex: 1;">
              <label style="color: #f59e0b; display: block; margin-bottom: 4px; font-weight: 700;">Type:</label>
              <select id="manual_type" style="width: 100%; background: #0f172a; border: 1px solid #d97706; border-radius: 6px; color: #f59e0b; font-weight: 700; padding: 8px 10px; font-size: 13px; outline: none; box-sizing: border-box; cursor: pointer;">
                <option value="">-- Chọn Type --</option>
                <option value="MTI" style="background:#166534; color:#fff;">MTI</option>
                <option value="MTO" style="background:#1e3a5f; color:#fff;">MTO</option>
                <option value="RTS" style="background:#4c1d95; color:#fff;">RTS</option>
                <option value="Liquidate" style="background:#7c2d12; color:#fff;">Liquidate</option>
                <option value="Transfer" style="background:#134e4a; color:#fff;">Transfer</option>
                <option value="IB" style="background:#1e3a5f; color:#fff;">IB</option>
              </select>
            </div>
          </div>
          <!-- ROW 4: Warehouse + Operater -->
          <div style="display: flex; gap: 10px;">
            <div style="flex: 1;">
              <label style="color: #38bdf8; display: block; margin-bottom: 4px; font-weight: 700;" id="label_warehouse">🏭 From WH (IN):</label>
              <input type="text" list="manual_wh_list" id="manual_warehouse" placeholder="VD: Thuận Thành / VNNL" style="width: 100%; background: #0f172a; border: 1px solid #0ea5e9; border-radius: 6px; color: #38bdf8; padding: 8px 10px; font-size: 13px; outline: none; box-sizing: border-box;">
              <datalist id="manual_wh_list">
                <option value="VNA">
                <option value="VNNL">
                <option value="VNG">
                <option value="VNVL">
                <option value="VNSL">
                <option value="VNS">
              </datalist>
            </div>
            <div style="flex: 1;">
              <label style="color: #94a3b8; display: block; margin-bottom: 4px;">Operater:</label>
              <select id="manual_operater" style="width: 100%; background: #0f172a; border: 1px solid rgba(255,255,255,0.18); border-radius: 6px; color: #e2e8f0; padding: 8px 10px; font-size: 13px; outline: none; box-sizing: border-box; cursor: pointer;">
                <option value="">-- Chọn Operater --</option>
                <option value="LÊ VĂN TÚ">LÊ VĂN TÚ</option>
                <option value="BÙI QUỐC PHÁP">BÙI QUỐC PHÁP</option>
                <option value="NGUYỄN THẾ CÔNG">NGUYỄN THẾ CÔNG</option>
                <option value="LÊ VIỆT HÙNG">LÊ VIỆT HÙNG</option>
                <option value="BÙI VĂN TUẤN">BÙI VĂN TUẤN</option>
                <option value="Huỳnh Vũ Tuyết Nhi">Huỳnh Vũ Tuyết Nhi</option>
              </select>
            </div>
          </div>
          <!-- ROW 5: Kho đi (Kho xuất phát) + Shipping ID -->
          <div style="display: flex; gap: 10px;">
            <div style="flex: 1;">
              <label style="color: #94a3b8; display: block; margin-bottom: 4px;">Kho xuất phát:</label>
              <input type="text" id="manual_kho_di" placeholder="Kho xuất phát" style="width: 100%; background: #0f172a; border: 1px solid rgba(255,255,255,0.18); border-radius: 6px; color: #fff; padding: 8px 10px; font-size: 13px; outline: none; box-sizing: border-box;">
            </div>
            <div style="flex: 1;">
              <label style="color: #a78bfa; display: block; margin-bottom: 4px; font-weight: 700;">Shipping ID:</label>
              <input type="text" id="manual_shipping_id" placeholder="VD: HDVNDL26..." style="width: 100%; background: #0f172a; border: 1px solid #7c3aed; border-radius: 6px; color: #a78bfa; padding: 8px 10px; font-size: 13px; outline: none; box-sizing: border-box;">
            </div>
          </div>
          <!-- ROW 6: Ghi chú -->
          <div>
            <label style="color: #94a3b8; display: block; margin-bottom: 4px;">Ghi chú / LDD:</label>
            <input type="text" id="manual_ghi_chu" placeholder="Ghi chú thêm..." style="width: 100%; background: #0f172a; border: 1px solid rgba(255,255,255,0.18); border-radius: 6px; color: #fff; padding: 8px 10px; font-size: 13px; outline: none; box-sizing: border-box;">
          </div>
          <!-- BUTTONS -->
          <div style="margin-top: 10px; display: flex; gap: 10px;">
            <button type="button" onclick="submitManualXeDen()" style="flex: 1; background: #16a34a; border: none; border-radius: 6px; color: #fff; padding: 10px; font-weight: 800; font-size: 13px; cursor: pointer;">💾 Lưu Xe Đến VNDL</button>
            <button type="button" onclick="filterTruckModal('all')" style="background: rgba(255,255,255,0.08); border: 1px solid rgba(255,255,255,0.15); border-radius: 6px; color: #cbd5e1; padding: 10px 16px; cursor: pointer;">Hủy</button>
          </div>
        </div>
      </div>

      <!-- FOOTER -->
      <div style="padding: 10px 20px; border-top: 1px solid rgba(255,255,255,0.08); display: flex; justify-content: space-between; align-items: center; font-size: 11px; color: var(--text-dim); background: rgba(255,255,255,0.02);">
        <span>Nguồn: SQLite [Daily truck tracker]</span>
        <button onclick="closeTruckTrackerModal()" style="background: rgba(255,255,255,0.08); border: 1px solid rgba(255,255,255,0.15); border-radius: 6px; color: #cbd5e1; padding: 4px 12px; cursor: pointer; font-size: 12px;">Đóng</button>
      </div>

    </div>
  </div>

  <div id="toast" class="toast-msg">Đã cập nhật vào SQL!</div>

  <script>
    let currentData = {data_json};
    let currentBucket = '{initial_bucket}';
    let currentFilter = 'ALL';
    let selectedMonth = 'ALL';
    let selectedDate = 'ALL';
    let selectedDates = [];
    let rangeStartDate = null;
    let rangeEndDate = null;
    let calViewYear = 2026;
    let calViewMonth = 11;
    let lastDataHash = '';
    let rtsViewMode = 'orders';

    function setRTSViewMode(mode) {{
      rtsViewMode = mode;
      renderFilterTabs();
      renderTableHead();
      renderTable();
    }}

    const MONTH_NAMES_EN = ['', 'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

    function parseReceiveDate(raw) {{
      const s = String(raw || '').trim();
      if (!s) return {{ ddmmyyyy: '', yyyymmdd: '', mmyyyy: '', sortKey: '' }};
      if (s.includes('/')) {{
        const parts = s.split('/');
        if (parts.length === 3) {{
          let d = parts[0].padStart(2, '0');
          let m = parts[1].padStart(2, '0');
          let y = parts[2].length === 2 ? '20' + parts[2] : parts[2];
          if (parts[0].length === 4) {{
            y = parts[0];
            m = parts[1].padStart(2, '0');
            d = parts[2].padStart(2, '0');
          }}
          return {{
            ddmmyyyy: `${{d}}/${{m}}/${{y}}`,
            yyyymmdd: `${{y}}-${{m}}-${{d}}`,
            mmyyyy: `${{m}}/${{y}}`,
            sortKey: `${{y}}${{m}}${{d}}`
          }};
        }}
      }}
      if (s.includes('-')) {{
        const parts = s.slice(0, 10).split('-');
        if (parts.length === 3) {{
          let y = parts[0];
          let m = parts[1].padStart(2, '0');
          let d = parts[2].padStart(2, '0');
          if (parts[2].length === 4) {{
            d = parts[0].padStart(2, '0');
            m = parts[1].padStart(2, '0');
            y = parts[2];
          }}
          return {{
            ddmmyyyy: `${{d}}/${{m}}/${{y}}`,
            yyyymmdd: `${{y}}-${{m}}-${{d}}`,
            mmyyyy: `${{m}}/${{y}}`,
            sortKey: `${{y}}${{m}}${{d}}`
          }};
        }}
      }}
      return {{ ddmmyyyy: s, yyyymmdd: '', mmyyyy: '', sortKey: s }};
    }}

    function getTruckPlanDate(dateStr) {{
      if (!dateStr || dateStr === 'ALL') return dateStr;
      const parts = dateStr.split('/');
      if (parts.length === 3) {{
        const d = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10) - 1;
        const y = parseInt(parts[2], 10);
        const dt = new Date(y, m, d);
        dt.setDate(dt.getDate() - 1);
        const prevD = String(dt.getDate()).padStart(2, '0');
        const prevM = String(dt.getMonth() + 1).padStart(2, '0');
        const prevY = dt.getFullYear();
        return `${{prevD}}/${{prevM}}/${{prevY}}`;
      }}
      return dateStr;
    }}

    function sortKeyToDDMMYYYY(k) {{
      if (!k || k.length !== 8) return k;
      return `${{k.slice(6, 8)}}/${{k.slice(4, 6)}}/${{k.slice(0, 4)}}`;
    }}

    function formatMonthLabel(mmyyyy) {{
      const parts = String(mmyyyy || '').split('/');
      if (parts.length === 2) {{
        const mNum = parseInt(parts[0], 10);
        const y = parts[1];
        const en = MONTH_NAMES_EN[mNum] || parts[0];
        return `${{en}} ${{y}}`;
      }}
      return mmyyyy;
    }}

    function populateDateMonthSelectors() {{
      const items = getRawBucketItems();
      const monthCounts = {{}};
      const dateCounts = {{}};
      const dateSortMap = {{}};

      const yearsSet = new Set();
      items.forEach(r => {{
        const raw = r['Receive Date'] || r['Create Time'] || '';
        const p = parseReceiveDate(raw);
        if (p.mmyyyy) {{
          monthCounts[p.mmyyyy] = (monthCounts[p.mmyyyy] || 0) + 1;
          const yNum = parseInt(p.mmyyyy.split('/')[1], 10);
          if (yNum) yearsSet.add(yNum);
        }}
        if (p.ddmmyyyy) {{
          if (selectedMonth === 'ALL' || p.mmyyyy === selectedMonth) {{
            dateCounts[p.ddmmyyyy] = (dateCounts[p.ddmmyyyy] || 0) + 1;
            dateSortMap[p.ddmmyyyy] = p.sortKey;
          }}
        }}
      }});

      if (yearsSet.size === 0) yearsSet.add(new Date().getFullYear());
      const sortedYears = Array.from(yearsSet).sort().reverse();

      // Populate Month select - Đủ 12 tháng bằng tiếng Anh
      const monthSel = document.getElementById('monthFilter');
      if (monthSel) {{
        const curM = selectedMonth;
        let mHtml = '<option value="ALL">All Months</option>';
        sortedYears.forEach(yr => {{
          for (let m = 1; m <= 12; m++) {{
            const mStr = String(m).padStart(2, '0');
            const mVal = `${{mStr}}/${{yr}}`;
            const mName = MONTH_NAMES_EN[m];
            const cnt = monthCounts[mVal] || 0;
            const countBadge = cnt > 0 ? ` (${{cnt}})` : '';
            const isSel = (curM === mVal) ? 'selected' : '';
            mHtml += `<option value="${{mVal}}" ${{isSel}}>${{mName}} ${{yr}}${{countBadge}}</option>`;
          }}
        }});
        monthSel.innerHTML = mHtml;
      }}
    }}

    function initCalendarState() {{
      const today = getTodayInfo();
      calViewYear = parseInt(today.sortKey.substring(0, 4));
      calViewMonth = parseInt(today.sortKey.substring(4, 6));
    }}

    function onMonthChange(val) {{
      selectedMonth = val;
      if (val !== 'ALL') {{
        const parts = val.split('/');
        if (parts.length === 2) {{
          calViewMonth = parseInt(parts[0]);
          calViewYear = parseInt(parts[1]);
        }}
      }}
      selectedDates = [];
      syncSelectedDateVar();
      updateMultiDateTriggerText();
      renderCalendar();
      updateUI();
    }}

    function toggleMultiDatePicker(e) {{
      if (e) e.stopPropagation();
      const pop = document.getElementById('multiDatePopover');
      if (pop) {{
        const isShow = pop.classList.contains('show');
        if (!isShow) {{
          renderCalendar();
          pop.classList.add('show');
        }} else {{
          pop.classList.remove('show');
        }}
      }}
    }}

    function closeMultiDatePicker() {{
      const pop = document.getElementById('multiDatePopover');
      if (pop) pop.classList.remove('show');
    }}

    function navCalendarMonth(delta) {{
      calViewMonth += delta;
      if (calViewMonth > 12) {{
        calViewMonth = 1;
        calViewYear++;
      }} else if (calViewMonth < 1) {{
        calViewMonth = 12;
        calViewYear--;
      }}
      renderCalendar();
    }}

    function parseDdmmyyyyToDate(s) {{
      if (!s) return null;
      const parts = s.split('/');
      if (parts.length !== 3) return null;
      const d = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10) - 1;
      const y = parseInt(parts[2], 10);
      return new Date(y, m, d);
    }}

    function formatDateToDdmmyyyy(dt) {{
      const d = String(dt.getDate()).padStart(2, '0');
      const m = String(dt.getMonth() + 1).padStart(2, '0');
      const y = dt.getFullYear();
      return `${{d}}/${{m}}/${{y}}`;
    }}

    function selectDateRange(d1Str, d2Str) {{
      const dt1 = parseDdmmyyyyToDate(d1Str);
      const dt2 = parseDdmmyyyyToDate(d2Str);
      if (!dt1 || !dt2) return [d1Str];
      const start = (dt1 <= dt2) ? dt1 : dt2;
      const end = (dt1 <= dt2) ? dt2 : dt1;
      const res = [];
      const cur = new Date(start);
      while (cur <= end) {{
        res.push(formatDateToDdmmyyyy(cur));
        cur.setDate(cur.getDate() + 1);
      }}
      return res;
    }}

    function handleCalendarDateClick(ddmmyyyy) {{
      if (!rangeStartDate || (rangeStartDate && rangeEndDate)) {{
        // Bắt đầu chọn khoảng ngày mới (Click lần 1)
        rangeStartDate = ddmmyyyy;
        rangeEndDate = null;
        selectedDates = [ddmmyyyy];
      }} else {{
        // Hoàn tất chọn khoảng ngày (Click lần 2: chọn toàn bộ ngày trong khoảng từ ngày 1 đến ngày 2)
        rangeEndDate = ddmmyyyy;
        selectedDates = selectDateRange(rangeStartDate, rangeEndDate);
        const dt1 = parseDdmmyyyyToDate(rangeStartDate);
        const dt2 = parseDdmmyyyyToDate(rangeEndDate);
        if (dt1 > dt2) {{
          const tmp = rangeStartDate;
          rangeStartDate = rangeEndDate;
          rangeEndDate = tmp;
        }}
      }}
      syncSelectedDateVar();
      renderCalendar();
      updateMultiDateTriggerText();
      updateUI();
    }}

    function clearSelectedDates() {{
      rangeStartDate = null;
      rangeEndDate = null;
      selectedDates = [];
      syncSelectedDateVar();
      renderCalendar();
      updateMultiDateTriggerText();
      updateUI();
      showToast('🔄 Đã xóa bộ lọc ngày!');
    }}

    function selectTodayDate() {{
      const today = getTodayInfo();
      rangeStartDate = today.ddmmyyyy;
      rangeEndDate = today.ddmmyyyy;
      selectedDates = [today.ddmmyyyy];
      calViewYear = parseInt(today.sortKey.substring(0, 4));
      calViewMonth = parseInt(today.sortKey.substring(4, 6));
      syncSelectedDateVar();
      renderCalendar();
      updateMultiDateTriggerText();
      updateUI();
      showToast(`📅 Đã chọn hôm nay: ${{today.ddmmyyyy}}`);
    }}

    function syncSelectedDateVar() {{
      if (selectedDates.length === 0) {{
        selectedDate = 'ALL';
      }} else if (selectedDates.length === 1) {{
        selectedDate = selectedDates[0];
      }} else {{
        selectedDate = selectedDates.join(', ');
      }}
    }}

    function updateMultiDateTriggerText() {{
      const btnText = document.getElementById('multiDateBtnText');
      const trigger = document.getElementById('multiDatePickerTrigger');
      if (!btnText) return;
      if (selectedDates.length === 0) {{
        if (currentBucket === 'Inbound') {{
          btnText.textContent = 'Đến ngày D';
          if (trigger) trigger.title = 'Mặc định Inbound: Tính từ đầu tháng đến ngày D (Hôm nay)';
        }} else {{
          btnText.textContent = 'Tất cả ngày';
          if (trigger) trigger.title = 'Click để chọn khoảng ngày';
        }}
      }} else if (rangeStartDate && rangeEndDate && rangeStartDate !== rangeEndDate) {{
        btnText.textContent = `${{rangeStartDate}} - ${{rangeEndDate}}`;
        if (trigger) trigger.title = `Khoảng ngày: ${{rangeStartDate}} đến ${{rangeEndDate}} (${{selectedDates.length}} ngày)`;
      }} else if (selectedDates.length === 1) {{
        btnText.textContent = selectedDates[0];
        if (trigger) trigger.title = `Ngày: ${{selectedDates[0]}}`;
      }} else {{
        btnText.textContent = `${{selectedDates.length}} ngày chọn`;
        if (trigger) trigger.title = `Đã chọn: ${{selectedDates.join(', ')}}`;
      }}
    }}

    function renderCalendar() {{
      const titleEl = document.getElementById('calMonthTitle');
      const gridEl = document.getElementById('calGridDates');
      if (!titleEl || !gridEl) return;

      const monthNames = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
      titleEl.textContent = `${{monthNames[calViewMonth]}} ${{calViewYear}}`;

      const dataDates = new Set();
      const rawItems = getRawBucketItems();
      rawItems.forEach(r => {{
        const rawDate = r['Receive Date'] || r['Create Time'] || '';
        const p = parseReceiveDate(rawDate);
        if (p.ddmmyyyy) dataDates.add(p.ddmmyyyy);
      }});

      const today = getTodayInfo();
      const daysInMonth = new Date(calViewYear, calViewMonth, 0).getDate();
      const firstDayIndex = new Date(calViewYear, calViewMonth - 1, 1).getDay();
      const prevMonthDays = new Date(calViewYear, calViewMonth - 1, 0).getDate();

      let cellsHtml = '';

      // Ngày tháng trước
      for (let i = firstDayIndex - 1; i >= 0; i--) {{
        const dNum = prevMonthDays - i;
        cellsHtml += `<div class="cal-date-cell other-month">${{dNum}}</div>`;
      }}

      // Các ngày trong tháng hiện tại
      for (let d = 1; d <= daysInMonth; d++) {{
        const ddStr = String(d).padStart(2, '0');
        const mmStr = String(calViewMonth).padStart(2, '0');
        const ddmmyyyy = `${{ddStr}}/${{mmStr}}/${{calViewYear}}`;

        const isSel = selectedDates.includes(ddmmyyyy);
        const isTod = (ddmmyyyy === today.ddmmyyyy);
        const hasD = dataDates.has(ddmmyyyy);

        const isRangeStart = (rangeStartDate === ddmmyyyy);
        const isRangeEnd = (rangeEndDate === ddmmyyyy);
        const isInRange = isSel && !isRangeStart && !isRangeEnd;

        let cls = 'cal-date-cell';
        if (isRangeStart && isRangeEnd) cls += ' selected';
        else if (isRangeStart) cls += ' selected range-start';
        else if (isRangeEnd) cls += ' selected range-end';
        else if (isInRange) cls += ' selected in-range';
        else if (isSel) cls += ' selected';

        if (isTod) cls += ' today';
        if (hasD) cls += ' has-data';

        cellsHtml += `<div class="${{cls}}" onclick="handleCalendarDateClick('${{ddmmyyyy}}')" title="${{ddmmyyyy}}">${{d}}</div>`;
      }}

      // Ngày tháng sau
      const totalCells = firstDayIndex + daysInMonth;
      const nextDays = (7 - (totalCells % 7)) % 7;
      for (let n = 1; n <= nextDays; n++) {{
        cellsHtml += `<div class="cal-date-cell other-month">${{n}}</div>`;
      }}

      gridEl.innerHTML = cellsHtml;
    }}

    function resetDateMonthFilters() {{
      selectedMonth = 'ALL';
      selectedDates = [];
      rangeStartDate = null;
      rangeEndDate = null;
      selectedDate = 'ALL';
      const mSel = document.getElementById('monthFilter');
      if (mSel) mSel.value = 'ALL';
      initCalendarState();
      updateMultiDateTriggerText();
      renderCalendar();
      populateDateMonthSelectors();
      updateUI();
      showToast('🔄 Đã reset: Hiển thị tất cả ngày & tháng!');
    }}

    function updateFilterLabels() {{
      const pill = document.getElementById('activeFilterLabel');
      let labelStr = 'All Dates & Months';
      if (rangeStartDate && rangeEndDate && rangeStartDate !== rangeEndDate) {{
        labelStr = `As of: ${{rangeStartDate}} - ${{rangeEndDate}} (${{selectedDates.length}} ngày)`;
      }} else if (selectedDates.length === 1) {{
        labelStr = `As of: ${{selectedDates[0]}}`;
      }} else if (selectedDates.length > 1) {{
        labelStr = `As of: ${{selectedDates.length}} ngày`;
      }} else if (selectedMonth !== 'ALL') {{
        labelStr = `Month: ${{formatMonthLabel(selectedMonth)}}`;
      }}
      if (pill) pill.textContent = '📅 ' + labelStr;
    }}

    function getRawBucketItems() {{
      if (currentBucket === 'RTS') {{
        if (rtsViewMode === 'items') {{
          return (currentData.rts && currentData.rts.items) ? currentData.rts.items : [];
        }}
        return (currentData.rts && currentData.rts.orders) ? currentData.rts.orders : [];
      }} else if (currentBucket === 'Inbound') {{
        return (currentData.ib && currentData.ib.items) ? currentData.ib.items : [];
      }} else if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
        return (currentData.overview_vndl && currentData.overview_vndl.items) ? currentData.overview_vndl.items : [];
      }} else {{
        return (currentData.mti && currentData.mti.items) ? currentData.mti.items : (currentData.items || []);
      }}
    }}

    function getTodayInfo() {{
      const now = new Date();
      const dd = String(now.getDate()).padStart(2, '0');
      const mm = String(now.getMonth() + 1).padStart(2, '0');
      const yyyy = String(now.getFullYear());
      return {{
        ddmmyyyy: `${{dd}}/${{mm}}/${{yyyy}}`,
        sortKey: `${{yyyy}}${{mm}}${{dd}}`
      }};
    }}

    function applyLocalArrivedOverrides(items) {{
      if (currentBucket === 'RTS') return;
      try {{
        const keyPrefix = (currentBucket === 'Inbound') ? 'ib_' : 'mti_';
        const saved = JSON.parse(localStorage.getItem(keyPrefix + 'arrived_overrides') || '{{}}');
        const savedDates = JSON.parse(localStorage.getItem(keyPrefix + 'arrived_dates') || '{{}}');
        const origDates = JSON.parse(localStorage.getItem(keyPrefix + 'orig_dates') || '{{}}');
        const today = getTodayInfo();

        items.forEach(r => {{
          const id = r['Ship Order'] || r['ASN ID'] || r['PO ID'];
          if (!origDates[id] && r['Receive Date']) {{
            origDates[id] = r['Receive Date'];
          }}
          const put = Number(r.Putaway || 0);
          const qty = Number(r.Quantity || 0);
          const backlog = Number(r.Backlog || 0);

          if (saved[id] === 'Done') {{
            r.Arrived = 'Done';
            r.Status = 'Done';
          }} else if (put > 0) {{
            if (saved[id] === 'No') {{
              saved[id] = 'Yes';
              savedDates[id] = today.ddmmyyyy;
              localStorage.setItem(keyPrefix + 'arrived_overrides', JSON.stringify(saved));
              localStorage.setItem(keyPrefix + 'arrived_dates', JSON.stringify(savedDates));
            }}
            if (backlog <= 700 || put >= qty) {{
              r.Status = 'Done';
              r.Arrived = 'Done';
            }} else {{
              r.Status = 'In Progress';
              r.Arrived = 'Yes';
            }}
          }} else if (saved[id]) {{
            r.Arrived = saved[id];
            if (saved[id] === 'Done') r.Status = 'Done';
            else r.Status = (r.Arrived === 'Yes') ? 'Arriving' : 'Pending';
          }}

          const isRowDone = (String(r.Status || '').trim().toLowerCase() === 'done' || String(r.Arrived || '').trim().toLowerCase() === 'done');
          if (isRowDone) {{
            r.Status = 'Done';
            r.Arrived = 'Done';
            r.Backlog = 0;
            if (currentBucket === 'MTI') {{
              r.Putaway = r.Quantity;
            }}
          }}

          // Chỉ cập nhật Receive Date khi người dùng đã chủ động bấm 'Yes xe tới' và có lưu ngày đến cụ thể
          if (savedDates[id]) {{
            r['Receive Date'] = savedDates[id];
          }} else if (origDates[id]) {{
            // Giữ nguyên ngày gốc từ sheet Data MTI, không tự ý nhảy sang ngày hôm nay
            r['Receive Date'] = origDates[id];
          }}
        }});
        localStorage.setItem(keyPrefix + 'orig_dates', JSON.stringify(origDates));
      }} catch (e) {{}}
    }}

    function onBucketChange(val) {{
      currentBucket = val;
      currentFilter = 'ALL';
      selectedMonth = 'ALL';
      selectedDates = [];
      rangeStartDate = null;
      rangeEndDate = null;
      selectedDate = 'ALL';

      const mSel = document.getElementById('monthFilter');
      if (mSel) mSel.value = 'ALL';
      updateMultiDateTriggerText();
      renderCalendar();

      const searchInput = document.getElementById('searchInput');
      if (searchInput) {{
        searchInput.value = '';
        if (val === 'RTS') {{
          searchInput.placeholder = 'Tìm Order Number, Request ID, Nhà cung cấp (Supplier)...';
        }} else if (val === 'Inbound') {{
          searchInput.placeholder = 'Tìm PO ID, ASN ID...';
        }} else if (val === 'Overview' || val === 'Over view VNDL') {{
          searchInput.placeholder = 'Tìm Zone, Loại kệ (RACK / PALLET)...';
        }} else {{
          searchInput.placeholder = 'Tìm mã đơn Ship Order, ASN...';
        }}
      }}

      const dataMtiBox = document.getElementById('dataMtiSearchBox');
      const dataMtiInp = document.getElementById('dataMtiSearchInput');
      if (dataMtiInp) dataMtiInp.value = '';
      isDataMtiActive = false;
      if (dataMtiBox) {{
        dataMtiBox.style.display = (val === 'MTI') ? 'flex' : 'none';
      }}

      const badge = document.getElementById('headerBadge');
      const title = document.getElementById('headerTitleText');
      const cardTitle1 = document.getElementById('kpi-orders-title');
      const cardSub1 = document.getElementById('kpi-orders-sub');
      const cardTitle2 = document.getElementById('kpi-qty-title');
      const cardSub2 = document.getElementById('kpi-qty-sub');
      const cardTitle3 = document.getElementById('kpi-putaway-title');
      const cardSub3 = document.getElementById('kpi-putaway-sub');
      const cardTitle4 = document.getElementById('kpi-backlog-title');
      const cardSub4 = document.getElementById('kpi-backlog-sub');
      const pipeTitle = document.getElementById('pipelineTitle');

      if (val === 'RTS') {{
        badge.textContent = 'RTS';
        badge.style.background = 'linear-gradient(135deg, #7c3aed, #a855f7)';
        title.textContent = 'RTS Backlog Dashboard';
        cardTitle1.textContent = 'Tổng Order RTS';
        cardSub1.textContent = 'Tổng số Order Number';
        cardTitle2.textContent = 'Tổng SKU';
        cardSub2.textContent = 'Số lượng mã SKU';
        cardTitle3.textContent = 'Tổng Số Lượng (QTY)';
        cardSub3.textContent = 'Tổng số lượng Allocated';
        cardTitle4.textContent = 'Đã Xuất (Outbound)';
        cardSub4.textContent = 'Hàng đã hoàn tất Outbound';
        pipeTitle.textContent = 'Trạng Thái Xử Lý Đơn Hàng RTS';
      }} else if (val === 'Inbound') {{
        badge.textContent = 'INBOUND';
        badge.style.background = 'linear-gradient(135deg, #059669, #10b981)';
        title.textContent = 'Inbound Backlog Dashboard';
        cardTitle1.textContent = 'Tổng PO / ASN';
        cardSub1.textContent = 'Tổng số đơn nhập từ data IB';
        cardTitle2.textContent = 'Tổng Số Lượng (Quantity)';
        cardSub2.textContent = 'Kế hoạch nhập từ data IB';
        cardTitle3.textContent = 'Đã Lên Kệ (Putaway)';
        cardSub3.textContent = 'Thực tế lên kệ từ Put IB';
        cardTitle4.textContent = 'Số Lượng Tồn Đọng (Backlog)';
        cardSub4.textContent = 'Backlog = Quantity - Putaway';
        pipeTitle.textContent = 'Tiến Độ Xử Lý Toàn Bộ Đơn Hàng (Processing Pipeline)';
      }} else if (val === 'Overview' || val === 'Over view VNDL') {{
        badge.textContent = 'OVERVIEW VNDL';
        badge.style.background = 'linear-gradient(135deg, #0284c7, #6366f1)';
        title.textContent = 'Over View VNDL (Dung Tích & Sức Chứa Kho)';
        cardTitle1.textContent = 'Hàng Tồn On-hands';
        cardSub1.textContent = 'Tổng số lượng tồn từ ONH';
        cardTitle2.textContent = 'Tổng Sức Chứa (Max CBM)';
        cardSub2.textContent = 'Max_capacity (m³)';
        cardTitle3.textContent = 'Dung Tích Đã Dùng';
        cardSub3.textContent = 'Volume Used (m³)';
        cardTitle4.textContent = 'Dung Tích Còn Trống';
        cardSub4.textContent = 'Empty_Capacity (m³)';
        pipeTitle.textContent = 'Tổng Quan Sức Chứa & Tỷ Lệ Sử Dụng Kho VNDL';
      }} else {{
        badge.textContent = 'MTI';
        badge.style.background = 'linear-gradient(135deg, #2563eb, #38bdf8)';
        title.textContent = 'Inbound MTI Backlog Dashboard';
        cardTitle1.textContent = 'Tổng Ship Order';
        cardSub1.textContent = 'Tổng số đơn nhập từ data_mti';
        cardTitle2.textContent = 'Tổng Số Lượng (Quantity)';
        cardSub2.textContent = 'Kế hoạch nhập từ data_mti';
        cardTitle3.textContent = 'Đã Lên Kệ (Putaway)';
        cardSub3.textContent = 'Thực tế lên kệ từ Put MTI';
        cardTitle4.textContent = 'Số Lượng Tồn Đọng (Backlog)';
        cardSub4.textContent = 'Backlog = Quantity - Putaway';
        pipeTitle.textContent = 'Tiến Độ Xử Lý Toàn Bộ Đơn Hàng (Processing Pipeline)';
      }}

      renderFilterTabs();
      renderTableHead();
      populateDateMonthSelectors();
      updateUI();
    }}

    function renderFilterTabs() {{
      const container = document.getElementById('filterTabs');
      if (currentBucket === 'RTS') {{
        const cntOrders = (currentData.rts && currentData.rts.orders) ? currentData.rts.orders.length : 0;
        const cntItems = (currentData.rts && currentData.rts.items) ? currentData.rts.items.length : 0;
        container.innerHTML = `
          <button class="filter-tab ${{currentFilter === 'ALL' ? 'active' : ''}}" data-filter="ALL" onclick="setFilter('ALL')">Tất Cả (<span id="cnt-rts-all">0</span>)</button>
          <button class="filter-tab ${{currentFilter === 'Allocated' ? 'active' : ''}}" style="color: #fde047; font-weight: 700;" data-filter="Allocated" onclick="setFilter('Allocated')">☑ Allocated (<span id="cnt-rts-alloc">0</span>)</button>
          <button class="filter-tab ${{currentFilter === 'Outbound' ? 'active' : ''}}" style="color: #4ade80; font-weight: 700;" data-filter="Outbound" onclick="setFilter('Outbound')">☑ Outbound (<span id="cnt-rts-out">0</span>)</button>
          <button class="filter-tab ${{currentFilter === 'Pending Pick' ? 'active' : ''}}" style="color: #f87171; font-weight: 700;" data-filter="Pending Pick" onclick="setFilter('Pending Pick')">☑ Pending Pick (<span id="cnt-rts-pick">0</span>)</button>
          <div style="margin-left: auto; display: flex; align-items: center; gap: 6px;">
            <span style="font-size: 11px; color: var(--text-dim); font-weight: 700; text-transform: uppercase;">Chế độ xem:</span>
            <button type="button" class="filter-tab ${{rtsViewMode === 'orders' ? 'active' : ''}}" onclick="setRTSViewMode('orders')" style="font-size: 11px; padding: 3px 10px; cursor: pointer;">📦 Theo Đơn (${{cntOrders}})</button>
            <button type="button" class="filter-tab ${{rtsViewMode === 'items' ? 'active' : ''}}" onclick="setRTSViewMode('items')" style="font-size: 11px; padding: 3px 10px; cursor: pointer;">📋 Chi Tiết Từng Dòng SKU (${{cntItems}})</button>
          </div>
        `;
      }} else if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
        container.innerHTML = `
          <button class="filter-tab ${{currentFilter === 'ALL' ? 'active' : ''}}" data-filter="ALL" onclick="setFilter('ALL')">Tất Cả Zone (<span id="cnt-ov-all">0</span>)</button>
          <button class="filter-tab ${{currentFilter === 'RACK' ? 'active' : ''}}" style="color: #38bdf8; font-weight: 700;" data-filter="RACK" onclick="setFilter('RACK')">Kệ Rack (<span id="cnt-ov-rack">0</span>)</button>
          <button class="filter-tab ${{currentFilter === 'PALLET' ? 'active' : ''}}" style="color: #f43f5e; font-weight: 700;" data-filter="PALLET" onclick="setFilter('PALLET')">Pallet (<span id="cnt-ov-pallet">0</span>)</button>
        `;
      }} else {{
        container.innerHTML = `
          <button class="filter-tab ${{currentFilter === 'ALL' ? 'active' : ''}}" data-filter="ALL" onclick="setFilter('ALL')">Tất Cả (<span id="cnt-all">0</span>)</button>
          <button class="filter-tab ${{currentFilter === 'Pending' ? 'active' : ''}}" data-filter="Pending" onclick="setFilter('Pending')">Pending - Xe Chưa Đến (<span id="cnt-pending">0</span>)</button>
          <button class="filter-tab ${{currentFilter === 'Arriving' ? 'active' : ''}}" data-filter="Arriving" onclick="setFilter('Arriving')">Arriving - Đã Đến (<span id="cnt-arriving">0</span>)</button>
          <button class="filter-tab ${{currentFilter === 'In Progress' ? 'active' : ''}}" data-filter="In Progress" onclick="setFilter('In Progress')">Đang Làm (<span id="cnt-inprog">0</span>)</button>
          <button class="filter-tab ${{currentFilter === 'Done' ? 'active' : ''}}" data-filter="Done" onclick="setFilter('Done')">Done (<span id="cnt-done">0</span>)</button>
        `;
      }}
    }}


    // Toggle hiển thị chi tiết SKU khi click vào dòng order RTS
    function toggleRTSSkuDetail(rowKey, orderNum, event) {{
      if (event && (event.target.tagName === 'SELECT' || event.target.tagName === 'INPUT' || event.target.tagName === 'BUTTON')) return;
      const detailRow = document.getElementById('rts_sku_row_' + rowKey);
      const arrow = document.getElementById('rts_arrow_' + rowKey);
      if (!detailRow) return;
      const isOpen = detailRow.style.display !== 'none';
      if (isOpen) {{
        detailRow.style.display = 'none';
        if (arrow) {{ arrow.style.transform = ''; arrow.textContent = '▶'; }}
        return;
      }}
      // Build SKU detail
      const allItems = (currentData.rts && currentData.rts.items) ? currentData.rts.items : [];
      const skus = allItems.filter(it => (it['Order Number'] || '') === orderNum);
      const detailDiv = document.getElementById('rts_sku_detail_' + rowKey);
      if (!detailDiv) return;
      if (skus.length === 0) {{
        detailDiv.innerHTML = '<span style="color:#94a3b8;font-size:12px;">Không có chi tiết SKU.</span>';
      }} else {{
        let skuHtml = `
          <table style="width:100%;border-collapse:collapse;font-size:12px;margin-top:4px;">
            <thead>
              <tr style="border-bottom:1px solid rgba(255,255,255,0.12);color:#94a3b8;">
                <th style="padding:5px 8px;text-align:left;font-weight:600;">SKU</th>
                <th style="padding:5px 8px;text-align:left;font-weight:600;">SKU Name</th>
                <th style="padding:5px 8px;text-align:center;font-weight:600;">QTY</th>
                <th style="padding:5px 8px;text-align:left;font-weight:600;">Status</th>
                <th style="padding:5px 8px;text-align:left;font-weight:600;">Ship Date</th>
              </tr>
            </thead>
            <tbody>
        `;
        skus.forEach(it => {{
          const st = it.Status || 'Allocated';
          const stColor = st === 'Outbound' ? '#4ade80' : st === 'Pending Pick' ? '#f87171' : '#fde047';
          skuHtml += `<tr style="border-bottom:1px solid rgba(255,255,255,0.05);">
            <td style="padding:4px 8px;color:#38bdf8;font-weight:700;">${{it['SKU'] || ''}}</td>
            <td style="padding:4px 8px;color:#e2e8f0;max-width:220px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${{it['SKU Name'] || ''}}">${{it['SKU Name'] || ''}}</td>
            <td style="padding:4px 8px;text-align:center;color:#4ade80;font-weight:700;">${{fmtNum(it['QTY'] || 0)}}</td>
            <td style="padding:4px 8px;color:${{stColor}};font-weight:600;">${{st}}</td>
            <td style="padding:4px 8px;color:#94a3b8;">${{it['Ship Date'] || ''}}</td>
          </tr>`;
        }});
        skuHtml += '</tbody></table>';
        detailDiv.innerHTML = skuHtml;
      }}
      detailRow.style.display = '';
      if (arrow) {{ arrow.style.transform = 'rotate(90deg)'; arrow.textContent = '▼'; }}
    }}

    function renderTableHead() {{
      const tr = document.getElementById('tableHeadRow');
      if (currentBucket === 'RTS') {{
        if (rtsViewMode === 'items') {{
          tr.innerHTML = `
            <th>ORDER NUMBER</th>
            <th>REQUEST ID</th>
            <th>MÃ NCC</th>
            <th>SKU</th>
            <th>SKU NAME</th>
            <th>QTY</th>
            <th>CREATE TIME</th>
            <th>SHIP DATE</th>
            <th>STATUS</th>
            <th>ĐỊA CHỈ CHUYỂN</th>
            <th>THÔNG TIN XE</th>
            <th>MAIL</th>
          `;
        }} else {{
          tr.innerHTML = `
            <th>ORDER NUMBER</th>
            <th>REQUEST ID</th>
            <th>MÃ NCC</th>
            <th>TOTAL SKU</th>
            <th>TOTAL QTY</th>
            <th>CREATE TIME</th>
            <th>SHIP DATE</th>
            <th>STATUS</th>
            <th>ĐỊA CHỈ CHUYỂN</th>
            <th>THÔNG TIN XE</th>
            <th>MAIL</th>
          `;
        }}
      }} else if (currentBucket === 'Inbound') {{
        tr.innerHTML = `
          <th>PO ID</th>
          <th>ASN ID</th>
          <th>Receive Date</th>
          <th>Expected Delivery</th>
          <th>Total SKU</th>
          <th>Quantity</th>
          <th>Putaway</th>
          <th>Backlog</th>
          <th>Status</th>
          <th>Arrived</th>
        `;
      }} else if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
        tr.innerHTML = `
          <th>Loại Kệ (Type)</th>
          <th>Zone</th>
          <th>Empty Qty</th>
          <th>Used Qty</th>
          <th>Location Qty</th>
          <th>Max_capacity (m³)</th>
          <th>Volume Used (m³)</th>
          <th>Empty_Capacity (m³)</th>
          <th>Qty Onhands</th>
          <th>AVG Items</th>
          <th>% Lấp Đầy</th>
        `;
      }} else {{
        tr.innerHTML = `
          <th>Ship Order</th>
          <th>ASN</th>
          <th>Receive Date</th>
          <th>Total SKU</th>
          <th>Quantity</th>
          <th>Putaway</th>
          <th>Backlog</th>
          <th>Status</th>
          <th>Arrived</th>
        `;
      }}
    }}

    function getDateMonthFilteredItems() {{
      const items = getRawBucketItems();
      if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
        return items;
      }}
      applyLocalArrivedOverrides(items);
      return items.filter(r => {{
        const rawDate = r['Receive Date'] || r['Create Time'] || '';
        const p = parseReceiveDate(rawDate);
        if (selectedMonth !== 'ALL' && p.mmyyyy !== selectedMonth) return false;
        if (selectedDates.length > 0) {{
          if (!selectedDates.includes(p.ddmmyyyy)) return false;
        }} else if (selectedDate !== 'ALL') {{
          if (p.ddmmyyyy !== selectedDate) return false;
        }} else if (currentBucket === 'Inbound') {{
          // Nếu không chọn gì thì auto là ngày D (tính từ đầu tháng đến ngày D hôm nay)
          const today = getTodayInfo();
          if (p.sortKey && p.sortKey > today.sortKey) return false;
        }}
        return true;
      }});
    }}

    function showToast(msg) {{
      const t = document.getElementById('toast');
      t.textContent = msg;
      t.classList.add('show');
      setTimeout(() => t.classList.remove('show'), 2500);
    }}

    function fmtNum(n) {{
      return Number(n || 0).toLocaleString('vi-VN');
    }}

    function updateUI() {{
      updateFilterLabels();
      const activeItems = getDateMonthFilteredItems();

      if (currentBucket === 'RTS') {{
        // Cập nhật giao diện RTS theo từng Order (mỗi order 1 dòng)
        const uniqueOrders = activeItems.length;
        const totalQty = activeItems.reduce((s, r) => s + Number(r['Total QTY'] || r.QTY || 0), 0);
        const distinctSkus = activeItems.reduce((s, r) => s + Number(r['Total SKU'] || 1), 0);
        const outboundOrders = activeItems.filter(r => r.Status === 'Outbound');
        const outboundQty = outboundOrders.reduce((s, r) => s + Number(r['Total QTY'] || r.QTY || 0), 0);
        const rate = totalQty > 0 ? Number((outboundQty / totalQty * 100).toFixed(1)) : 0;

        const cntAlloc = activeItems.filter(r => r.Status === 'Allocated').length;
        const cntOut = activeItems.filter(r => r.Status === 'Outbound').length;
        const cntPick = activeItems.filter(r => r.Status === 'Pending Pick').length;

        document.getElementById('kpi-orders').textContent = fmtNum(uniqueOrders);
        document.getElementById('kpi-qty').textContent = fmtNum(distinctSkus);
        document.getElementById('kpi-putaway').textContent = fmtNum(totalQty);
        document.getElementById('kpi-backlog').textContent = fmtNum(outboundQty);

        const bdDiv = document.getElementById('kpi-backlog-breakdown');
        if (bdDiv) bdDiv.style.display = 'none';
        const cardActualPut = document.getElementById('kpi-card-actual-put');
        if (cardActualPut) cardActualPut.style.display = 'none';
        const cardTruck = document.getElementById('kpi-card-truck-tracker');
        if (cardTruck) cardTruck.style.display = 'none';

        document.getElementById('pipeline-rate').textContent = rate + '% Outbound';
        document.getElementById('pipe-lbl-1').textContent = '1. TỔNG ORDER';
        document.getElementById('stat-total-asn').textContent = fmtNum(uniqueOrders);
        document.getElementById('stat-total-asn').style.color = '#38bdf8';

        document.getElementById('pipe-lbl-2').textContent = '2. ALLOCATED';
        document.getElementById('stat-pending').textContent = fmtNum(cntAlloc);
        document.getElementById('stat-pending').style.color = '#fde047';

        document.getElementById('pipe-lbl-3').textContent = '3. PENDING PICK';
        document.getElementById('stat-arriving').textContent = fmtNum(cntPick);
        document.getElementById('stat-arriving').style.color = '#f87171';

        document.getElementById('pipe-lbl-4').textContent = '4. OUTBOUND';
        document.getElementById('stat-in-prog').textContent = fmtNum(cntOut);
        document.getElementById('stat-in-prog').style.color = '#4ade80';

        const stat5 = document.querySelector('.pipeline-stats .stat-item:nth-child(5)');
        if (stat5) stat5.style.display = 'none';

        const denom = uniqueOrders || 1;
        document.getElementById('bar-done').style.width = (cntOut / denom * 100).toFixed(1) + '%';
        document.getElementById('bar-done').style.background = '#22c55e';
        document.getElementById('bar-prog').style.width = '0%';
        document.getElementById('bar-arriving').style.width = (cntPick / denom * 100).toFixed(1) + '%';
        document.getElementById('bar-arriving').style.background = '#ef4444';
        document.getElementById('bar-pend').style.width = (cntAlloc / denom * 100).toFixed(1) + '%';
        document.getElementById('bar-pend').style.background = '#eab308';

        const cAll = document.getElementById('cnt-rts-all');
        if (cAll) cAll.textContent = uniqueOrders;
        const cAl = document.getElementById('cnt-rts-alloc');
        if (cAl) cAl.textContent = cntAlloc;
        const cOut = document.getElementById('cnt-rts-out');
        if (cOut) cOut.textContent = cntOut;
        const cPick = document.getElementById('cnt-rts-pick');
        if (cPick) cPick.textContent = cntPick;

      }} else if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
        const ov = currentData.overview_vndl || {{}};
        const items = ov.items || [];
        const tot = ov.total || {{}};

        const totLoc = tot['Location Qty'] || items.reduce((s, r) => s + Number(r['Location Qty'] || 0), 0);
        const totUsed = tot['Used Qty'] || items.reduce((s, r) => s + Number(r['Used Qty'] || 0), 0);
        const totEmpty = tot['Empty Qty'] || items.reduce((s, r) => s + Number(r['Empty Qty'] || 0), 0);
        const totMaxCap = tot['Max_capacity(m³)'] || items.reduce((s, r) => s + Number(r['Max_capacity(m³)'] || 0), 0);
        const totVolUsed = tot['Volume Used(m³)'] || items.reduce((s, r) => s + Number(r['Volume Used(m³)'] || 0), 0);
        const totOnhand = tot['Qty Onhands'] || items.reduce((s, r) => s + Number(r['Qty Onhands'] || 0), 0);

        const occupancyRate = totLoc > 0 ? Number((totUsed / totLoc * 100).toFixed(1)) : 0;
        const cbmRate = totMaxCap > 0 ? Number((totVolUsed / totMaxCap * 100).toFixed(1)) : 0;

        const totEmptyCap = tot['Empty_Capacity(m³)'] || (totMaxCap - totVolUsed);

        document.getElementById('kpi-orders').textContent = fmtNum(totOnhand);
        document.getElementById('kpi-qty').textContent = fmtNum(totMaxCap) + ' m³';
        document.getElementById('kpi-putaway').textContent = fmtNum(totVolUsed) + ` m³ (${{cbmRate}}%)`;
        document.getElementById('kpi-backlog').textContent = fmtNum(totEmptyCap) + ' m³';

        const bdDiv = document.getElementById('kpi-backlog-breakdown');
        if (bdDiv) bdDiv.style.display = 'none';
        const cardActualPut = document.getElementById('kpi-card-actual-put');
        if (cardActualPut) cardActualPut.style.display = 'none';
        const cardTruck = document.getElementById('kpi-card-truck-tracker');
        if (cardTruck) cardTruck.style.display = 'none';

        document.getElementById('pipeline-rate').textContent = occupancyRate + '% Vị Trí Đã Dùng • ' + cbmRate + '% Dung Tích Sử Dụng';
        document.getElementById('pipe-lbl-1').textContent = '1. HÀNG TỒN ON-HANDS';
        document.getElementById('stat-total-asn').textContent = fmtNum(totOnhand);
        document.getElementById('stat-total-asn').style.color = '#4ade80';

        document.getElementById('pipe-lbl-2').textContent = '2. TỔNG SỨC CHỨA';
        document.getElementById('stat-pending').textContent = fmtNum(totMaxCap) + ' m³';
        document.getElementById('stat-pending').style.color = '#38bdf8';

        document.getElementById('pipe-lbl-3').textContent = '3. DUNG TÍCH ĐÃ DÙNG';
        document.getElementById('stat-arriving').textContent = fmtNum(totVolUsed) + ` m³ (${{cbmRate}}%)`;
        document.getElementById('stat-arriving').style.color = '#fde047';

        document.getElementById('pipe-lbl-4').textContent = '4. DUNG TÍCH TRỐNG';
        document.getElementById('stat-in-prog').textContent = fmtNum(totEmptyCap) + ' m³';
        document.getElementById('stat-in-prog').style.color = '#60a5fa';

        const stat5 = document.querySelector('.pipeline-stats .stat-item:nth-child(5)');
        if (stat5) {{
          stat5.style.display = 'block';
          document.getElementById('pipe-lbl-5').textContent = '5. VỊ TRÍ ĐÃ DÙNG (Y)';
          document.getElementById('stat-done').textContent = `${{fmtNum(totUsed)}} / ${{fmtNum(totLoc)}} (${{occupancyRate}}%)`;
          document.getElementById('stat-done').style.color = '#f87171';
        }}

        const denom = totLoc || 1;
        document.getElementById('bar-done').style.width = (totUsed / denom * 100).toFixed(1) + '%';
        document.getElementById('bar-done').style.background = '#ef4444';
        document.getElementById('bar-prog').style.width = '0%';
        document.getElementById('bar-arriving').style.width = (totEmpty / denom * 100).toFixed(1) + '%';
        document.getElementById('bar-arriving').style.background = '#3b82f6';
        document.getElementById('bar-pend').style.width = '0%';

        const cAll = document.getElementById('cnt-ov-all');
        if (cAll) cAll.textContent = items.length;
        const cRack = document.getElementById('cnt-ov-rack');
        if (cRack) cRack.textContent = items.filter(r => (r.Type || '').includes('RACK')).length;
        const cPallet = document.getElementById('cnt-ov-pallet');
        if (cPallet) cPallet.textContent = items.filter(r => (r.Type || '').includes('PALLET')).length;

      }} else {{
        const stat5 = document.querySelector('.pipeline-stats .stat-item:nth-child(5)');
        if (stat5) stat5.style.display = '';

        document.getElementById('bar-done').style.background = '';
        document.getElementById('bar-prog').style.background = '';
        document.getElementById('bar-arriving').style.background = '';
        document.getElementById('bar-pend').style.background = '';

        document.getElementById('stat-total-asn').style.color = '';
        document.getElementById('stat-pending').style.color = '';
        document.getElementById('stat-arriving').style.color = '';
        document.getElementById('stat-in-prog').style.color = '';
        // Cập nhật giao diện MTI / Inbound
        const totalOrders = activeItems.length;
        let totalQty = activeItems.reduce((s, r) => s + Number(r.Quantity || 0), 0);
        if (currentBucket === 'Inbound' && selectedDate === 'ALL') {{
          const today = getTodayInfo();
          const curYearMonth = today.sortKey.substring(0, 6);
          const startOfMonthKey = curYearMonth + '01';
          const endOfDKey = today.sortKey;
          const qtyItems = activeItems.filter(r => {{
            const rawDate = r['Receive Date'] || r['Create Time'] || '';
            const p = parseReceiveDate(rawDate);
            if (!p.sortKey) return false;
            if (selectedMonth === 'ALL' || p.sortKey.substring(0, 6) === curYearMonth) {{
              return p.sortKey >= startOfMonthKey && p.sortKey <= endOfDKey;
            }}
            return true;
          }});
          totalQty = qtyItems.reduce((s, r) => s + Number(r.Quantity || 0), 0);
        }}
        let totalPut = activeItems.reduce((s, r) => {{
          if (currentBucket === 'MTI' && (r.Status === 'Done' || r.Arrived === 'Done')) {{
            return s + Number(r.Quantity || 0);
          }}
          return s + Number(r.Putaway || 0);
        }}, 0);
        const totalBacklog = activeItems.reduce((s, r) => {{
          if (r.Status === 'Done' || r.Arrived === 'Done') return s;
          return s + Number(r.Backlog || 0);
        }}, 0);
        const cntPending = activeItems.filter(r => r.Status === 'Pending').length;
        const cntArriving = activeItems.filter(r => r.Status === 'Arriving').length;
        const cntInProg = activeItems.filter(r => r.Status === 'In Progress').length;
        const cntDone = activeItems.filter(r => r.Status === 'Done').length;
        const rate = totalQty > 0 ? Number((totalPut / totalQty * 100).toFixed(1)) : 0;

        document.getElementById('kpi-orders').textContent = fmtNum(totalOrders);
        document.getElementById('kpi-qty').textContent = fmtNum(totalQty);

        // CẬP NHẬT PLAN VOLUME MTI TỪ currentData.plan_mti (Sheet Data Plan MTI)
        const planMtiList = currentData.plan_mti || [];
        const today = getTodayInfo();
        const planFiltered = planMtiList.filter(r => {{
          if (!r.date) return false;
          const p = parseReceiveDate(r.date);
          if (selectedDates && selectedDates.length > 0) return selectedDates.includes(p.ddmmyyyy);
          if (selectedDate && selectedDate !== 'ALL') return p.ddmmyyyy === selectedDate;
          if (selectedMonth && selectedMonth !== 'ALL') {{
            return p.mmyyyy === selectedMonth && (p.sortKey <= today.sortKey || p.mmyyyy < today.mmyyyy);
          }}
          // Nếu không chọn gì: Từ ngày 1 tháng hiện tại → ngày D
          const curMonthKey = today.sortKey.substring(0, 6);
          const startOfCurMonth = curMonthKey + '01';
          return p.sortKey >= startOfCurMonth && p.sortKey <= today.sortKey;
        }});

        let totalPlanQty = 0;
        const planWhs = new Set();
        planFiltered.forEach(r => {{
          totalPlanQty += Number(r.qty || 0);
          if (r.warehouse) planWhs.add(r.warehouse);
        }});

        // Fallback sang data_mti nếu plan_mti không có dòng nào (ví dụ xem dữ liệu cũ chưa có plan)
        if (totalPlanQty === 0) {{
          const allDataMti = currentData.data_mti || [];
          const planFilteredOld = allDataMti.filter(r => {{
            const [d] = r;
            if (!d) return false;
            const p = parseReceiveDate(d);
            if (selectedDates && selectedDates.length > 0) return selectedDates.includes(p.ddmmyyyy);
            if (selectedDate && selectedDate !== 'ALL') return p.ddmmyyyy === selectedDate;
            if (selectedMonth && selectedMonth !== 'ALL') return p.mmyyyy === selectedMonth;
            return true;
          }});
          planFilteredOld.forEach(r => {{
            totalPlanQty += Number(r[4] || 0);
          }});
        }}

        const cardPlanVol = document.getElementById('kpi-card-plan-volume');
        const planVolEl = document.getElementById('kpi-plan-volume');
        const planVolSub = document.getElementById('kpi-plan-sub');
        if (cardPlanVol) {{
          cardPlanVol.style.display = (currentBucket === 'MTI') ? '' : 'none';
        }}
        if (planVolEl) planVolEl.textContent = fmtNum(totalPlanQty);
        if (planVolSub) {{
          planVolSub.textContent = (planWhs.size > 0) 
            ? `Kế hoạch (${{planWhs.size}} kho) • Click xem`
            : `Kế hoạch từ Data Plan MTI • Click xem`;
        }}

        // CẬP NHẬT GAP VOLUME CARD
        const cardGapVol = document.getElementById('kpi-card-gap-volume');
        const gapVolEl = document.getElementById('kpi-gap-volume');
        const gapSubEl = document.getElementById('kpi-gap-sub');
        const gapPctEl = document.getElementById('kpi-gap-pct');
        if (cardGapVol) {{
          cardGapVol.style.display = (currentBucket === 'MTI') ? '' : 'none';
        }}
        if (currentBucket === 'MTI') {{
          // Actual = totalQty (from items)
          const gapVal = totalQty - totalPlanQty;
          const gapPct = totalPlanQty > 0 ? ((gapVal / totalPlanQty) * 100).toFixed(1) : (totalQty > 0 ? '+∞' : '0.0');
          const gapPrefix = gapVal >= 0 ? '+' : '';
          if (gapVolEl) gapVolEl.textContent = gapPrefix + fmtNum(gapVal);
          if (gapSubEl) gapSubEl.textContent = `Plan: ${{fmtNum(totalPlanQty)}} | Actual: ${{fmtNum(totalQty)}}`;
          if (gapPctEl) gapPctEl.textContent = gapPrefix + gapPct + '% Gap';
          if (gapVolEl) gapVolEl.style.color = gapVal >= 0 ? '#4ade80' : '#fb7185';
          if (gapPctEl) gapPctEl.style.color = gapVal >= 0 ? '#4ade80' : '#fb7185';
        }}

        const cardQtyTitle = document.getElementById('kpi-qty-title');
        const cardQtySub = document.getElementById('kpi-qty-sub');
        if (cardQtyTitle) {{
          cardQtyTitle.textContent = (currentBucket === 'MTI') ? 'Actual Volume MTI' : 'Tổng Số Lượng (Quantity)';
        }}
        if (cardQtySub) {{
          if (currentBucket === 'Inbound') {{
            cardQtySub.textContent = (selectedDate === 'ALL') ? 'Kế hoạch nhập từ đầu tháng đến ngày D' : 'Kế hoạch nhập từ data IB';
          }} else if (currentBucket === 'MTI') {{
            cardQtySub.textContent = 'Thực tế nhập theo SO';
          }} else {{
            cardQtySub.textContent = 'Kế hoạch nhập';
          }}
        }}
        document.getElementById('kpi-putaway').textContent = fmtNum(totalPut);
        document.getElementById('kpi-backlog').textContent = fmtNum(totalBacklog);


        const bdDiv = document.getElementById('kpi-backlog-breakdown');
        if (bdDiv) bdDiv.style.display = 'flex';
        const pendingBacklogQty = activeItems.filter(r => (r.Status === 'Pending' || r.Arrived === 'No') && r.Status !== 'Done' && r.Arrived !== 'Done').reduce((s, r) => s + Number(r.Backlog || 0), 0);
        const arrivedBacklogQty = Math.max(0, totalBacklog - pendingBacklogQty);
        const elPend = document.getElementById('kpi-backlog-pending');
        if (elPend) elPend.textContent = fmtNum(pendingBacklogQty) + ' QTY';
        const elArr = document.getElementById('kpi-backlog-arriving');
        if (elArr) elArr.textContent = fmtNum(arrivedBacklogQty) + ' QTY';

        const cardActualPut = document.getElementById('kpi-card-actual-put');
        if (cardActualPut) cardActualPut.style.display = '';
        const cardTruck = document.getElementById('kpi-card-truck-tracker');
        if (cardTruck) {{
          cardTruck.style.display = (currentBucket === 'MTI') ? '' : 'none';
          const truckData = currentData.truck_tracker || {{}};
          const count = truckData.total || (truckData.items ? truckData.items.length : 0);
          const countEl = document.getElementById('kpi-truck-count');
          if (countEl) {{
            countEl.textContent = count > 0 ? (count + ' xe') : '🚚 Tra cứu xe';
          }}
        }}
        const actualPuts = getFilteredActualPutaway();
        const totalActualPutQty = actualPuts.reduce((s, r) => s + Number(r.qty || 0), 0);
        let displayActualPut = totalActualPutQty;
        if (currentBucket === 'MTI') {{
          if (cntDone === totalOrders && totalOrders > 0) {{
            displayActualPut = totalPut;
          }} else {{
            displayActualPut = Math.max(totalActualPutQty, totalPut);
          }}
        }}
        const elActualPut = document.getElementById('kpi-actual-putaway');
        if (elActualPut) elActualPut.textContent = fmtNum(displayActualPut);
        const cardActualSub = document.getElementById('kpi-actual-sub');
        if (cardActualSub) {{
          cardActualSub.textContent = (currentBucket === 'Inbound') ? 'Thực tế lên kệ (Put IB)' : 'Thực tế lên kệ (Put MTI)';
        }}

        document.getElementById('pipeline-rate').textContent = rate + '% Hoàn Thành';
        document.getElementById('pipe-lbl-1').textContent = '1. TỔNG ASN';
        document.getElementById('stat-total-asn').textContent = fmtNum(totalOrders);
        document.getElementById('pipe-lbl-2').textContent = '2. PENDING (XE CHƯA ĐẾN)';
        document.getElementById('stat-pending').textContent = fmtNum(cntPending);
        document.getElementById('pipe-lbl-3').textContent = '3. ARRIVING (ĐÃ ĐẾN)';
        document.getElementById('stat-arriving').textContent = fmtNum(cntArriving);
        document.getElementById('pipe-lbl-4').textContent = '4. ĐANG LÀM (IN PROGRESS)';
        document.getElementById('stat-in-prog').textContent = fmtNum(cntInProg);
        document.getElementById('pipe-lbl-5').textContent = '5. DONE';
        document.getElementById('stat-done').textContent = fmtNum(cntDone);

        const denom = totalOrders || 1;
        const pctDone = totalOrders > 0 ? (cntDone / denom * 100).toFixed(1) : 0;
        const pctProg = totalOrders > 0 ? (cntInProg / denom * 100).toFixed(1) : 0;
        const pctArr = totalOrders > 0 ? (cntArriving / denom * 100).toFixed(1) : 0;
        const pctPend = totalOrders > 0 ? Math.max(0, 100 - pctDone - pctProg - pctArr).toFixed(1) : 0;
        document.getElementById('bar-done').style.width = pctDone + '%';
        document.getElementById('bar-prog').style.width = pctProg + '%';
        document.getElementById('bar-arriving').style.width = pctArr + '%';
        document.getElementById('bar-pend').style.width = pctPend + '%';

        const cAll = document.getElementById('cnt-all');
        if (cAll) cAll.textContent = totalOrders;
        const cPend = document.getElementById('cnt-pending');
        if (cPend) cPend.textContent = cntPending;
        const cArr = document.getElementById('cnt-arriving');
        if (cArr) cArr.textContent = cntArriving;
        const cProg = document.getElementById('cnt-inprog');
        if (cProg) cProg.textContent = cntInProg;
        const cDone = document.getElementById('cnt-done');
        if (cDone) cDone.textContent = cntDone;
      }}

      // Cập nhật số xe / chuyến trên ô KPI Daily truck tracker
      const truckCard = document.getElementById('kpi-card-truck-tracker');
      if (truckCard) {{
        if (currentBucket === 'MTI') {{
          truckCard.style.display = '';
          const allTrucks = (currentData.truck_tracker && currentData.truck_tracker.items) ? currentData.truck_tracker.items : [];
          const today = getTodayInfo();
          const targetTruckDate = (selectedDate && selectedDate !== 'ALL') ? selectedDate : today.ddmmyyyy;
          // Ở tab MTI: số chuyến ngày D lấy data từ ngày D - 1 ở sheet Daily truck tracker
          const planTruckDate = (currentBucket === 'MTI') ? getTruckPlanDate(targetTruckDate) : targetTruckDate;

          const dayTrucks = allTrucks.filter(r => {{
            const pTh = getTruckVal(r, ['Pick thêm', 'Pick them']).toLowerCase();
            const kDen = getTruckVal(r, ['Kho đến', 'Kho den']).toLowerCase();
            const dTh = getTruckVal(r, ['Drop thêm', 'Drop them']).toLowerCase();
            const hasDaNang = (pTh.includes('đà nẵng') || pTh.includes('da nang') || pTh.includes('vnd') ||
                               kDen.includes('đà nẵng') || kDen.includes('da nang') || kDen.includes('vnd') ||
                               dTh.includes('đà nẵng') || dTh.includes('da nang') || dTh.includes('vnd'));
            if (!hasDaNang) return false;

            const d1 = getTruckVal(r, ['Ngày MTO', 'Ngay MTO', 'Ngày', 'Ngay', 'Date']);
            // Kế hoạch xe đến ngày D được lấy từ Ngày MTO = D - 1 (leadtime 24h)
            if (d1 && parseReceiveDate(d1).ddmmyyyy === planTruckDate) return true;  // MTO = D-1
            return false;
          }});
          const truckCountEl = document.getElementById('kpi-truck-count');
          const truckSubEl = document.getElementById('kpi-truck-sub');
          if (truckCountEl) {{
            truckCountEl.textContent = `${{dayTrucks.length}} Chuyến`;
          }}
          if (truckSubEl) {{
            const isToday = (targetTruckDate === today.ddmmyyyy);
            const dateDesc = (currentBucket === 'MTI' && planTruckDate !== targetTruckDate)
              ? (isToday ? `Kế hoạch ${{planTruckDate}} (D-1)` : `Kế hoạch ${{planTruckDate}}`)
              : (isToday ? `Ngày ${{targetTruckDate}} (D)` : `Ngày ${{targetTruckDate}}`);
            truckSubEl.textContent = `${{dateDesc}} • Click xem`;
          }}

          // Cross-check thông tin xe đến VNDL
          const xeDenList = (currentData.truck_tracker && currentData.truck_tracker.xe_den_vndl) ? currentData.truck_tracker.xe_den_vndl : [];
          const dayXeDen = xeDenList.filter(x => {{
            const xDate = getTruckVal(x, ['Dấu thời gian', 'Thời gian', 'Ngày', 'Ngay', 'date', 'ngay', 'Time']);
            if (!xDate) return false;
            return parseReceiveDate(xDate).ddmmyyyy === targetTruckDate;
          }});

          let arrivedCount = 0;
          const matchedXdSet = new Set();
          dayTrucks.forEach(trk => {{
            const p1 = cleanTruckPlate(getTruckVal(trk, ['Số xe', 'So xe', 'Biển số', 'so_xe', 'Biển số xe']));
            const ph1 = cleanTruckPhone(getTruckVal(trk, ['Thông tin tài xế', 'SĐT', 'Số điện thoại', 'sdt', 'Phone', 'Số ĐT']));
            const matchedXd = dayXeDen.find(xd => {{
              const p2 = cleanTruckPlate(getTruckVal(xd, ['Biển số xe', 'Biển số', 'Số xe', 'So xe', 'so_xe']));
              const ph2 = cleanTruckPhone(getTruckVal(xd, ['Số ĐT', 'Số điện thoại', 'SĐT', 'sdt', 'Phone', 'Thông tin tài xế']));
              if (p1 && p2 && (p1.includes(p2) || p2.includes(p1))) return true;
              if (ph1 && ph2 && ph1.length >= 8 && ph2.length >= 8 && (ph1.includes(ph2) || ph2.includes(ph1))) return true;
              return false;
            }});
            if (matchedXd) {{
              arrivedCount++;
              matchedXdSet.add(matchedXd);
            }}
          }});

          const extraArrived = dayXeDen.filter(xd => !matchedXdSet.has(xd)).length;
          const totalArrived = arrivedCount + extraArrived;
          const notArrivedCount = Math.max(0, dayTrucks.length - arrivedCount);

          if (dayTrucks.length === 0 && totalArrived > 0) {{
            truckCountEl.textContent = `${{totalArrived}} Chuyến`;
          }} else {{
            truckCountEl.textContent = `${{dayTrucks.length}} Chuyến`;
          }}

          const arrBadge = document.getElementById('kpi-truck-arrived-badge');
          const notArrBadge = document.getElementById('kpi-truck-not-arrived-badge');
          if (arrBadge) arrBadge.textContent = `${{totalArrived}} xe đã đến`;
          if (notArrBadge) notArrBadge.textContent = `${{notArrivedCount}} xe chưa đến`;
        }} else {{
          truckCard.style.display = 'none';
        }}
      }}

      if (currentData.updated_at) {{
        document.getElementById('last-update').textContent = 'Cập nhật: ' + currentData.updated_at;
      }}

      renderTable(activeItems);
    }}

    function renderTable(baseItems) {{
      if (isDataMtiActive) return;
      const tbody = document.getElementById('tableBody');
      const q = (document.getElementById('searchInput').value || '').trim().toLowerCase();
      const items = baseItems || getDateMonthFilteredItems();

      const filtered = items.filter(r => {{
        if (currentFilter !== 'ALL') {{
          if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
            if (currentFilter === 'RACK' && !(r.Type || '').includes('RACK')) return false;
            if (currentFilter === 'PALLET' && !(r.Type || '').includes('PALLET')) return false;
          }} else if (currentBucket === 'RTS') {{
            if (r.Status !== currentFilter) return false;
          }} else {{
            if (r.Status !== currentFilter) return false;
          }}
        }}
        if (q) {{
          const txt = JSON.stringify(r).toLowerCase();
          if (!txt.includes(q)) return false;
        }}
        return true;
      }});

      if (filtered.length === 0) {{
        tbody.innerHTML = '<tr><td colspan="9" style="text-align:center; padding: 40px; color: var(--text-dim);">Không có dòng dữ liệu nào khớp bộ lọc.</td></tr>';
        return;
      }}

      let html = '';
      if (currentBucket === 'RTS') {{
        const widgetContainer = document.getElementById('overviewWidgetsContainer');
        if (widgetContainer) widgetContainer.style.display = 'none';

        try {{
          const st = JSON.parse(localStorage.getItem('rts_transfer_address') || '{{}}');
          let changed = false;
          for (const k in st) {{
            const v = String(st[k] || '').trim().toUpperCase();
            if (!['VNNL', 'VNA', 'VNVL', 'VNG', 'VNSL'].includes(v)) {{
              delete st[k];
              changed = true;
            }}
          }}
          if (changed) localStorage.setItem('rts_transfer_address', JSON.stringify(st));
          localStorage.removeItem('rts_return_address');
        }} catch (e) {{}}

        const savedTruck = JSON.parse(localStorage.getItem('rts_truck_info') || '{{}}');
        const savedMail = JSON.parse(localStorage.getItem('rts_mail_status') || '{{}}');
        const savedTransfer = JSON.parse(localStorage.getItem('rts_transfer_address') || '{{}}');

        for (const r of filtered) {{
          const on = r['Order Number'] || '';
          const sku = r['SKU'] || 'ALL';
          const status = r.Status || 'Allocated';
          const truckInfo = savedTruck[on] !== undefined ? savedTruck[on] : (r['Thông tin xe'] || r.truck_info || '');
          const mailStatus = savedMail[on] !== undefined ? savedMail[on] : (r['Mail'] || r.mail_status || 'Chưa mail');
          const retAddr = r['Return Address'] || r['Địa chỉ trả hàng'] || r.return_address || '';
          let transferAddr = '';
          const cand = (savedTransfer[on] !== undefined ? savedTransfer[on] : (r['Địa chỉ chuyển'] || r.transfer_address || '')).toString().trim();
          if (['VNNL', 'VNA', 'VNVL', 'VNG', 'VNSL'].includes(cand.toUpperCase())) {{
            transferAddr = cand.toUpperCase();
          }} else if (cand && cand.length <= 25 && !cand.toLowerCase().includes('đường') && !cand.toLowerCase().includes('tầng') && !cand.toLowerCase().includes('số ') && !cand.toLowerCase().includes('phố') && !cand.toLowerCase().includes('quận') && !cand.toLowerCase().includes('huyện')) {{
            transferAddr = cand;
          }}
          const supp = r['Supplier'] || '';

          if (rtsViewMode === 'items') {{
            html += `
              <tr>
                <td class="order-code">${{on}}</td>
                <td style="color: #38bdf8; font-weight: 600;">${{r['Request ID'] || ''}}</td>
                <td title="Địa chỉ trả NCC: ${{retAddr}}">${{supp}}</td>
                <td style="color: #38bdf8; font-weight: 600;">${{sku}}</td>
                <td style="color: #f1f5f9; font-size: 12px; max-width: 220px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${{r['SKU Name'] || ''}}">${{r['SKU Name'] || ''}}</td>
                <td style="color: #4ade80; text-align: center;"><strong>${{fmtNum(r['QTY'] || 0)}}</strong></td>
                <td>${{r['Create Time'] || ''}}</td>
                <td>${{r['Ship Date'] || ''}}</td>
                <td>
                  <select class="arrived-select ${{status.toLowerCase().replace(/[^a-z]/g, '')}}" onchange="toggleRTSStatus('${{on}}', '${{sku}}', this.value, this)">
                    <option value="Allocated" ${{status === 'Allocated' ? 'selected' : ''}}>Allocated</option>
                    <option value="Outbound" ${{status === 'Outbound' ? 'selected' : ''}}>Outbound</option>
                    <option value="Pending Pick" ${{status === 'Pending Pick' ? 'selected' : ''}}>Pending Pick</option>
                  </select>
                </td>
                <td>
                  <div class="whs-combobox">
                    <input type="text" 
                      id="whs_inp_${{on}}"
                      class="rts-whs-input ${{getWhsClass(transferAddr)}}" 
                      value="${{transferAddr}}" 
                      placeholder="Chọn kho... ▼" 
                      autocomplete="off"
                      onclick="toggleWhsDropdown('${{on}}', event)"
                      oninput="handleWhsInput(this, '${{on}}')"
                      onkeydown="handleWhsKeydown(event, '${{on}}', this)" 
                      onchange="updateRTSAddress('${{on}}', this.value)" 
                      title="Click để chọn kho hoặc gõ tay rồi Enter">
                    <div id="whs_dd_${{on}}" class="whs-dropdown-menu">
                      <div class="whs-opt-item whs-opt-vnnl" onclick="selectWhs('${{on}}', 'VNNL')"><span>VNNL</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-vna" onclick="selectWhs('${{on}}', 'VNA')"><span>VNA</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-vnvl" onclick="selectWhs('${{on}}', 'VNVL')"><span>VNVL</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-vng" onclick="selectWhs('${{on}}', 'VNG')"><span>VNG</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-vnsl" onclick="selectWhs('${{on}}', 'VNSL')"><span>VNSL</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-clear" onclick="selectWhs('${{on}}', '')">✕ Để trống</div>
                    </div>
                  </div>
                </td>
                <td>
                  <input type="text" class="rts-truck-input" value="${{truckInfo}}" placeholder="Nhập xe..." 
                    onkeydown="if(event.key==='Enter'){{ updateRTSTruckInfo('${{on}}', this.value); this.blur(); }}" 
                    onchange="updateRTSTruckInfo('${{on}}', this.value)" 
                    style="background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.15); border-radius: 4px; padding: 4px 8px; color: #fff; font-size: 12px; width: 110px; outline: none;">
                </td>
                <td style="text-align: center;">
                  <select class="arrived-select ${{mailStatus === 'Mail' ? 'yes' : 'no'}}" onchange="updateRTSMailStatus('${{on}}', this.value)" style="min-width: 90px; padding: 4px 6px; font-size: 11px;">
                    <option value="Chưa mail" ${{mailStatus === 'Chưa mail' ? 'selected' : ''}}>⏳ Chưa mail</option>
                    <option value="Mail" ${{mailStatus === 'Mail' ? 'selected' : ''}}>✉️ Mail</option>
                  </select>
                </td>
              </tr>
            `;
          }} else {{
            const skuCount = Number(r['Total SKU'] || 0);
            const rowKey = on.replace(/[^a-zA-Z0-9]/g,'_');
            html += `
              <tr class="rts-order-row" onclick="toggleRTSSkuDetail('${{rowKey}}', '${{on}}', event)" style="cursor:pointer;" title="Click để xem chi tiết SKU">
                <td class="order-code">
                  <span id="rts_arrow_${{rowKey}}" style="display:inline-block;margin-right:5px;font-size:11px;transition:transform 0.2s;">▶</span>
                  ${{on}}
                </td>
                <td style="color: #38bdf8; font-weight: 600;">${{r['Request ID'] || ''}}</td>
                <td title="Địa chỉ trả NCC: ${{retAddr}}">${{supp}}</td>
                <td style="color: var(--accent-blue); text-align: center;"><strong>${{fmtNum(skuCount)}}</strong></td>
                <td style="color: #4ade80; text-align: center;"><strong>${{fmtNum(r['Total QTY'])}}</strong></td>
                <td>${{r['Create Time'] || ''}}</td>
                <td>${{r['Ship Date'] || ''}}</td>
                <td onclick="event.stopPropagation()">
                  <select class="arrived-select ${{status.toLowerCase().replace(/[^a-z]/g, '')}}" onchange="toggleRTSStatus('${{on}}', 'ALL', this.value, this)">
                    <option value="Allocated" ${{status === 'Allocated' ? 'selected' : ''}}>Allocated</option>
                    <option value="Outbound" ${{status === 'Outbound' ? 'selected' : ''}}>Outbound</option>
                    <option value="Pending Pick" ${{status === 'Pending Pick' ? 'selected' : ''}}>Pending Pick</option>
                  </select>
                </td>
                <td onclick="event.stopPropagation()">
                  <div class="whs-combobox">
                    <input type="text" 
                      id="whs_inp_${{on}}"
                      class="rts-whs-input ${{getWhsClass(transferAddr)}}" 
                      value="${{transferAddr}}" 
                      placeholder="Chọn kho... ▼" 
                      autocomplete="off"
                      onclick="toggleWhsDropdown('${{on}}', event)"
                      oninput="handleWhsInput(this, '${{on}}')"
                      onkeydown="handleWhsKeydown(event, '${{on}}', this)" 
                      onchange="updateRTSAddress('${{on}}', this.value)" 
                      title="Click để chọn kho hoặc gõ tay rồi Enter">
                    <div id="whs_dd_${{on}}" class="whs-dropdown-menu">
                      <div class="whs-opt-item whs-opt-vnnl" onclick="selectWhs('${{on}}', 'VNNL')"><span>VNNL</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-vna" onclick="selectWhs('${{on}}', 'VNA')"><span>VNA</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-vnvl" onclick="selectWhs('${{on}}', 'VNVL')"><span>VNVL</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-vng" onclick="selectWhs('${{on}}', 'VNG')"><span>VNG</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-vnsl" onclick="selectWhs('${{on}}', 'VNSL')"><span>VNSL</span><span style="font-size: 9px; opacity: 0.8;">●</span></div>
                      <div class="whs-opt-item whs-opt-clear" onclick="selectWhs('${{on}}', '')">✕ Để trống</div>
                    </div>
                  </div>
                </td>
                <td onclick="event.stopPropagation()">
                  <input type="text" class="rts-truck-input" value="${{truckInfo}}" placeholder="Nhập xe..." 
                    onkeydown="if(event.key==='Enter'){{ updateRTSTruckInfo('${{on}}', this.value); this.blur(); }}" 
                    onchange="updateRTSTruckInfo('${{on}}', this.value)" 
                    style="background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.15); border-radius: 4px; padding: 4px 8px; color: #fff; font-size: 12px; width: 120px; outline: none;">
                </td>
                <td style="text-align: center;" onclick="event.stopPropagation()">
                  <select class="arrived-select ${{mailStatus === 'Mail' ? 'yes' : 'no'}}" onchange="updateRTSMailStatus('${{on}}', this.value)" style="min-width: 95px; padding: 4px 8px; font-size: 11px;">
                    <option value="Chưa mail" ${{mailStatus === 'Chưa mail' ? 'selected' : ''}}>⏳ Chưa mail</option>
                    <option value="Mail" ${{mailStatus === 'Mail' ? 'selected' : ''}}>✉️ Mail</option>
                  </select>
                </td>
              </tr>
              <tr id="rts_sku_row_${{rowKey}}" style="display:none;">
                <td colspan="11" style="padding:0; background: rgba(14,30,60,0.95);">
                  <div id="rts_sku_detail_${{rowKey}}" style="padding: 10px 24px 14px 36px;"></div>
                </td>
              </tr>
            `;
          }}
        }}
      }} else if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
        const widgetContainer = document.getElementById('overviewWidgetsContainer');
        const ov = currentData.overview_vndl || {{}};
        const chartData = ov.chart_data || [];
        const tot = ov.total || {{}};

        if (widgetContainer) {{
          widgetContainer.style.display = 'block';

          let barsHtml = '';
          chartData.forEach(cd => {{
            barsHtml += `
              <div style="display: flex; align-items: center; margin-bottom: 12px; font-size: 13px;">
                <div style="width: 45px; font-weight: 800; color: #38bdf8; text-align: right; padding-right: 12px;">${{cd.zone}}</div>
                <div style="flex: 1; height: 26px; display: flex; border-radius: 6px; overflow: hidden; background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.1);">
                  <div style="width: ${{cd.empty_pct}}%; background: #3b82f6; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 11px; font-weight: 700; min-width: ${{cd.empty > 0 ? '30px' : '0'}};" title="Empty: ${{fmtNum(cd.empty)}} (${{cd.empty_pct}}%)">
                    ${{cd.empty > 0 ? fmtNum(cd.empty) : ''}}
                  </div>
                  <div style="width: ${{cd.used_pct}}%; background: #ef4444; display: flex; align-items: center; justify-content: center; color: #fff; font-size: 11px; font-weight: 700; min-width: ${{cd.used > 0 ? '40px' : '0'}};" title="Used: ${{fmtNum(cd.used)}} (${{cd.used_pct}}%)">
                    ${{cd.used > 0 ? fmtNum(cd.used) : ''}}
                  </div>
                </div>
              </div>
            `;
          }});

          widgetContainer.innerHTML = `
            <div style="display: grid; grid-template-columns: 1fr 380px; gap: 24px; align-items: start;">
              <!-- CỘT TRÁI: BIỂU ĐỒ TỔNG LOCATION & LOCATION TRỐNG -->
              <div style="background: rgba(30, 41, 59, 0.7); padding: 18px 20px; border-radius: 10px; border: 1px solid var(--border-color);">
                <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 16px;">
                  <h3 style="margin: 0; font-size: 15px; font-weight: 700; color: #f1f5f9; display: flex; align-items: center; gap: 8px;">
                    <span>📊</span> Tổng Location và Location trống theo Zone
                  </h3>
                  <div style="display: flex; align-items: center; gap: 14px; font-size: 12px; font-weight: 600;">
                    <span style="display: flex; align-items: center; gap: 6px;"><span style="width: 12px; height: 12px; background: #3b82f6; border-radius: 3px; display: inline-block;"></span> Empty Qty</span>
                    <span style="display: flex; align-items: center; gap: 6px;"><span style="width: 12px; height: 12px; background: #ef4444; border-radius: 3px; display: inline-block;"></span> Used Qty</span>
                  </div>
                </div>
                
                <div>
                  ${{barsHtml}}
                </div>

                <!-- Trục tỷ lệ 0% -> 100% -->
                <div style="display: flex; margin-left: 45px; border-top: 1px solid rgba(255,255,255,0.15); padding-top: 6px; font-size: 10px; color: var(--text-dim); justify-content: space-between;">
                  <span>0%</span>
                  <span>25%</span>
                  <span>50%</span>
                  <span>75%</span>
                  <span>100%</span>
                </div>
              </div>

              <!-- CỘT PHẢI: KẾ HOẠCH & SIZE WIDGETS -->
              <div style="display: flex; flex-direction: column; gap: 16px;">
                <!-- NHẬP THEO PLAN -->
                <div style="background: rgba(30, 41, 59, 0.7); padding: 16px 18px; border-radius: 10px; border: 1px solid var(--border-color);">
                  <div style="font-size: 13px; font-weight: 700; color: #fb7185; margin-bottom: 10px; text-transform: uppercase;">
                    📋 Nhập theo Plan
                  </div>
                  <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                    <tr style="border-bottom: 1px solid rgba(255,255,255,0.08);">
                      <td style="padding: 6px 0; color: var(--text-dim);">Total OB:</td>
                      <td style="padding: 6px 0; text-align: right; font-weight: 700; color: #38bdf8;">10.000</td>
                    </tr>
                    <tr>
                      <td style="padding: 6px 0; color: var(--text-dim);">FC INB:</td>
                      <td style="padding: 6px 0; text-align: right; font-weight: 700; color: #4ade80;">50.000</td>
                    </tr>
                  </table>
                </div>

                <!-- SỐ LƯỢNG NHẬN THEO SIZE -->
                <div style="background: rgba(30, 41, 59, 0.7); padding: 16px 18px; border-radius: 10px; border: 1px solid var(--border-color);">
                  <div style="font-size: 13px; font-weight: 700; color: #fb7185; margin-bottom: 10px; text-transform: uppercase;">
                    📦 Số Lượng Nhận Theo Size
                  </div>
                  <table style="width: 100%; border-collapse: collapse; font-size: 12px;">
                    <tr style="border-bottom: 1px solid rgba(255,255,255,0.08);">
                      <td style="padding: 5px 0; font-weight: 600;">Kệ (Rack):</td>
                      <td style="padding: 5px 0; text-align: right; color: #fde047; font-weight: 700;">197.514</td>
                    </tr>
                    <tr style="border-bottom: 1px solid rgba(255,255,255,0.08);">
                      <td style="padding: 5px 0; font-weight: 600;">Pallet:</td>
                      <td style="padding: 5px 0; text-align: right; color: #38bdf8; font-weight: 700;">87.407</td>
                    </tr>
                    <tr>
                      <td style="padding: 5px 0; font-weight: 800; color: #f43f5e;">SUM Total:</td>
                      <td style="padding: 5px 0; text-align: right; color: #4ade80; font-weight: 800;">284.921</td>
                    </tr>
                  </table>
                </div>
              </div>
            </div>
          `;
        }}

        let displayRows = filtered;
        if (currentFilter === 'RACK') {{
          displayRows = filtered.filter(r => (r.Type || '').includes('RACK'));
        }} else if (currentFilter === 'PALLET') {{
          displayRows = filtered.filter(r => (r.Type || '').includes('PALLET'));
        }}

        for (const r of displayRows) {{
          const occRate = Number(r.Occupancy_Rate || ((r['Used Qty'] / (r['Location Qty'] || 1)) * 100)).toFixed(1);
          html += `
            <tr>
              <td style="font-weight: 700; color: var(--text-main);">${{r.Type || ''}}</td>
              <td class="order-code" style="color: #38bdf8; font-size: 14px;"><strong>${{r.Zone || ''}}</strong></td>
              <td style="color: #38bdf8; font-weight: 700;">${{fmtNum(r['Empty Qty'])}}</td>
              <td style="color: #f87171; font-weight: 700;">${{fmtNum(r['Used Qty'])}}</td>
              <td><strong>${{fmtNum(r['Location Qty'])}}</strong></td>
              <td style="color: #fde047;"><strong>${{fmtNum(r['Max_capacity(m³)'] || r['Max_capacity'])}}</strong></td>
              <td style="color: #fb7185;"><strong>${{fmtNum(r['Volume Used(m³)'] || r['Volume Used'])}}</strong></td>
              <td style="color: #38bdf8;"><strong>${{fmtNum(r['Empty_Capacity(m³)'] || r['Empty_Capacity'])}}</strong></td>
              <td style="color: #4ade80;"><strong>${{fmtNum(r['Qty Onhands'])}}</strong></td>
              <td>${{fmtNum(r['AVG Items available'])}}</td>
              <td>
                <div style="display: flex; align-items: center; gap: 8px;">
                  <div style="flex: 1; height: 8px; background: rgba(255,255,255,0.1); border-radius: 4px; overflow: hidden;">
                    <div style="width: ${{occRate}}%; height: 100%; background: ${{occRate > 90 ? '#ef4444' : (occRate > 70 ? '#f59e0b' : '#10b981')}};"></div>
                  </div>
                  <span style="font-size: 11px; font-weight: 700;">${{occRate}}%</span>
                </div>
              </td>
            </tr>
          `;
        }}

        if (tot && Object.keys(tot).length > 0) {{
          html += `
            <tr style="background: rgba(244, 63, 94, 0.18); font-weight: 800; border-top: 2px solid #f43f5e;">
              <td style="color: #fb7185;">TOTAL</td>
              <td style="color: #fb7185;">ALL</td>
              <td style="color: #38bdf8;">${{fmtNum(tot['Empty Qty'])}}</td>
              <td style="color: #f87171;">${{fmtNum(tot['Used Qty'])}}</td>
              <td>${{fmtNum(tot['Location Qty'])}}</td>
              <td style="color: #fde047;">${{fmtNum(tot['Max_capacity(m³)'] || tot['Max_capacity'])}}</td>
              <td style="color: #fb7185;">${{fmtNum(tot['Volume Used(m³)'] || tot['Volume Used'])}}</td>
              <td style="color: #38bdf8;">${{fmtNum(tot['Empty_Capacity(m³)'] || tot['Empty_Capacity'])}}</td>
              <td style="color: #4ade80;">${{fmtNum(tot['Qty Onhands'])}}</td>
              <td>${{fmtNum(tot['AVG Items available'])}}</td>
              <td style="color: #f43f5e;">${{tot.Occupancy_Rate || 95.1}}%</td>
            </tr>
          `;
        }}

      }} else if (currentBucket === 'Inbound') {{
        const widgetContainer = document.getElementById('overviewWidgetsContainer');
        if (widgetContainer) widgetContainer.style.display = 'none';

        for (const r of filtered) {{
          const orderId = r['PO ID'] || '';
          const asnId = r['ASN ID'] || '';
          let status = r.Status || 'Pending';
          let arrived = r.Arrived || 'No';
          let isDone = (String(status).trim().toLowerCase() === 'done' || String(arrived).trim().toLowerCase() === 'done');
          if (isDone) {{
            status = 'Done';
            arrived = 'Done';
          }}
          let badgeClass = 'badge-pending';
          if (status === 'Done') badgeClass = 'badge-done';
          else if (status === 'In Progress') badgeClass = 'badge-inprog';
          else if (status === 'Arriving') badgeClass = 'badge-arriving';

          const backlog = isDone ? 0 : Number(r.Backlog || 0);
          const backlogClass = backlog > 0 ? 'has-backlog' : 'zero';
          let selectClass = 'no';
          if (arrived === 'Yes') selectClass = 'yes';
          else if (arrived === 'Done') selectClass = 'done';

          const toggleKey = r['ASN ID'] || orderId;

          html += `
            <tr>
              <td class="order-code">${{orderId}}</td>
              <td>${{asnId}}</td>
              <td>${{r['Receive Date'] || ''}}</td>
              <td style="color: #38bdf8; font-weight: 600;">${{r['Expected Delivery'] || ''}}</td>
              <td>${{fmtNum(r['Total SKU'])}}</td>
              <td><strong>${{fmtNum(r['Quantity'])}}</strong></td>
              <td style="color: #4ade80;"><strong>${{fmtNum(r['Putaway'])}}</strong></td>
              <td class="backlog-num ${{backlogClass}}">${{fmtNum(backlog)}}</td>
              <td><span class="status-badge ${{badgeClass}}">${{status}}</span></td>
              <td>
                <select class="arrived-select ${{selectClass}}" onchange="toggleArrived('${{toggleKey}}', this.value)">
                  <option value="No" ${{arrived === 'No' ? 'selected' : ''}}>⏳ No (Chưa đến)</option>
                  <option value="Yes" ${{arrived === 'Yes' ? 'selected' : ''}}>🚚 Yes (Đã đến)</option>
                  <option value="Done" ${{arrived === 'Done' ? 'selected' : ''}}>✅ Done</option>
                </select>
              </td>
            </tr>
          `;
        }}
      }} else {{
        const widgetContainer = document.getElementById('overviewWidgetsContainer');
        if (widgetContainer) widgetContainer.style.display = 'none';

        for (const r of filtered) {{
          const orderId = r['Ship Order'] || r['PO ID'] || '';
          const asnId = r['ASN'] || r['ASN ID'] || '';
          let status = r.Status || 'Pending';
          let arrived = r.Arrived || 'No';
          let isDone = (String(status).trim().toLowerCase() === 'done' || String(arrived).trim().toLowerCase() === 'done');
          if (isDone) {{
            status = 'Done';
            arrived = 'Done';
          }}
          let badgeClass = 'badge-pending';
          if (status === 'Done') badgeClass = 'badge-done';
          else if (status === 'In Progress') badgeClass = 'badge-inprog';
          else if (status === 'Arriving') badgeClass = 'badge-arriving';

          const backlog = isDone ? 0 : Number(r.Backlog || 0);
          const putaway = isDone ? Number(r.Quantity || 0) : Number(r.Putaway || 0);
          const backlogClass = backlog > 0 ? 'has-backlog' : 'zero';
          let selectClass = 'no';
          if (arrived === 'Yes') selectClass = 'yes';
          else if (arrived === 'Done') selectClass = 'done';

          const toggleKey = r['Ship Order'] || r['ASN ID'] || orderId;

          html += `
            <tr>
              <td class="order-code">${{orderId}}</td>
              <td>${{asnId}}</td>
              <td>${{r['Receive Date'] || ''}}</td>
              <td>${{fmtNum(r['Total SKU'])}}</td>
              <td><strong>${{fmtNum(r['Quantity'])}}</strong></td>
              <td style="color: #4ade80;"><strong>${{fmtNum(putaway)}}</strong></td>
              <td class="backlog-num ${{backlogClass}}">${{fmtNum(backlog)}}</td>
              <td><span class="status-badge ${{badgeClass}}">${{status}}</span></td>
              <td>
                <select class="arrived-select ${{selectClass}}" onchange="toggleArrived('${{toggleKey}}', this.value)">
                  <option value="No" ${{arrived === 'No' ? 'selected' : ''}}>⏳ No (Chưa đến)</option>
                  <option value="Yes" ${{arrived === 'Yes' ? 'selected' : ''}}>🚚 Yes (Đã đến)</option>
                  <option value="Done" ${{arrived === 'Done' ? 'selected' : ''}}>✅ Done</option>
                </select>
              </td>
            </tr>
          `;
        }}
      }}
      tbody.innerHTML = html;
    }}

    async function toggleArrived(targetKey, newVal) {{
      const today = getTodayInfo();
      const keyPrefix = (currentBucket === 'Inbound') ? 'ib_' : 'mti_';
      try {{
        const saved = JSON.parse(localStorage.getItem(keyPrefix + 'arrived_overrides') || '{{}}');
        const savedDates = JSON.parse(localStorage.getItem(keyPrefix + 'arrived_dates') || '{{}}');
        saved[targetKey] = newVal;
        if (newVal === 'Yes' || newVal === 'Done') {{
          savedDates[targetKey] = today.ddmmyyyy;
        }} else {{
          delete savedDates[targetKey];
        }}
        localStorage.setItem(keyPrefix + 'arrived_overrides', JSON.stringify(saved));
        localStorage.setItem(keyPrefix + 'arrived_dates', JSON.stringify(savedDates));
      }} catch (e) {{}}

      const items = getRawBucketItems();
      for (const r of items) {{
        const rKey = r['Ship Order'] || r['ASN ID'] || r['PO ID'];
        if (rKey === targetKey) {{
          r.Arrived = newVal;
          const put = Number(r.Putaway || 0);
          const qty = Number(r.Quantity || 0);
          const backlog = Number(r.Backlog || 0);
          if (newVal === 'Done') {{
            r.Status = 'Done';
            r['Receive Date'] = today.ddmmyyyy;
          }} else if (put === 0) {{
            r.Status = (newVal === 'Yes') ? 'Arriving' : 'Pending';
            if (newVal === 'Yes') r['Receive Date'] = today.ddmmyyyy;
          }} else if (backlog <= 700 || put >= qty) {{
            r.Status = 'Done';
          }} else {{
            r.Status = 'In Progress';
          }}
          break;
        }}
      }}

      updateUI();
      showToast(`Đã lưu Arrived = ${{newVal}} vào SQL!`);

      const bParam = (currentBucket === 'Inbound') ? 'Inbound' : 'MTI';
      for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
        try {{
          const resp = await fetch(`${{base}}/api/set_arrived?ship_order=${{encodeURIComponent(targetKey)}}&arrived=${{encodeURIComponent(newVal)}}&bucket=${{bParam}}`);
          if (resp.ok) break;
        }} catch (e) {{}}
      }}
    }}

    async function toggleRTSStatus(orderNumber, sku, newStatus, selectElem) {{
      if (selectElem) {{
        selectElem.className = 'arrived-select ' + newStatus.toLowerCase().replace(/[^a-z]/g, '');
      }}
      const items = (currentData.rts && currentData.rts.items) ? currentData.rts.items : [];
      for (const r of items) {{
        if (r['Order Number'] === orderNumber && (r['SKU'] === sku || !sku || sku === 'ALL')) {{
          r.Status = newStatus;
        }}
      }}
      const orders = (currentData.rts && currentData.rts.orders) ? currentData.rts.orders : [];
      for (const o of orders) {{
        if (o['Order Number'] === orderNumber) {{
          o.Status = newStatus;
        }}
      }}
      updateUI();
      showToast(`Đã cập nhật Status [${{newStatus}}] cho đơn ${{orderNumber}}!`);

      for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
        try {{
          const resp = await fetch(`${{base}}/api/set_rts_status?order_number=${{encodeURIComponent(orderNumber)}}&sku=${{encodeURIComponent(sku)}}&status=${{encodeURIComponent(newStatus)}}`);
          if (resp.ok) break;
        }} catch (e) {{}}
      }}
    }}

    async function updateRTSAddress(orderNumber, val) {{
      const cleanVal = (val || '').trim();
      const savedTransfer = JSON.parse(localStorage.getItem('rts_transfer_address') || localStorage.getItem('rts_return_address') || '{{}}');
      savedTransfer[orderNumber] = cleanVal;
      localStorage.setItem('rts_transfer_address', JSON.stringify(savedTransfer));
      localStorage.setItem('rts_return_address', JSON.stringify(savedTransfer));

      const orders = (currentData.rts && currentData.rts.orders) ? currentData.rts.orders : [];
      for (const o of orders) {{
        if (o['Order Number'] === orderNumber) {{
          o['Địa chỉ chuyển'] = cleanVal;
          o['transfer_address'] = cleanVal;
        }}
      }}
      const items = (currentData.rts && currentData.rts.items) ? currentData.rts.items : [];
      for (const it of items) {{
        if (it['Order Number'] === orderNumber) {{
          it['Địa chỉ chuyển'] = cleanVal;
          it['transfer_address'] = cleanVal;
        }}
      }}
      showToast(`Đã lưu Địa chỉ chuyển đơn ${{orderNumber}}: ${{cleanVal || '(trống)'}}!`);

      for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
        try {{
          const resp = await fetch(`${{base}}/api/set_rts_address?order_number=${{encodeURIComponent(orderNumber)}}&address=${{encodeURIComponent(cleanVal)}}`);
          if (resp.ok) break;
        }} catch (e) {{}}
      }}
    }}

    async function updateRTSTruckInfo(orderNumber, val) {{
      const cleanVal = (val || '').trim();
      const savedTruck = JSON.parse(localStorage.getItem('rts_truck_info') || '{{}}');
      savedTruck[orderNumber] = cleanVal;
      localStorage.setItem('rts_truck_info', JSON.stringify(savedTruck));

      const orders = (currentData.rts && currentData.rts.orders) ? currentData.rts.orders : [];
      for (const o of orders) {{
        if (o['Order Number'] === orderNumber) {{
          o['Thông tin xe'] = cleanVal;
        }}
      }}
      const items = (currentData.rts && currentData.rts.items) ? currentData.rts.items : [];
      for (const it of items) {{
        if (it['Order Number'] === orderNumber) {{
          it['Thông tin xe'] = cleanVal;
        }}
      }}
      showToast(`Đã lưu Thông tin xe đơn ${{orderNumber}}: ${{cleanVal || '(trống)'}}!`);

      for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
        try {{
          const resp = await fetch(`${{base}}/api/set_rts_truck_info?order_number=${{encodeURIComponent(orderNumber)}}&truck_info=${{encodeURIComponent(cleanVal)}}`);
          if (resp.ok) break;
        }} catch (e) {{}}
      }}
    }}

    async function updateRTSMailStatus(orderNumber, val) {{
      const cleanVal = (val || '').trim() || 'Chưa mail';
      const savedMail = JSON.parse(localStorage.getItem('rts_mail_status') || '{{}}');
      savedMail[orderNumber] = cleanVal;
      localStorage.setItem('rts_mail_status', JSON.stringify(savedMail));

      const orders = (currentData.rts && currentData.rts.orders) ? currentData.rts.orders : [];
      for (const o of orders) {{
        if (o['Order Number'] === orderNumber) {{
          o['Mail'] = cleanVal;
        }}
      }}
      const items = (currentData.rts && currentData.rts.items) ? currentData.rts.items : [];
      for (const it of items) {{
        if (it['Order Number'] === orderNumber) {{
          it['Mail'] = cleanVal;
        }}
      }}
      renderTable();
      showToast(`Đã lưu trạng thái Mail [${{cleanVal}}] đơn ${{orderNumber}}!`);

      for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
        try {{
          const resp = await fetch(`${{base}}/api/set_rts_mail_status?order_number=${{encodeURIComponent(orderNumber)}}&mail_status=${{encodeURIComponent(cleanVal)}}`);
          if (resp.ok) break;
        }} catch (e) {{}}
      }}
    }}

    function getWhsClass(val) {{
      if (!val) return '';
      const v = String(val).trim().toUpperCase();
      if (v === 'VNNL') return 'whs-vnnl';
      if (v === 'VNA') return 'whs-vna';
      if (v === 'VNVL') return 'whs-vnvl';
      if (v === 'VNG') return 'whs-vng';
      if (v === 'VNSL') return 'whs-vnsl';
      return '';
    }}

    function toggleWhsDropdown(orderNumber, ev) {{
      ev.stopPropagation();
      const dd = document.getElementById(`whs_dd_${{orderNumber}}`);
      if (!dd) return;
      const isOpen = dd.style.display === 'block';
      closeAllWhsDropdowns();
      if (!isOpen) {{
        const inp = ev.currentTarget;
        const rect = inp.getBoundingClientRect();
        dd.style.top = `${{rect.bottom + 2}}px`;
        dd.style.left = `${{rect.left}}px`;
        dd.style.width = `${{Math.max(95, rect.width)}}px`;
        dd.style.display = 'block';
      }}
    }}

    function selectWhs(orderNumber, code) {{
      closeAllWhsDropdowns();
      const inp = document.getElementById(`whs_inp_${{orderNumber}}`);
      if (inp) {{
        inp.value = code;
        inp.className = `rts-whs-input ${{getWhsClass(code)}}`;
      }}
      updateRTSAddress(orderNumber, code);
    }}

    function handleWhsKeydown(ev, orderNumber, inp) {{
      if (ev.key === 'Enter') {{
        closeAllWhsDropdowns();
        inp.className = `rts-whs-input ${{getWhsClass(inp.value)}}`;
        updateRTSAddress(orderNumber, inp.value);
        inp.blur();
      }} else if (ev.key === 'Escape') {{
        closeAllWhsDropdowns();
        inp.blur();
      }}
    }}

    function handleWhsInput(inp, orderNumber) {{
      inp.className = `rts-whs-input ${{getWhsClass(inp.value)}}`;
    }}

    function closeAllWhsDropdowns() {{
      document.querySelectorAll('.whs-dropdown-menu').forEach(m => m.style.display = 'none');
    }}

    document.addEventListener('click', (e) => {{
      if (!e.target.closest('.whs-combobox') && !e.target.closest('.whs-dropdown-menu')) {{
        closeAllWhsDropdowns();
      }}
    }});

    window.addEventListener('scroll', () => closeAllWhsDropdowns(), true);

    function cleanTruckPlate(p) {{
      if (!p) return '';
      return p.toString().toUpperCase().replace(/[^A-Z0-9]/g, '');
    }}

    function cleanTruckPhone(p) {{
      if (!p) return '';
      const str = p.toString();
      const m = str.match(/(?:0|\+84)[35789]\d{8}/);
      if (m) return m[0].replace('+84', '0');
      const digits = str.replace(/\D/g, '');
      return digits.length >= 10 ? digits.slice(-10) : digits;
    }}

    let isDataMtiActive = false;

    function handleDataMtiSearch(val) {{
      const q = (val || '').trim().toLowerCase();
      const clearBtn = document.getElementById('clearDataMtiBtn');
      if (clearBtn) clearBtn.style.display = q ? 'block' : 'none';

      if (!q) {{
        isDataMtiActive = false;
        renderTableHead();
        renderTable();
        return;
      }}

      isDataMtiActive = true;
      const allDataMti = currentData.data_mti || [];
      const matches = allDataMti.filter(r => {{
        const [d, pallet, mtb, sku, qty, kho, so, asn] = r;
        return (pallet && String(pallet).toLowerCase().includes(q)) ||
               (mtb && String(mtb).toLowerCase().includes(q)) ||
               (sku && String(sku).toLowerCase().includes(q)) ||
               (so && String(so).toLowerCase().includes(q)) ||
               (asn && String(asn).toLowerCase().includes(q));
      }});

      renderDataMtiTable(matches, q);
    }}

    function clearDataMtiSearch() {{
      const inp = document.getElementById('dataMtiSearchInput');
      if (inp) {{
        inp.value = '';
        inp.focus();
      }}
      handleDataMtiSearch('');
    }}

    function renderDataMtiTable(rows, query) {{
      const trHead = document.getElementById('tableHeadRow');
      trHead.innerHTML = `
        <th style="width: 50px; text-align: center;">STT</th>
        <th style="width: 100px;">NGÀY</th>
        <th style="color: #38bdf8;">MTP (PALLET)</th>
        <th style="color: #fbbf24;">MTB (THÙNG)</th>
        <th style="color: #4ade80;">SKU</th>
        <th style="text-align: center;">QTY</th>
        <th>KHO CHUYỂN</th>
        <th>SHIP ORDER</th>
        <th>ASN ID</th>
      `;

      const tbody = document.getElementById('tableBody');
      if (rows.length === 0) {{
        tbody.innerHTML = `
          <tr>
            <td colspan="9" style="text-align: center; padding: 40px; color: var(--text-dim);">
              <div style="font-size: 28px; margin-bottom: 8px;">📦</div>
              Không tìm thấy dòng nào khớp với "<strong>${{query}}</strong>" trong bảng data_mti.
            </td>
          </tr>
        `;
        return;
      }}

      const limit = 250;
      const displayRows = rows.slice(0, limit);
      let html = '';
      displayRows.forEach((r, idx) => {{
        const [d, pallet, mtb, sku, qty, kho, so, asn] = r;
        html += `
          <tr>
            <td style="text-align: center; color: #94a3b8; font-size: 11px;">${{idx + 1}}</td>
            <td style="color: #cbd5e1; font-size: 12px;">${{d || ''}}</td>
            <td style="color: #38bdf8; font-weight: 700; font-family: 'JetBrains Mono', monospace;">${{pallet || ''}}</td>
            <td style="color: #fbbf24; font-weight: 700; font-family: 'JetBrains Mono', monospace;">${{mtb || ''}}</td>
            <td style="color: #4ade80; font-family: 'JetBrains Mono', monospace; font-size: 12px; font-weight: 600;">${{sku || ''}}</td>
            <td style="text-align: center; font-weight: 800; color: #fff;">${{fmtNum(qty || 0)}}</td>
            <td><span style="background: rgba(192, 132, 252, 0.2); border: 1px solid #9333ea; color: #c084fc; padding: 2px 7px; border-radius: 4px; font-size: 11px; font-weight: 800;">${{kho || ''}}</span></td>
            <td style="color: #cbd5e1; font-size: 12px; font-weight: 600;">${{so || ''}}</td>
            <td style="color: #94a3b8; font-size: 12px;">${{asn || ''}}</td>
          </tr>
        `;
      }});

      if (rows.length > limit) {{
        html += `
          <tr>
            <td colspan="9" style="text-align: center; padding: 12px; color: #38bdf8; font-size: 12px; background: rgba(56, 189, 248, 0.08); font-weight: 600;">
              ✨ Đang hiển thị ${{limit}} / ${{fmtNum(rows.length)}} dòng tìm thấy trong data_mti.
            </td>
          </tr>
        `;
      }}

      tbody.innerHTML = html;
    }}

    function handleSearch() {{
      renderTable();
    }}

    function setFilter(filter) {{
      currentFilter = filter;
      document.querySelectorAll('.filter-tab').forEach(b => {{
        b.classList.toggle('active', b.dataset.filter === filter);
      }});
      renderTable();
    }}

    function getFilteredActualPutaway() {{
      const putsData = currentData.actual_putaway || {{}};
      let list = [];
      if (currentBucket === 'Inbound') {{
        list = putsData.ib || [];
      }} else {{
        list = putsData.mti || [];
      }}

      return list.filter(r => {{
        const rawDate = r.date || '';
        const p = parseReceiveDate(rawDate);
        if (selectedMonth !== 'ALL' && p.mmyyyy !== selectedMonth) return false;
        if (selectedDates.length > 0) {{
          if (!selectedDates.includes(p.ddmmyyyy)) return false;
        }} else if (selectedDate !== 'ALL') {{
          if (p.ddmmyyyy !== selectedDate) return false;
        }} else if (currentBucket === 'Inbound') {{
          // Nếu không chọn gì thì auto là ngày D (tính từ đầu tháng đến ngày D hôm nay)
          const today = getTodayInfo();
          if (p.sortKey && p.sortKey > today.sortKey) return false;
        }}
        return true;
      }});
    }}



    // ==========================================
    // ACTUAL VOLUME MTI MODAL (Thực tế theo Kho)
    // ==========================================
    function openActualVolumeModal() {{
      const modal = document.getElementById('actualVolumeModal');
      if (modal) {{
        modal.style.display = 'flex';
        document.body.style.overflow = 'hidden';
        renderActualVolumeModal();
      }}
    }}

    function closeActualVolumeModal(e) {{
      if (e && e.target && e.target.id !== 'actualVolumeModal' && !e.target.closest('button')) return;
      const modal = document.getElementById('actualVolumeModal');
      if (modal) {{
        modal.style.display = 'none';
        document.body.style.overflow = '';
      }}
    }}

    function renderActualVolumeModal() {{
      const listEl = document.getElementById('modalActualVolumeList');
      const subEl = document.getElementById('modal-actual-volume-subtitle');
      if (!listEl) return;

      const allDataMti = currentData.data_mti || [];
      let targetDate = (selectedDate && selectedDate !== 'ALL') ? selectedDate : '';
      const subText = targetDate ? `Ngày ${{targetDate}}` : (selectedDates && selectedDates.length > 1 ? `${{selectedDates.length}} ngày đã chọn` : 'Tất cả các ngày');
      if (subEl) subEl.textContent = `Thống kê hàng về thực tế theo Kho chuyển • ${{subText}}`;

      const filtered = allDataMti.filter(r => {{
        const [d] = r;
        if (!d) return false;
        const p = parseReceiveDate(d);
        if (selectedDates && selectedDates.length > 0) return selectedDates.includes(p.ddmmyyyy);
        if (targetDate) return p.ddmmyyyy === targetDate;
        if (selectedMonth && selectedMonth !== 'ALL') return p.mmyyyy === selectedMonth;
        return true;
      }});

      const whMap = {{}};
      let grandQty = 0;
      const allSkusSet = new Set();

      filtered.forEach(r => {{
        const [d, pallet, mtb, sku, qty, kho] = r;
        const whName = (kho || 'Chưa xác định').trim().toUpperCase();
        const qVal = parseInt(qty, 10) || 0;

        if (!whMap[whName]) {{
          whMap[whName] = {{ wh: whName, qty: 0, skus: new Set() }};
        }}
        whMap[whName].qty += qVal;
        if (sku) {{
          whMap[whName].skus.add(String(sku).trim());
          allSkusSet.add(String(sku).trim());
        }}
        grandQty += qVal;
      }});

      const whList = Object.values(whMap).sort((a, b) => b.qty - a.qty);

      if (whList.length === 0) {{
        listEl.innerHTML = `<div style="text-align: center; padding: 30px; color: var(--text-dim);">Không có dữ liệu Actual Volume MTI trong ngày này.</div>`;
        document.getElementById('modalActualTotalWh').textContent = '0 kho chuyển';
        document.getElementById('modalActualTotalSku').textContent = '0 SKU';
        document.getElementById('modalActualTotalQty').textContent = 'Tổng: 0 QTY';
        return;
      }}

      let html = '';
      whList.forEach((item, idx) => {{
        const pct = grandQty > 0 ? ((item.qty / grandQty) * 100).toFixed(1) : 0;
        html += `
          <div style="display: grid; grid-template-columns: 50px 1fr 110px 120px 80px; align-items: center; padding: 10px 12px; background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06); border-radius: 8px;">
            <span style="color: var(--text-dim);">${{idx + 1}}</span>
            <strong style="color: #f8fafc; display: flex; align-items: center; gap: 6px;">
              <span style="color: #c084fc;">🏢</span> ${{item.wh}}
            </strong>
            <span style="text-align: right; color: #38bdf8; font-weight: 700;">${{fmtNum(item.skus.size)}} SKU</span>
            <span style="text-align: right; color: #4ade80; font-weight: 800; font-family: 'JetBrains Mono', monospace;">${{fmtNum(item.qty)}}</span>
            <span style="text-align: right; color: #f59e0b; font-weight: 700;">${{pct}}%</span>
          </div>
        `;
      }});

      listEl.innerHTML = html;
      document.getElementById('modalActualTotalWh').textContent = `${{whList.length}} kho chuyển`;
      document.getElementById('modalActualTotalSku').textContent = `${{fmtNum(allSkusSet.size)}} SKU`;
      document.getElementById('modalActualTotalQty').textContent = `Tổng: ${{fmtNum(grandQty)}} QTY`;
    }}

    // ==========================================
    // ==========================================
    // PLAN VOLUME MTI MODAL (Kế hoạch theo Ngày, Tháng & Kho chuyển)
    // ==========================================
    let planModalViewMode = 'wh'; // 'wh', 'daily', 'monthly'
    let planModalDate = 'ALL';
    let planModalMonth = 'ALL';

    function setPlanModalViewMode(mode) {{
      planModalViewMode = mode;
      const btnWh = document.getElementById('planTabWh');
      const btnDaily = document.getElementById('planTabDaily');
      const btnMonthly = document.getElementById('planTabMonthly');
      if (btnWh) {{
        btnWh.style.background = (mode === 'wh') ? '#6366f1' : 'transparent';
        btnWh.style.color = (mode === 'wh') ? '#fff' : '#94a3b8';
      }}
      if (btnDaily) {{
        btnDaily.style.background = (mode === 'daily') ? '#6366f1' : 'transparent';
        btnDaily.style.color = (mode === 'daily') ? '#fff' : '#94a3b8';
      }}
      if (btnMonthly) {{
        btnMonthly.style.background = (mode === 'monthly') ? '#6366f1' : 'transparent';
        btnMonthly.style.color = (mode === 'monthly') ? '#fff' : '#94a3b8';
      }}
      renderPlanVolumeModal();
    }}

    function populatePlanModalDateSelect() {{
      const select = document.getElementById('planModalDateSelect');
      if (!select) return;
      const planMtiData = currentData.plan_mti || [];
      const dateSet = new Set();
      planMtiData.forEach(r => {{
        if (r.date) dateSet.add(r.date.trim());
      }});

      const dates = Array.from(dateSet).sort((a, b) => {{
        const pa = parseReceiveDate(a).sortKey;
        const pb = parseReceiveDate(b).sortKey;
        return pa.localeCompare(pb);
      }});

      const todayStr = getTodayInfo().ddmmyyyy;
      let opts = '<option value="ALL">📅 Tất cả ngày</option>';
      if (dateSet.has(todayStr)) {{
        opts += `<option value="${{todayStr}}">⚡ Hôm nay (${{todayStr}})</option>`;
      }}
      dates.forEach(d => {{
        if (d !== todayStr) {{
          opts += `<option value="${{d}}">${{d}}</option>`;
        }}
      }});
      select.innerHTML = opts;
      select.value = planModalDate || 'ALL';
    }}

    function onPlanModalFilterChange() {{
      const mSel = document.getElementById('planModalMonthSelect');
      const dSel = document.getElementById('planModalDateSelect');
      if (mSel) planModalMonth = mSel.value;
      if (dSel) planModalDate = dSel.value;
      renderPlanVolumeModal();
    }}

    function setPlanModalToday() {{
      const todayStr = getTodayInfo().ddmmyyyy;
      planModalDate = todayStr;
      const dSel = document.getElementById('planModalDateSelect');
      if (dSel) dSel.value = todayStr;
      renderPlanVolumeModal();
    }}

    function resetPlanModalFilter() {{
      planModalDate = 'ALL';
      planModalMonth = 'ALL';
      const mSel = document.getElementById('planModalMonthSelect');
      const dSel = document.getElementById('planModalDateSelect');
      if (mSel) mSel.value = 'ALL';
      if (dSel) dSel.value = 'ALL';
      renderPlanVolumeModal();
    }}

    function openPlanVolumeModal() {{
      const modal = document.getElementById('planVolumeModal');
      if (!modal) return;
      
      // Đồng bộ filter ban đầu từ Dashboard
      if (selectedDate && selectedDate !== 'ALL') {{
        planModalDate = selectedDate;
      }} else {{
        planModalDate = 'ALL';
      }}
      if (selectedMonth && selectedMonth !== 'ALL') {{
        planModalMonth = selectedMonth;
      }} else {{
        planModalMonth = 'ALL';
      }}

      populatePlanModalDateSelect();
      const mSel = document.getElementById('planModalMonthSelect');
      if (mSel) mSel.value = planModalMonth;
      const dSel = document.getElementById('planModalDateSelect');
      if (dSel) dSel.value = planModalDate;

      modal.style.display = 'flex';
      document.body.style.overflow = 'hidden';
      renderPlanVolumeModal();
    }}

    function closePlanVolumeModal(e) {{
      if (e && e.target && e.target.id !== 'planVolumeModal' && !e.target.closest('button')) return;
      const modal = document.getElementById('planVolumeModal');
      if (modal) {{
        modal.style.display = 'none';
        document.body.style.overflow = '';
      }}
    }}

    function renderPlanVolumeModal() {{
      const listEl = document.getElementById('modalPlanVolumeList');
      const subEl = document.getElementById('modal-plan-volume-subtitle');
      const headerEl = document.getElementById('modalPlanTableHeader');
      const totalWhEl = document.getElementById('modalPlanTotalWh');
      const totalSkuEl = document.getElementById('modalPlanTotalSku');
      const totalQtyEl = document.getElementById('modalPlanTotalQty');
      if (!listEl) return;

      const planMtiData = currentData.plan_mti || [];
      const todayStr = getTodayInfo().ddmmyyyy;

      // Lọc dữ liệu theo tháng và ngày
      const today = getTodayInfo();
      const filtered = planMtiData.filter(r => {{
        if (!r.date) return false;
        const p = parseReceiveDate(r.date);
        if (planModalDate && planModalDate !== 'ALL') {{
          return p.ddmmyyyy === planModalDate;
        }}
        if (planModalMonth && planModalMonth !== 'ALL') {{
          return p.mmyyyy === planModalMonth && (p.sortKey <= today.sortKey || p.mmyyyy < today.mmyyyy);
        }}
        // Nếu không chọn gì: Từ ngày 1 tháng hiện tại → ngày D
        const curMonthKeyM = today.sortKey.substring(0, 6);
        const startOfCurMonthM = curMonthKeyM + '01';
        return p.sortKey >= startOfCurMonthM && p.sortKey <= today.sortKey;
      }});

      // Tính tổng QTY lọc được
      let grandQty = 0;
      filtered.forEach(r => {{
        grandQty += (parseInt(r.qty, 10) || 0);
      }});

      // Cập nhật subtitle
      let filterDesc = 'Tất cả các ngày';
      if (planModalDate && planModalDate !== 'ALL') {{
        filterDesc = (planModalDate === todayStr) ? `Hôm nay: ${{planModalDate}}` : `Ngày ${{planModalDate}}`;
      }} else if (planModalMonth && planModalMonth !== 'ALL') {{
        filterDesc = `Tháng ${{planModalMonth}}`;
      }}
      if (subEl) {{
        subEl.innerHTML = `Nguồn: Sheet <b>Data Plan MTI</b> &bull; <span style="color: #4ade80; background: rgba(34,197,94,0.15); padding: 2px 7px; border-radius: 4px; border: 1px solid rgba(74,222,128,0.3); font-weight: 700;">${{filterDesc}}</span>`;
      }}

      // ==========================================
      // CHẾ ĐỘ 1: THEO KHO CHUYỂN (Warehouse Mode)
      // ==========================================
      if (planModalViewMode === 'wh') {{
        if (headerEl) {{
          headerEl.innerHTML = `
            <span>STT</span>
            <span>🏭 KHO CHUYỂN</span>
            <span style="text-align: right;">📅 SỐ NGÀY PLAN</span>
            <span style="text-align: right;">📊 SỐ LƯỢNG (QTY)</span>
            <span style="text-align: right;">TỶ TRỌNG</span>
          `;
        }}

        const whMap = {{}};
        filtered.forEach(r => {{
          let wh = (r.warehouse || 'Chưa xác định').trim().toUpperCase();
          if (wh === 'VNN') wh = 'VNNL';
          const q = parseInt(r.qty, 10) || 0;
          if (!whMap[wh]) whMap[wh] = {{ wh: wh, qty: 0, dates: new Set() }};
          whMap[wh].qty += q;
          if (r.date) whMap[wh].dates.add(r.date.trim());
        }});

        const whList = Object.values(whMap).sort((a, b) => b.qty - a.qty);

        if (whList.length === 0) {{
          listEl.innerHTML = `<div style="text-align: center; padding: 40px; color: var(--text-dim);">Không có kế hoạch MTI cho ${{filterDesc}}.</div>`;
          if (totalWhEl) totalWhEl.textContent = '0 kho chuyển';
          if (totalSkuEl) totalSkuEl.textContent = '-';
          if (totalQtyEl) totalQtyEl.textContent = 'Tổng: 0 QTY';
          return;
        }}

        let html = '';
        whList.forEach((item, idx) => {{
          const pct = grandQty > 0 ? ((item.qty / grandQty) * 100).toFixed(1) : 0;
          html += `
            <div style="display: grid; grid-template-columns: 50px 1fr 140px 120px 80px; align-items: center; padding: 10px 12px; background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.06); border-radius: 8px; transition: all 0.2s ease;">
              <span style="color: var(--text-dim); font-weight: 700;">${{idx + 1}}</span>
              <strong style="color: #f8fafc; display: flex; align-items: center; gap: 8px;">
                <span style="color: #818cf8;">🏢</span> ${{item.wh}}
              </strong>
              <span style="text-align: right; color: #38bdf8; font-weight: 700;">${{item.dates.size}} ngày</span>
              <span style="text-align: right; color: #4ade80; font-weight: 800; font-family: 'JetBrains Mono', monospace;">${{fmtNum(item.qty)}}</span>
              <span style="text-align: right; color: #f59e0b; font-weight: 700;">${{pct}}%</span>
            </div>
          `;
        }});

        listEl.innerHTML = html;
        if (totalWhEl) totalWhEl.textContent = `${{whList.length}} kho chuyển`;
        if (totalSkuEl) totalSkuEl.textContent = '-';
        if (totalQtyEl) totalQtyEl.textContent = `Tổng: ${{fmtNum(grandQty)}} QTY`;

      // ==========================================
      // CHẾ ĐỘ 2: THEO TỪNG NGÀY (Daily Plan Mode)
      // ==========================================
      }} else if (planModalViewMode === 'daily') {{
        if (headerEl) {{
          headerEl.innerHTML = `
            <span>STT</span>
            <span>📅 NGÀY ĐẾN (PLAN)</span>
            <span style="text-align: center;">🏢 PHÂN BỔ KHO</span>
            <span style="text-align: right;">📊 TỔNG QTY (NGÀY)</span>
            <span style="text-align: right;">TỶ TRỌNG</span>
          `;
        }}

        const dailyMap = {{}};
        filtered.forEach(r => {{
          const d = (r.date || '').trim();
          if (!d) return;
          let wh = (r.warehouse || 'Chưa xác định').trim().toUpperCase();
          if (wh === 'VNN') wh = 'VNNL';
          const q = parseInt(r.qty, 10) || 0;
          if (!dailyMap[d]) {{
            dailyMap[d] = {{ date: d, qty: 0, whs: {{}} }};
          }}
          dailyMap[d].qty += q;
          dailyMap[d].whs[wh] = (dailyMap[d].whs[wh] || 0) + q;
        }});

        // Sắp xếp ngày theo thứ tự thời gian
        const dailyList = Object.values(dailyMap).sort((a, b) => {{
          const pa = parseReceiveDate(a.date).sortKey;
          const pb = parseReceiveDate(b.date).sortKey;
          return pa.localeCompare(pb);
        }});

        if (dailyList.length === 0) {{
          listEl.innerHTML = `<div style="text-align: center; padding: 40px; color: var(--text-dim);">Không có kế hoạch MTI cho ${{filterDesc}}.</div>`;
          if (totalWhEl) totalWhEl.textContent = '0 ngày';
          if (totalSkuEl) totalSkuEl.textContent = '-';
          if (totalQtyEl) totalQtyEl.textContent = 'Tổng: 0 QTY';
          return;
        }}

        let html = '';
        dailyList.forEach((item, idx) => {{
          const pct = grandQty > 0 ? ((item.qty / grandQty) * 100).toFixed(1) : 0;
          const isToday = (item.date === todayStr);
          const todayBadge = isToday ? ` <span style="font-size: 10px; color: #4ade80; background: rgba(34,197,94,0.2); padding: 1px 5px; border-radius: 4px; font-weight: 800; border: 1px solid rgba(74,222,128,0.4);">HÔM NAY</span>` : '';
          
          // Tạo badge phân bổ các kho
          const whBadges = Object.entries(item.whs).map(([w, q]) => {{
            return `<span style="font-size: 10px; background: rgba(56, 189, 248, 0.12); color: #38bdf8; border: 1px solid rgba(56, 189, 248, 0.25); padding: 1px 4px; border-radius: 4px; white-space: nowrap;">${{w}}: <b>${{fmtNum(q)}}</b></span>`;
          }}).join(' ');

          html += `
            <div onclick="selectPlanDate('${{item.date}}')" style="display: grid; grid-template-columns: 50px 160px 1fr 120px 80px; align-items: center; padding: 10px 12px; background: ${{isToday ? 'rgba(34,197,94,0.06)' : 'rgba(255,255,255,0.03)'}}; border: 1px solid ${{isToday ? 'rgba(74,222,128,0.3)' : 'rgba(255,255,255,0.06)'}}; border-radius: 8px; cursor: pointer; transition: all 0.2s ease;" title="Click để xem chi tiết kho của ngày này">
              <span style="color: var(--text-dim); font-weight: 700;">${{idx + 1}}</span>
              <strong style="color: ${{isToday ? '#4ade80' : '#f8fafc'}}; display: flex; align-items: center; gap: 4px;">
                📅 ${{item.date}}${{todayBadge}}
              </strong>
              <div style="display: flex; gap: 4px; flex-wrap: wrap; justify-content: center;">
                ${{whBadges}}
              </div>
              <span style="text-align: right; color: #4ade80; font-weight: 800; font-family: 'JetBrains Mono', monospace;">${{fmtNum(item.qty)}}</span>
              <span style="text-align: right; color: #f59e0b; font-weight: 700;">${{pct}}%</span>
            </div>
          `;
        }});

        listEl.innerHTML = html;
        if (totalWhEl) totalWhEl.textContent = `${{dailyList.length}} ngày có plan`;
        if (totalSkuEl) totalSkuEl.textContent = '-';
        if (totalQtyEl) totalQtyEl.textContent = `Tổng: ${{fmtNum(grandQty)}} QTY`;

      // ==========================================
      // CHẾ ĐỘ 3: THEO THÁNG (Monthly Plan Mode)
      // ==========================================
      }} else if (planModalViewMode === 'monthly') {{
        if (headerEl) {{
          headerEl.innerHTML = `
            <span>STT</span>
            <span>🗓️ THÁNG</span>
            <span style="text-align: center;">📅 SỐ NGÀY PLAN</span>
            <span style="text-align: right;">📊 TỔNG QTY (THÁNG)</span>
            <span style="text-align: right;">TỶ TRỌNG</span>
          `;
        }}

        const monthMap = {{}};
        // Tính trên toàn bộ dataset plan
        planMtiData.forEach(r => {{
          const d = (r.date || '').trim();
          if (!d) return;
          const p = parseReceiveDate(d);
          const m = p.mmyyyy || 'Khác';
          const q = parseInt(r.qty, 10) || 0;
          if (!monthMap[m]) {{
            monthMap[m] = {{ month: m, qty: 0, dates: new Set(), sortKey: p.sortKey.substring(0, 6) }};
          }}
          monthMap[m].qty += q;
          monthMap[m].dates.add(d);
        }});

        const monthList = Object.values(monthMap).sort((a, b) => a.sortKey.localeCompare(b.sortKey));
        const totalSeasonQty = monthList.reduce((sum, item) => sum + item.qty, 0);

        let html = '';
        monthList.forEach((item, idx) => {{
          const pct = totalSeasonQty > 0 ? ((item.qty / totalSeasonQty) * 100).toFixed(1) : 0;
          const isCurrent = (planModalMonth === item.month);
          html += `
            <div onclick="selectPlanMonth('${{item.month}}')" style="display: grid; grid-template-columns: 50px 160px 1fr 120px 80px; align-items: center; padding: 12px 14px; background: ${{isCurrent ? 'rgba(99,102,241,0.12)' : 'rgba(255,255,255,0.03)'}}; border: 1px solid ${{isCurrent ? '#6366f1' : 'rgba(255,255,255,0.06)'}}; border-radius: 8px; cursor: pointer; transition: all 0.2s ease;" title="Click để xem kế hoạch theo ngày của tháng này">
              <span style="color: var(--text-dim); font-weight: 700;">${{idx + 1}}</span>
              <strong style="color: #f8fafc; font-size: 14px; display: flex; align-items: center; gap: 8px;">
                🗓️ Tháng ${{item.month}}
              </strong>
              <span style="text-align: center; color: #38bdf8; font-weight: 700;">${{item.dates.size}} ngày</span>
              <span style="text-align: right; color: #4ade80; font-weight: 800; font-size: 14px; font-family: 'JetBrains Mono', monospace;">${{fmtNum(item.qty)}}</span>
              <span style="text-align: right; color: #f59e0b; font-weight: 700;">${{pct}}%</span>
            </div>
          `;
        }});

        listEl.innerHTML = html;
        if (totalWhEl) totalWhEl.textContent = `${{monthList.length}} tháng`;
        if (totalSkuEl) totalSkuEl.textContent = '-';
        if (totalQtyEl) totalQtyEl.textContent = `Tổng toàn mùa: ${{fmtNum(totalSeasonQty)}} QTY`;
      }}
    }}

    function selectPlanDate(dateStr) {{
      planModalDate = dateStr;
      const dSel = document.getElementById('planModalDateSelect');
      if (dSel) dSel.value = dateStr;
      setPlanModalViewMode('wh');
    }}

    function selectPlanMonth(monthStr) {{
      planModalMonth = monthStr;
      planModalDate = 'ALL';
      const mSel = document.getElementById('planModalMonthSelect');
      const dSel = document.getElementById('planModalDateSelect');
      if (mSel) mSel.value = monthStr;
      if (dSel) dSel.value = 'ALL';
      setPlanModalViewMode('daily');
    }}

    function openActualPutawayModal() {{
      const modal = document.getElementById('actualPutawayModal');
      if (modal) {{
        modal.style.display = 'flex';
        document.body.style.overflow = 'hidden';
        const searchInput = document.getElementById('modalOperatorSearch');
        if (searchInput) {{
          searchInput.value = '';
          setTimeout(() => searchInput.focus(), 50);
        }}
        renderModalOperators();
      }}
    }}

    function closeActualPutawayModal(e) {{
      if (e && e.target && e.target.id !== 'actualPutawayModal' && !e.target.closest('button')) return;
      const modal = document.getElementById('actualPutawayModal');
      if (modal) {{
        modal.style.display = 'none';
        document.body.style.overflow = '';
      }}
    }}

    function renderModalOperators() {{
      const listEl = document.getElementById('modalOperatorList');
      if (!listEl) return;
      const q = (document.getElementById('modalOperatorSearch')?.value || '').trim().toLowerCase();

      const records = getFilteredActualPutaway();
      const opMap = {{}};
      let totalQty = 0;
      for (const r of records) {{
        const op = (r.op || 'Unknown').trim();
        if (!opMap[op]) {{
          opMap[op] = {{ op: op, qty: 0, cnt: 0 }};
        }}
        opMap[op].qty += Number(r.qty || 0);
        opMap[op].cnt += Number(r.cnt || 0);
        totalQty += Number(r.qty || 0);
      }}

      let ops = Object.values(opMap);
      ops.sort((a, b) => b.qty - a.qty);

      if (q) {{
        ops = ops.filter(item => item.op.toLowerCase().includes(q));
      }}

      let subText = 'Tất cả các ngày & tháng';
      if (selectedDate !== 'ALL') {{
        subText = `Ngày ${{selectedDate}}`;
      }} else if (selectedMonth !== 'ALL') {{
        subText = `Tháng ${{selectedMonth}}`;
      }}
      const sourceName = (currentBucket === 'Inbound') ? 'Put IB' : 'Put MTI';
      const subEl = document.getElementById('modal-putaway-subtitle');
      if (subEl) subEl.textContent = `Thống kê năng suất từ ${{sourceName}} • ${{subText}}`;

      if (ops.length === 0) {{
        listEl.innerHTML = '<div style="text-align: center; color: var(--text-dim); padding: 30px;">Không tìm thấy operator nào.</div>';
        document.getElementById('modalTotalOps').textContent = '0 operators';
        document.getElementById('modalTotalQty').textContent = 'Total: 0 QTY';
        return;
      }}

      let html = '';
      ops.forEach((item, idx) => {{
        let dot = '🟡';
        if (idx < 4 || item.qty >= 3000) dot = '🔴';
        else if (idx < 7 || item.qty >= 1500) dot = '🟠';

        html += `
          <div style="display: flex; justify-content: space-between; align-items: center; padding: 7px 6px; border-bottom: 1px dashed rgba(255,255,255,0.08); font-size: 13px;">
            <div style="display: flex; align-items: center; gap: 8px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 65%;">
              <span style="font-size: 11px;">${{dot}}</span>
              <span style="color: #f1f5f9; text-overflow: ellipsis; overflow: hidden;" title="${{item.op}}">${{item.op}}</span>
            </div>
            <div style="color: #94a3b8; font-weight: 600; white-space: nowrap;">
              : &nbsp;<span style="color: #f8fafc; font-weight: 700;">${{fmtNum(item.qty)}}</span> &nbsp;|&nbsp; <span style="color: #64748b;">${{fmtNum(item.cnt)}}</span>
            </div>
          </div>
        `;
      }});

      listEl.innerHTML = html;
      document.getElementById('modalTotalOps').textContent = `${{ops.length}} operators`;
      document.getElementById('modalTotalQty').textContent = `Total: ${{fmtNum(totalQty)}} QTY`;
    }}

    let currentTruckFilter = 'all';

    function openTruckTrackerModal(filterType = 'all') {{
      currentTruckFilter = filterType || 'all';
      const modal = document.getElementById('truckTrackerModal');
      if (modal) {{
        modal.style.display = 'flex';
        document.body.style.overflow = 'hidden';
        renderTruckTrackerModal();
      }}
    }}

    function filterTruckModal(type) {{
      currentTruckFilter = type;
      renderTruckTrackerModal();
    }}

    async function closeTruckTrackerModal(e) {{
      if (e && e.target && e.target.id !== 'truckTrackerModal' && !e.target.closest('button')) return;
      
      // Tự động lưu tất cả thẻ xe đang có dữ liệu số Pallet / Type / Shipping ID khi bấm Đóng
      const forms = document.querySelectorAll('[id^="card_pallet_"]');
      for (const pInp of forms) {{
        const cardId = pInp.id.replace('card_pallet_', '');
        const pVal = (pInp.value || '').trim();
        const sVal = (document.getElementById('card_ship_' + cardId)?.value || '').trim();
        const tVal = (document.getElementById('card_type_' + cardId)?.value || '').trim();
        if (pVal || sVal || tVal) {{
          try {{
            await saveXeDenFromCard(cardId, false);
          }} catch (err) {{}}
        }}
      }}

      const modal = document.getElementById('truckTrackerModal');
      if (modal) {{
        modal.style.display = 'none';
        document.body.style.overflow = '';
      }}
      updateUI();
    }}

    function getTruckVal(row, possibleKeys) {{
      if (!row) return '';
      for (const k of possibleKeys) {{
        for (const rowKey of Object.keys(row)) {{
          if (rowKey.trim().toLowerCase() === k.toLowerCase()) {{
            return String(row[rowKey] || '').trim();
          }}
        }}
      }}
      return '';
    }}

    let truckModalCustomDate = '';

    function changeTruckModalDate(offsetDays) {{
      const cur = truckModalCustomDate || (selectedDate && selectedDate !== 'ALL' ? selectedDate : getTodayInfo().ddmmyyyy);
      const dt = parseDdmmyyyyToDate(cur);
      if (dt) {{
        dt.setDate(dt.getDate() + offsetDays);
        const d = String(dt.getDate()).padStart(2, '0');
        const m = String(dt.getMonth() + 1).padStart(2, '0');
        const y = dt.getFullYear();
        truckModalCustomDate = `${{d}}/${{m}}/${{y}}`;
        renderTruckTrackerModal();
      }}
    }}

    function setTruckModalDateToday() {{
      truckModalCustomDate = getTodayInfo().ddmmyyyy;
      renderTruckTrackerModal();
    }}

    function renderTruckTrackerModal() {{
      const listEl = document.getElementById('truckTripsList');
      const formEl = document.getElementById('truckManualForm');
      const dateEl = document.getElementById('truckDateDisplay');
      const totalEl = document.getElementById('truckTotalDisplay');
      const summaryEl = document.getElementById('truckArrivedSummary');
      const tabsEl = document.getElementById('truckModalTabs');
      if (!listEl) return;

      const truckData = currentData.truck_tracker || {{ columns: [], items: [], total: 0, xe_den_vndl: [] }};
      const allItems = truckData.items || [];
      const xeDenList = truckData.xe_den_vndl || [];

      // Xác định ngày hiển thị: ngày D (hỗ trợ chuyển ngày nhanh trong modal)
      let targetDate = '';
      if (truckModalCustomDate) {{
        targetDate = truckModalCustomDate;
      }} else if (selectedDate && selectedDate !== 'ALL') {{
        targetDate = selectedDate;
      }} else {{
        const today = getTodayInfo();
        targetDate = today.ddmmyyyy;
      }}

      // Ngày kế hoạch xuất phát D-1
      const planTruckDate = getTruckPlanDate(targetDate);
      if (dateEl) {{
        dateEl.textContent = `${{targetDate}} (Kế hoạch xuất phát D-1: ${{planTruckDate}})`;
      }}

      // Lọc các chuyến kế hoạch cho ngày D: xuất phát D-1 hoặc dự kiến đến D (hoặc ngày MTO = D)
      function _matchDate(r) {{
        // Kiểm tra kho Đà Nẵng qua 3 cột: Pick thêm, Kho đến, Drop thêm
        const pTh = getTruckVal(r, ['Pick thêm', 'Pick them']).toLowerCase();
        const kDen = getTruckVal(r, ['Kho đến', 'Kho den']).toLowerCase();
        const dTh = getTruckVal(r, ['Drop thêm', 'Drop them']).toLowerCase();
        const hasDaNang = (pTh.includes('đà nẵng') || pTh.includes('da nang') || pTh.includes('vnd') ||
                           kDen.includes('đà nẵng') || kDen.includes('da nang') || kDen.includes('vnd') ||
                           dTh.includes('đà nẵng') || dTh.includes('da nang') || dTh.includes('vnd'));
        if (!hasDaNang) return false;

        if (!targetDate || targetDate === 'ALL') return true;
        const d1 = getTruckVal(r, ['Ngày MTO', 'Ngay MTO', 'Ngày', 'Ngay', 'Date']);
        const p1 = d1 ? parseReceiveDate(d1).ddmmyyyy : '';
        // Kế hoạch xe đến ngày D được lấy từ Ngày MTO = D - 1 (leadtime 24h)
        if (p1 && p1 === planTruckDate) return true;   // MTO = D-1
        return false;
      }}

      let trips = allItems.filter(_matchDate);
      window._currentTruckTrips = trips;

      // Lọc xe đến VNDL thực tế: xe đến vào đúng ngày D
      const dayXeDen = xeDenList.filter(x => {{
        const xDate = getTruckVal(x, ['Dấu thời gian', 'Thời gian', 'Ngày', 'Ngay', 'date', 'ngay', 'Time']);
        if (!xDate) return false;
        if (!targetDate || targetDate === 'ALL') return true;
        const xd = parseReceiveDate(xDate).ddmmyyyy;
        return xd === targetDate;
      }});

      // Đối chiếu từng chuyến kế hoạch với xe đến thực tế
      let arrivedTrips = [];
      let notArrivedTrips = [];
      const matchedXeDenSet = new Set();

      trips.forEach((r, idx) => {{
        const p1 = cleanTruckPlate(getTruckVal(r, ['Số xe', 'So xe', 'Biển số', 'so_xe', 'Biển số xe']));
        const ph1 = cleanTruckPhone(getTruckVal(r, ['Thông tin tài xế', 'SĐT', 'Số điện thoại', 'sdt', 'Phone', 'Số ĐT']));
        const matchedXd = dayXeDen.find(xd => {{
          const p2 = cleanTruckPlate(getTruckVal(xd, ['Biển số xe', 'Biển số', 'Số xe', 'So xe', 'so_xe']));
          const ph2 = cleanTruckPhone(getTruckVal(xd, ['Số ĐT', 'Số điện thoại', 'SĐT', 'sdt', 'Phone', 'Thông tin tài xế']));
          if (p1 && p2 && (p1.includes(p2) || p2.includes(p1))) return true;
          if (ph1 && ph2 && ph1.length >= 8 && ph2.length >= 8 && (ph1.includes(ph2) || ph2.includes(ph1))) return true;
          return false;
        }});

        r._isArrived = !!matchedXd;
        r._matchedXeDen = matchedXd || null;

        // Tính Leadtime: So sánh ngày xe đến thực tế với ngày kế hoạch D
        if (matchedXd) {{
          const arrDateStr = getTruckVal(matchedXd, ['Dấu thời gian', 'Thời gian', 'Ngày', 'Ngay', 'date', 'ngay']);
          const arrP = parseReceiveDate(arrDateStr);
          const tgtP = parseReceiveDate(targetDate);
          if (arrP.sortKey && tgtP.sortKey) {{
            r._leadtime = (arrP.sortKey > tgtP.sortKey) ? 'Late' : 'On Time';
          }} else {{
            r._leadtime = 'On Time';
          }}
        }} else {{
          r._leadtime = 'Pending';
        }}

        if (r._isArrived) {{
          matchedXeDenSet.add(matchedXd);
          arrivedTrips.push(r);
        }} else {{
          notArrivedTrips.push(r);
        }}
      }});

      // Thêm các xe thực tế đến nhưng ngoài kế hoạch
      const unplanArrived = dayXeDen.filter(xd => !matchedXeDenSet.has(xd)).map(xd => {{
        const arrDateStr = getTruckVal(xd, ['Dấu thời gian', 'Thời gian', 'Ngày', 'Ngay', 'date', 'ngay']);
        const arrP = parseReceiveDate(arrDateStr);
        const tgtP = parseReceiveDate(targetDate);
        let leadtime = 'On Time';
        if (arrP.sortKey && tgtP.sortKey && arrP.sortKey > tgtP.sortKey) {{
          leadtime = 'Late';
        }}
        return {{
          _isArrived: true,
          _isUnplanned: true,
          _leadtime: leadtime,
          _matchedXeDen: xd,
          'Số xe': getTruckVal(xd, ['Biển số xe', 'Biển số', 'Số xe', 'so_xe']),
          'Thông tin tài xế': `${{getTruckVal(xd, ['Name', 'tai_xe', 'Tài xế']) || ''}} ${{getTruckVal(xd, ['Số ĐT', 'Số điện thoại', 'sdt']) ? `(${{getTruckVal(xd, ['Số ĐT', 'Số điện thoại', 'sdt'])}})` : ''}}`.trim(),
          'Kho xuất phát': getTruckVal(xd, ['Kho xuất phát', 'IN From WH', 'kho_di']),
          'Loại xe': getTruckVal(xd, ['Loại xe', 'IN Type']),
          'Pallet thực tế': getTruckVal(xd, ['IN Pallet', 'Pallet', 'pallet']) || 0,
          'Thời gian': getTruckVal(xd, ['Dấu thời gian', 'Thời gian', 'gio_den']),
          'Ghi chú': getTruckVal(xd, ['IN LDD', 'IN Operater', 'ghi_chu'])
        }};
      }});

      const allArrived = arrivedTrips.concat(unplanArrived);
      const totalAll = trips.length + unplanArrived.length;

      if (totalEl) totalEl.textContent = totalAll;
      if (summaryEl) {{
        summaryEl.innerHTML = `
          <span style="color: #4ade80;">✓ ${{allArrived.length}} xe đã đến</span>
          <span style="color: rgba(255,255,255,0.3); margin: 0 4px;">/</span>
          <span style="color: #fb7185;">⏳ ${{notArrivedTrips.length}} xe chưa đến</span>
        `;
      }}

      // Render Tabs
      if (tabsEl) {{
        tabsEl.innerHTML = `
          <button type="button" class="filter-tab ${{currentTruckFilter === 'all' ? 'active' : ''}}" onclick="filterTruckModal('all')" style="font-size: 11px; padding: 4px 10px; cursor: pointer;">Tất cả (${{totalAll}})</button>
          <button type="button" class="filter-tab ${{currentTruckFilter === 'arrived' ? 'active' : ''}}" onclick="filterTruckModal('arrived')" style="font-size: 11px; padding: 4px 10px; cursor: pointer; color: #4ade80; font-weight: 700;">✓ Xe đã đến (${{allArrived.length}})</button>
          <button type="button" class="filter-tab ${{currentTruckFilter === 'not_arrived' ? 'active' : ''}}" onclick="filterTruckModal('not_arrived')" style="font-size: 11px; padding: 4px 10px; cursor: pointer; color: #fb7185; font-weight: 700;">⏳ Xe chưa đến (${{notArrivedTrips.length}})</button>
        `;
      }}

      // Nếu đang mở tab nhập thủ công
      if (currentTruckFilter === 'new') {{
        listEl.style.display = 'none';
        if (formEl) {{
          formEl.style.display = 'block';
          const inPlate = document.getElementById('manual_so_xe');
          if (inPlate) setTimeout(() => inPlate.focus(), 50);
        }}
        return;
      }}

      // Hiển thị danh sách chuyến
      if (formEl) formEl.style.display = 'none';
      listEl.style.display = 'flex';

      let displayTrips = (unplanArrived.length > 0) ? trips.concat(unplanArrived) : trips;
      if (currentTruckFilter === 'arrived') {{
        displayTrips = allArrived;
      }} else if (currentTruckFilter === 'not_arrived') {{
        displayTrips = notArrivedTrips;
      }}

      if (displayTrips.length === 0) {{
        let emptyMsg = 'Không có chuyến xe nào trong ngày này.';
        if (currentTruckFilter === 'arrived') emptyMsg = 'Chưa có xe nào đến trong ngày này.';
        if (currentTruckFilter === 'not_arrived') emptyMsg = 'Tất cả các xe trong kế hoạch hôm nay đều đã đến!';
        listEl.innerHTML = `
          <div style="text-align: center; padding: 35px 15px; color: var(--text-dim);">
            <div style="font-size: 32px; margin-bottom: 8px;">🚚</div>
            <div>${{emptyMsg}}</div>
          </div>
        `;
        return;
      }}

      let html = '';
      displayTrips.forEach((r, idx) => {{
        if (r._isUnplanned) {{
          const mXd = r._matchedXeDen || r;
          const plt = getTruckVal(mXd, ['IN Pallet', 'Pallet', 'pallet']) || r['Pallet thực tế'] || 0;
          const timeStr = getTruckVal(mXd, ['Dấu thời gian', 'Thời gian', 'gio_den']) || r['Thời gian'] || 'Đã vào kho';
          const plate = r['Số xe'] || 'Chưa rõ biển số';
          const driverInfo = r['Thông tin tài xế'] || 'Chưa rõ tài xế';
          const fromWh = r['Kho xuất phát'] || 'VNDL';
          const truckType = r['Loại xe'] || '-';
          const ldd = r['Ghi chú'] || '';

          const curMode = getTruckVal(mXd, ['OUT Pallet', 'OUT Type']) ? 'OUT' : 'IN';
          const curType = getTruckVal(mXd, [curMode === 'IN' ? 'IN Type' : 'OUT Type', 'Loại xe']) || '';
          const curPallet = getTruckVal(mXd, [curMode === 'IN' ? 'IN Pallet' : 'OUT Pallet', 'Pallet', 'pallet']) || '';
          const curWh = getTruckVal(mXd, [curMode === 'IN' ? 'IN From WH' : 'OUT To WH', 'Kho xuất phát', 'From WH']) || fromWh || '';
          const curOp = getTruckVal(mXd, [curMode === 'IN' ? 'IN Operater' : 'OUT Operater', 'IN Operator', 'OUT Operator']) || '';
          const curShipId = getTruckVal(mXd, ['Shipping ID', 'Shiping ID', 'ShippingId']) || '';
          const rawDriverName = (getTruckVal(mXd, ['Name', 'tai_xe', 'Tài xế']) || '').replace(/[\x27\x22]/g, ' ').replace(/\s+/g, ' ').trim();
          const rawDriverPhone = (getTruckVal(mXd, ['Số ĐT', 'Số điện thoại', 'sdt']) || '').replace(/[\x27\x22]/g, ' ').replace(/\s+/g, ' ').trim();

          const _inData = {{ type: getTruckVal(mXd, ['IN Type']) || '', pallet: getTruckVal(mXd, ['IN Pallet']) || '', wh: getTruckVal(mXd, ['IN From WH']) || '', op: getTruckVal(mXd, ['IN Operater','IN Operator']) || '' }};
          const _outData = {{ type: getTruckVal(mXd, ['OUT Type']) || '', pallet: getTruckVal(mXd, ['OUT Pallet']) || '', wh: getTruckVal(mXd, ['OUT To WH']) || '', op: getTruckVal(mXd, ['OUT Operater','OUT Operator']) || '' }};
          const formHtml = getCardXeDenFormHtml(`card_unplan_${{idx}}`, curMode, curType, curPallet, curWh, curOp, curShipId, plate, rawDriverName, rawDriverPhone, fromWh, _inData, _outData);

          html += `
            <div style="background: rgba(34, 197, 94, 0.05); border: 1.5px solid rgba(34, 197, 94, 0.35); border-radius: 10px; padding: 14px 16px;">
              <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 6px;">
                <div style="font-size: 14px; font-weight: 800; color: #4ade80; display: flex; align-items: center; gap: 6px;">
                  <span>🚚</span> XE ĐẾN THỰC TẾ (Ngoài kế hoạch)
                </div>
                <div style="display: inline-flex; align-items: center; gap: 8px;">
                  <div style="display: inline-flex; align-items: center; gap: 6px; background: rgba(74, 222, 128, 0.15); border: 1px solid #16a34a; border-radius: 6px; padding: 4px 10px; color: #4ade80; font-weight: 800; font-size: 11px;">
                    <span>✓ ĐÃ ĐẾN</span> &bull; <span>📦 ${{plt}} Pallet</span> &bull; <span>⏰ ${{timeStr}}</span>
                  </div>
                  <span style="background: rgba(34, 197, 94, 0.2); color: #4ade80; border: 1px solid rgba(74, 222, 128, 0.4); padding: 3px 8px; border-radius: 4px; font-weight: 800; font-size: 11px;">⚡ On Time</span>
                </div>
              </div>
              <div style="display: flex; gap: 16px; justify-content: space-between; align-items: flex-start; flex-wrap: wrap;">
                <div style="display: flex; flex-direction: column; gap: 5px; font-size: 13px; flex: 1; min-width: 250px;">
                  <div><span style="display: inline-block; width: 140px; color: #94a3b8;">🔴 Kho xuất phát</span>: <strong style="color: #f8fafc;">${{getKhoDisplayName(fromWh)}}</strong></div>
                  <div><span style="display: inline-block; width: 140px; color: #94a3b8;">🟢 Biển số xe</span>: <strong style="color: #38bdf8;">${{plate}}</strong></div>
                  <div><span style="display: inline-block; width: 140px; color: #94a3b8;">🔵 Loại xe</span>: <span style="color: #cbd5e1;">${{truckType}}</span></div>
                  <div><span style="display: inline-block; width: 140px; color: #94a3b8;">👤 Tài xế</span>: <strong style="color: #f8fafc;">${{driverInfo}}</strong></div>
                  <div><span style="display: inline-block; width: 140px; color: #94a3b8;">⏱️ Leadtime</span>: <span style="color: #4ade80; font-weight: 700;">On Time</span></div>
                  ${{ldd ? `<div><span style="display: inline-block; width: 140px; color: #94a3b8;">📝 LDD / Ghi chú</span>: <span style="color: #cbd5e1;">${{ldd}}</span></div>` : ''}}
                </div>
                ${{formHtml}}
              </div>
            </div>
          `;
          return;
        }}

        const chuyenVal = getTruckVal(r, ['Số chuyến', 'So chuyen', 'Chuyến', 'Chuyen']);
        const chuyenNum = chuyenVal || (idx + 1);
        const khoDi = getTruckVal(r, ['Kho đi', 'Kho di', 'From', 'Kho xuất']);
        const pickThem = getTruckVal(r, ['Pick thêm', 'Pick them', 'Pick']);
        const khoDen = getTruckVal(r, ['Kho đến', 'Kho den', 'To', 'Kho nhập']);
        const dropThem = getTruckVal(r, ['Drop thêm', 'Drop them', 'Drop']);
        const soXe = getTruckVal(r, ['Số xe', 'So xe', 'Biển số', 'Bien so', 'License Plate']);
        const loaiXe = getTruckVal(r, ['Loại xe', 'Loai xe', 'Truck Type', 'Type']);
        const vendor = getTruckVal(r, ['Vendor', 'Nhà xe', 'Nha xe', 'Đơn vị vận chuyển']);
        const taiXe = getTruckVal(r, ['Thông tin tài xế', 'Thong tin tai xe', 'Tài xế', 'Tai xe', 'Driver', 'Tài xế / SĐT']);
        const thoiGian = getTruckVal(r, ['Thời gian lấy hàng', 'Thoi gian lay hang', 'Pickup Time', 'Thời gian', 'Time']);
        const note = getTruckVal(r, ['Note', 'Ghi chú', 'Ghi chu', 'Notes']);
        let savedShipIds = {{}};
        try {{ savedShipIds = JSON.parse(localStorage.getItem('truck_shipping_ids_map') || '{{}}'); }} catch(e) {{}}
        const tripRowId = r._rowid || r.id || (idx + 1);
        const tripKey = `${{getTruckVal(r, ['Ngày MTO', 'Dự kiến giao hàng', 'Ngày'])}}_${{chuyenNum}}_${{soXe}}_${{tripRowId}}`;
        const shipIdVal = getTruckVal(r, ['Shipping ID', 'Shiping ID', 'ShippingId', 'Ship Order']) || (savedShipIds[tripKey] || '');

        let statusBadge = '';
        const mXd = r._matchedXeDen || {{}};
        const leadBadge = (r._leadtime === 'Late')
          ? `<span style="background: rgba(239, 68, 68, 0.2); color: #f87171; border: 1px solid rgba(248, 113, 113, 0.4); padding: 3px 8px; border-radius: 4px; font-weight: 800; font-size: 11px;">⚠️ Late</span>`
          : (r._leadtime === 'On Time'
              ? `<span style="background: rgba(34, 197, 94, 0.2); color: #4ade80; border: 1px solid rgba(74, 222, 128, 0.4); padding: 3px 8px; border-radius: 4px; font-weight: 800; font-size: 11px;">⚡ On Time</span>`
              : `<span style="background: rgba(148, 163, 184, 0.15); color: #94a3b8; border: 1px solid rgba(148, 163, 184, 0.3); padding: 3px 8px; border-radius: 4px; font-weight: 700; font-size: 11px;">⏳ Pending</span>`);

        const curMode = getTruckVal(mXd, ['OUT Pallet', 'OUT Type']) ? 'OUT' : 'IN';
        const curType = getTruckVal(mXd, [curMode === 'IN' ? 'IN Type' : 'OUT Type']) || 'MTI';
        const curPallet = getTruckVal(mXd, [curMode === 'IN' ? 'IN Pallet' : 'OUT Pallet', 'Pallet', 'pallet']) || '';
        const curWh = getTruckVal(mXd, [curMode === 'IN' ? 'IN From WH' : 'OUT To WH', 'Kho xuất phát', 'From WH']) || khoDi || '';
        const curOp = getTruckVal(mXd, [curMode === 'IN' ? 'IN Operater' : 'OUT Operater', 'IN Operator', 'OUT Operator']) || '';
        const curShipId = getTruckVal(mXd, ['Shipping ID', 'Shiping ID', 'ShippingId']) || shipIdVal || '';
        const rawDriverName = (getTruckVal(mXd, ['Name', 'tai_xe', 'Tài xế']) || taiXe || '').replace(/[\x27\x22]/g, ' ').replace(/\s+/g, ' ').trim();
        const rawDriverPhone = (getTruckVal(mXd, ['Số ĐT', 'Số điện thoại', 'sdt']) || '').replace(/[\x27\x22]/g, ' ').replace(/\s+/g, ' ').trim();

        if (r._isArrived && r._matchedXeDen) {{
          const plt = getTruckVal(mXd, ['IN Pallet', 'Pallet', 'pallet']) || curPallet || 0;
          const timeStr = getTruckVal(mXd, ['Dấu thời gian', 'Thời gian', 'gio_den', 'created_at']) || 'Đã vào kho';
          statusBadge = `
            <div style="display: inline-flex; align-items: center; gap: 8px;">
              <div style="display: inline-flex; align-items: center; gap: 6px; background: rgba(74, 222, 128, 0.15); border: 1px solid #16a34a; border-radius: 6px; padding: 4px 10px; color: #4ade80; font-weight: 800; font-size: 11px;">
                <span>✓ ĐÃ ĐẾN</span> &bull; <span>📦 ${{plt}} Pallet</span> &bull; <span>⏰ ${{timeStr}}</span>
              </div>
              ${{leadBadge}}
            </div>
          `;
        }} else {{
          statusBadge = `
            <div style="display: inline-flex; align-items: center; gap: 8px;">
              <div style="display: inline-flex; align-items: center; gap: 6px; background: rgba(251, 113, 133, 0.15); border: 1px solid #e11d48; border-radius: 6px; padding: 4px 10px; color: #fb7185; font-weight: 800; font-size: 11px;">
                <span>⏳ CHƯA ĐẾN</span> &bull; <span style="font-weight: 500; opacity: 0.85;">(Chưa có thông tin xe đến khớp)</span>
              </div>
              ${{leadBadge}}
            </div>
          `;
        }}

        const _inData = {{ type: getTruckVal(mXd, ['IN Type']) || '', pallet: getTruckVal(mXd, ['IN Pallet']) || '', wh: getTruckVal(mXd, ['IN From WH']) || '', op: getTruckVal(mXd, ['IN Operater','IN Operator']) || '' }};
        const _outData = {{ type: getTruckVal(mXd, ['OUT Type']) || '', pallet: getTruckVal(mXd, ['OUT Pallet']) || '', wh: getTruckVal(mXd, ['OUT To WH']) || '', op: getTruckVal(mXd, ['OUT Operater','OUT Operator']) || '' }};
        const formHtml = getCardXeDenFormHtml(`card_plan_${{idx}}`, curMode, curType, curPallet, curWh, curOp, curShipId, soXe, rawDriverName, rawDriverPhone, khoDi, _inData, _outData, r._leadtime || 'On Time');

        const displaySoXe = String(soXe || '').replace(/[\x27\x22]/g, ' ').replace(/\s+/g, ' ').trim();
        const displayTaiXe = String(taiXe || '').replace(/[\x27\x22]/g, ' ').trim();

        html += `
          <div style="background: rgba(255,255,255,0.03); border: 1px solid rgba(255,255,255,0.1); border-radius: 10px; padding: 14px 16px;">
            <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; flex-wrap: wrap; gap: 6px;">
              <div style="font-size: 14px; font-weight: 800; color: #f1f5f9; display: flex; align-items: center; gap: 6px;">
                <span>🚚</span> CHUYẾN ${{chuyenNum}}
              </div>
              ${{statusBadge}}
            </div>
            <div style="display: flex; gap: 16px; justify-content: space-between; align-items: flex-start; flex-wrap: wrap;">
              <div style="display: flex; flex-direction: column; gap: 4px; font-size: 13px; flex: 1; min-width: 250px;">
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">🔴 Kho đi</span>: <strong style="color: #f8fafc;">${{getKhoDisplayName(khoDi)}}</strong></div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">🟠 Pick thêm</span>: <span style="color: #cbd5e1;">${{pickThem}}</span></div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">🟡 Kho đến</span>: <strong style="color: #f8fafc;">${{khoDen}}</strong></div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">⚪ Drop thêm</span>: <span style="color: #cbd5e1;">${{dropThem}}</span></div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">🔷 Số xe</span>: <strong style="color: #38bdf8;">${{displaySoXe}}</strong></div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">🚐 Loại xe</span>: <span style="color: #cbd5e1;">${{loaiXe}}</span></div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">🏷️ Vendor</span>: <strong style="color: #f59e0b;">${{vendor}}</strong></div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">👨‍✈️ Tài xế</span>: <strong style="color: #4ade80; white-space: pre-wrap;">${{displayTaiXe}}</strong></div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">⏱️ Leadtime</span>: ${{leadBadge}}</div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">⏰ Thời gian lấy</span>: <span style="color: #cbd5e1;">${{thoiGian}}</span></div>
                <div><span style="display: inline-block; width: 140px; color: #94a3b8;">📝 Note</span>: <span style="color: #cbd5e1;">${{note}}</span></div>
              </div>
              ${{formHtml}}
            </div>
          </div>
        `;
      }});
      listEl.innerHTML = html;
    }}

    function getCardXeDenFormHtml(cardId, curMode, curType, curPallet, curWh, curOp, curShipId, plate, driverName, driverPhone, fromWh, inData, outData, leadtime) {{
      inData = inData || {{}};
      outData = outData || {{}};
      const isOut = (curMode === 'OUT');
      const typeOptions = ['MTI', 'MTO', 'RTS', 'Liquidate', 'Transfer', 'IB'].map(t => 
        `<option value="${{t}}" ${{curType === t ? 'selected' : ''}}>${{t}}</option>`
      ).join('');

      const opOptions = [
        'LÊ VĂN TÚ', 'BÙI QUỐC PHÁP', 'NGUYỄN THẾ CÔNG',
        'LÊ VIỆT HÙNG', 'BÙI VĂN TUẤN', 'Huỳnh Vũ Tuyết Nhi'
      ].map(op => 
        `<option value="${{op}}" ${{curOp === op ? 'selected' : ''}}>${{op}}</option>`
      ).join('');

      const cleanPlate = (plate || '').replace(/[\x27\x22]/g, ' ').replace(/\s+/g, ' ').trim();
      const cleanDriver = (driverName || '').replace(/[\x27\x22]/g, ' ').replace(/\s+/g, ' ').trim();
      const cleanPhone = (driverPhone || '').replace(/[\x27\x22]/g, ' ').replace(/\s+/g, ' ').trim();
      const cleanFromWh = (fromWh || '').replace(/['"]/g, ' ').replace(/\s+/g, ' ').trim();

      let inOutSummary = '';
      const hasIn = Boolean(inData.type || inData.pallet);
      const hasOut = Boolean(outData.type || outData.pallet);
      if (hasIn && hasOut) {{
        inOutSummary = ` <span style="font-size: 10px; color: #4ade80; background: rgba(34, 197, 94, 0.2); padding: 2px 6px; border-radius: 4px; border: 1px solid rgba(74, 222, 128, 0.4);">✓ IN + OUT</span>`;
      }} else if (hasIn) {{
        inOutSummary = ` <span style="font-size: 10px; color: #4ade80; background: rgba(34, 197, 94, 0.15); padding: 2px 6px; border-radius: 4px;">✓ Đã có IN</span>`;
      }} else if (hasOut) {{
        inOutSummary = ` <span style="font-size: 10px; color: #fb923c; background: rgba(234, 88, 12, 0.15); padding: 2px 6px; border-radius: 4px;">✓ Đã có OUT</span>`;
      }}

      return `
        <div style="background: rgba(245, 158, 11, 0.08); border: 1.5px solid #f59e0b; box-shadow: 0 4px 15px rgba(245, 158, 11, 0.12); border-radius: 8px; padding: 10px 12px; display: flex; flex-direction: column; gap: 8px; min-width: 290px; flex: 1;">
          <input type="hidden" id="card_plate_${{cardId}}" value="${{cleanPlate}}">
          <input type="hidden" id="card_driver_${{cardId}}" value="${{cleanDriver}}">
          <input type="hidden" id="card_phone_${{cardId}}" value="${{cleanPhone}}">
          <input type="hidden" id="card_fromwh_${{cardId}}" value="${{cleanFromWh}}">
          <input type="hidden" id="card_lead_${{cardId}}" value="${{leadtime || 'On Time'}}">
          <input type="hidden" id="card_in_type_${{cardId}}" value="${{inData.type || ''}}">
          <input type="hidden" id="card_in_pallet_${{cardId}}" value="${{inData.pallet || ''}}">
          <input type="hidden" id="card_in_wh_${{cardId}}" value="${{inData.wh || ''}}">
          <input type="hidden" id="card_in_op_${{cardId}}" value="${{inData.op || ''}}">
          <input type="hidden" id="card_out_type_${{cardId}}" value="${{outData.type || ''}}">
          <input type="hidden" id="card_out_pallet_${{cardId}}" value="${{outData.pallet || ''}}">
          <input type="hidden" id="card_out_wh_${{cardId}}" value="${{outData.wh || ''}}">
          <input type="hidden" id="card_out_op_${{cardId}}" value="${{outData.op || ''}}">
          <div style="display: flex; justify-content: space-between; align-items: center; border-bottom: 1px dashed rgba(255,255,255,0.1); padding-bottom: 5px;">
            <span style="font-size: 11px; font-weight: 800; color: #f59e0b; text-transform: uppercase;">📝 CẬP NHẬT XE ĐẾN</span>${{inOutSummary}}
            <div style="display: inline-flex; gap: 4px; background: rgba(0,0,0,0.3); padding: 2px; border-radius: 6px; border: 1px solid rgba(255,255,255,0.1);">
              <button type="button" id="card_btn_in_${{cardId}}" onclick="setCardInOutMode('${{cardId}}', 'IN')" style="font-size: 10px; font-weight: 800; padding: 3px 8px; border-radius: 4px; cursor: pointer; ${{!isOut ? 'background: #16a34a; color: #fff; border: 1px solid #4ade80;' : 'background: rgba(255,255,255,0.05); color: #94a3b8; border: 1px solid rgba(255,255,255,0.1);'}}">IN</button>
              <button type="button" id="card_btn_out_${{cardId}}" onclick="setCardInOutMode('${{cardId}}', 'OUT')" style="font-size: 10px; font-weight: 800; padding: 3px 8px; border-radius: 4px; cursor: pointer; ${{isOut ? 'background: #ea580c; color: #fff; border: 1px solid #fb923c;' : 'background: rgba(255,255,255,0.05); color: #94a3b8; border: 1px solid rgba(255,255,255,0.1);'}}">OUT</button>
              <input type="hidden" id="card_mode_${{cardId}}" value="${{isOut ? 'OUT' : 'IN'}}">
            </div>
          </div>
          <div style="display: flex; gap: 6px;">
            <div style="flex: 1;">
              <label style="font-size: 10px; color: #94a3b8; display: block; margin-bottom: 2px;">Type:</label>
              <select id="card_type_${{cardId}}" style="width: 100%; background: #0f172a; border: 1px solid #38bdf8; border-radius: 5px; color: #fff; padding: 4px 6px; font-size: 11px; outline: none;">
                <option value="">-- Chọn Type --</option>
                ${{typeOptions}}
              </select>
            </div>
            <div style="width: 90px;">
              <label style="font-size: 10px; color: #94a3b8; display: block; margin-bottom: 2px;">Số Pallet:</label>
              <input type="number" id="card_pallet_${{cardId}}" value="${{curPallet || ''}}" placeholder="Pallet..." style="width: 100%; background: #0f172a; border: 1px solid #4ade80; border-radius: 5px; color: #4ade80; padding: 4px 6px; font-size: 11px; font-weight: 700; outline: none;">
            </div>
          </div>
          <div style="display: flex; gap: 6px;">
            <div style="flex: 1;">
              <label style="font-size: 10px; color: #94a3b8; display: block; margin-bottom: 2px;">Kho (Warehouse):</label>
              <input type="text" list="card_wh_list_${{cardId}}" id="card_wh_${{cardId}}" value="${{curWh || ''}}" placeholder="${{isOut ? 'To WH...' : 'From WH...'}}" style="width: 100%; background: #0f172a; border: 1px solid #38bdf8; border-radius: 5px; color: #fff; padding: 4px 6px; font-size: 11px; outline: none;">
              <datalist id="card_wh_list_${{cardId}}">
                 <option value="VNNL">Bắc Ninh (VNNL)</option>
                 <option value="VNA">Yên Phong (VNA)</option>
                 <option value="VNG">Yên Phong (VNG)</option>
                 <option value="VNVL">VNVL</option>
                 <option value="VNSL">VNSL</option>
                 <option value="VNS">VNS</option>
                 <option value="Củ Chi mới">Củ Chi mới</option>
                 <option value="Xuyên Á">Xuyên Á</option>
                 <option value="Thuận Thành">Thuận Thành</option>
               </datalist>
            </div>
            <div style="flex: 1;">
              <label style="font-size: 10px; color: #94a3b8; display: block; margin-bottom: 2px;">Operater:</label>
              <select id="card_op_${{cardId}}" style="width: 100%; background: #0f172a; border: 1px solid #38bdf8; border-radius: 5px; color: #fff; padding: 4px 6px; font-size: 11px; outline: none;">
                <option value="">-- Chọn Operater --</option>
                ${{opOptions}}
              </select>
            </div>
          </div>
          <div style="display: flex; gap: 6px; align-items: flex-end;">
            <div style="flex: 1;">
              <label style="font-size: 10px; color: #94a3b8; display: block; margin-bottom: 2px;">Shipping ID:</label>
              <input type="text" id="card_ship_${{cardId}}" value="${{curShipId || ''}}" placeholder="Nhập Shipping ID..." style="width: 100%; background: #0f172a; border: 1px solid #38bdf8; border-radius: 5px; color: #fff; padding: 4px 6px; font-size: 11px; outline: none;">
            </div>
            <button type="button" onclick="saveXeDenFromCard('${{cardId}}')" style="background: linear-gradient(135deg, #0284c7, #2563eb); color: #fff; border: none; border-radius: 5px; padding: 5px 14px; font-size: 11px; font-weight: 800; cursor: pointer; white-space: nowrap; height: 26px;">
              💾 Lưu
            </button>
          </div>
        </div>
      `;
    }}




    // Map mã kho → tên đầy đủ
    function getKhoDisplayName(code) {{
      const khoMap = {{
        'VNNL': 'Bắc Ninh (VNNL)',
        'VNA':  'Yên Phong (VNA)',
        'VNG':  'Yên Phong (VNG)',
        'VNVL': 'VNVL',
        'VNSL': 'VNSL',
        'VNS':  'VNS',
      }};
      if (!code) return '';
      const upper = code.trim().toUpperCase();
      if (khoMap[upper]) return khoMap[upper];
      // Nếu tên kho chứa từ khoá → gắn mã vào
      const nameMap = [
        [['BẮC NINH','BACNINH','VNNL'], 'Bắc Ninh (VNNL)'],
        [['YÊN PHONG','YENPHONG','VNA'], 'Yên Phong (VNA)'],
        [['VNG'], 'Yên Phong (VNG)'],
      ];
      const up2 = code.trim().toUpperCase().replace(/[\s-]/g,'');
      for (const [keys, label] of nameMap) {{
        if (keys.some(k => up2.includes(k.replace(/[\s-]/g,'')))) return label;
      }}
      return code;
    }}

    function setCardInOutMode(cardId, mode) {{
      const modeInp = document.getElementById('card_mode_' + cardId);
      const oldMode = modeInp ? modeInp.value : 'IN';
      if (modeInp) modeInp.value = mode;

      // Lưu giá trị hiện tại vào hidden fields của mode cũ trước khi swap
      const typeEl = document.getElementById('card_type_' + cardId);
      const palletEl = document.getElementById('card_pallet_' + cardId);
      const whEl = document.getElementById('card_wh_' + cardId);
      const opEl = document.getElementById('card_op_' + cardId);

      if (oldMode !== mode) {{
        // Save current form values to old mode's hidden fields
        const oldPrefix = 'card_' + oldMode.toLowerCase() + '_';
        const saveType = document.getElementById(oldPrefix + 'type_' + cardId);
        const savePallet = document.getElementById(oldPrefix + 'pallet_' + cardId);
        const saveWh = document.getElementById(oldPrefix + 'wh_' + cardId);
        const saveOp = document.getElementById(oldPrefix + 'op_' + cardId);
        if (saveType && typeEl) saveType.value = typeEl.value;
        if (savePallet && palletEl) savePallet.value = palletEl.value;
        if (saveWh && whEl) saveWh.value = whEl.value;
        if (saveOp && opEl) saveOp.value = opEl.value;

        // Load new mode's values from hidden fields
        const newPrefix = 'card_' + mode.toLowerCase() + '_';
        const loadType = document.getElementById(newPrefix + 'type_' + cardId);
        const loadPallet = document.getElementById(newPrefix + 'pallet_' + cardId);
        const loadWh = document.getElementById(newPrefix + 'wh_' + cardId);
        const loadOp = document.getElementById(newPrefix + 'op_' + cardId);
        if (typeEl && loadType) typeEl.value = loadType.value || '';
        if (palletEl && loadPallet) palletEl.value = loadPallet.value || '';
        if (whEl && loadWh) whEl.value = loadWh.value || '';
        if (opEl && loadOp) opEl.value = loadOp.value || '';
      }}

      const btnIn = document.getElementById('card_btn_in_' + cardId);
      const btnOut = document.getElementById('card_btn_out_' + cardId);
      if (mode === 'IN') {{
        if (btnIn) {{ btnIn.style.background = '#16a34a'; btnIn.style.color = '#fff'; btnIn.style.border = '1px solid #4ade80'; }}
        if (btnOut) {{ btnOut.style.background = 'rgba(255,255,255,0.05)'; btnOut.style.color = '#94a3b8'; btnOut.style.border = '1px solid rgba(255,255,255,0.2)'; }}
        if (whEl) whEl.placeholder = 'From WH...';
      }} else {{
        if (btnOut) {{ btnOut.style.background = '#ea580c'; btnOut.style.color = '#fff'; btnOut.style.border = '1px solid #fb923c'; }}
        if (btnIn) {{ btnIn.style.background = 'rgba(255,255,255,0.05)'; btnIn.style.color = '#94a3b8'; btnIn.style.border = '1px solid rgba(255,255,255,0.2)'; }}
        if (whEl) whEl.placeholder = 'To WH...';
      }}
    }}

    async function saveXeDenFromCard(cardId, showAlert = true) {{
      const plate = (document.getElementById('card_plate_' + cardId)?.value || '').trim();
      const driverName = (document.getElementById('card_driver_' + cardId)?.value || '').trim();
      const driverPhone = (document.getElementById('card_phone_' + cardId)?.value || '').trim();
      const fromWh = (document.getElementById('card_fromwh_' + cardId)?.value || '').trim();
      const cardLeadtime = (document.getElementById('card_lead_' + cardId)?.value || 'On Time').trim();

      const curMode = (document.getElementById('card_mode_' + cardId)?.value || 'IN').trim();
      const curType = (document.getElementById('card_type_' + cardId)?.value || '').trim();
      const curPallet = (document.getElementById('card_pallet_' + cardId)?.value || '0').trim();
      const curWh = (document.getElementById('card_wh_' + cardId)?.value || '').trim();
      const curOp = (document.getElementById('card_op_' + cardId)?.value || '').trim();
      const shippingId = (document.getElementById('card_ship_' + cardId)?.value || '').trim();

      // Cập nhật giá trị visible hiện tại vào hidden field của mode đang active
      const curPrefix = 'card_' + curMode.toLowerCase() + '_';
      const curTypeEl = document.getElementById(curPrefix + 'type_' + cardId);
      const curPalletEl = document.getElementById(curPrefix + 'pallet_' + cardId);
      const curWhEl = document.getElementById(curPrefix + 'wh_' + cardId);
      const curOpEl = document.getElementById(curPrefix + 'op_' + cardId);
      if (curTypeEl) curTypeEl.value = curType;
      if (curPalletEl) curPalletEl.value = curPallet;
      if (curWhEl) curWhEl.value = curWh;
      if (curOpEl) curOpEl.value = curOp;

      // Lấy đầy đủ cả 2 chiều IN và OUT
      const inType = (document.getElementById('card_in_type_' + cardId)?.value || (curMode === 'IN' ? curType : '')).trim();
      const inPallet = (document.getElementById('card_in_pallet_' + cardId)?.value || (curMode === 'IN' ? curPallet : '')).trim();
      const inWh = (document.getElementById('card_in_wh_' + cardId)?.value || (curMode === 'IN' ? curWh : '')).trim();
      const inOp = (document.getElementById('card_in_op_' + cardId)?.value || (curMode === 'IN' ? curOp : '')).trim();

      const outType = (document.getElementById('card_out_type_' + cardId)?.value || (curMode === 'OUT' ? curType : '')).trim();
      const outPallet = (document.getElementById('card_out_pallet_' + cardId)?.value || (curMode === 'OUT' ? curPallet : '')).trim();
      const outWh = (document.getElementById('card_out_wh_' + cardId)?.value || (curMode === 'OUT' ? curWh : '')).trim();
      const outOp = (document.getElementById('card_out_op_' + cardId)?.value || (curMode === 'OUT' ? curOp : '')).trim();

      let targetDate = '';
      if (selectedDate && selectedDate !== 'ALL') {{
        targetDate = selectedDate;
      }} else {{
        targetDate = getTodayInfo().ddmmyyyy;
      }}

      const params = new URLSearchParams({{
        so_xe: plate || '',
        sdt: driverPhone || '',
        tai_xe: driverName || '',
        pallet: curPallet || '0',
        kho_di: fromWh || '',
        inout_mode: curMode,
        type_val: curType,
        operater: curOp,
        warehouse: curWh,
        in_type: inType,
        in_pallet: inPallet,
        in_wh: inWh,
        in_op: inOp,
        out_type: outType,
        out_pallet: outPallet,
        out_wh: outWh,
        out_op: outOp,
        shipping_id: shippingId,
        leadtime: cardLeadtime,
        ngay: targetDate
      }});

      let saved = false;
      for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
        try {{
          const resp = await fetch(`${{base}}/api/add_xe_den_vndl?${{params.toString()}}`);
          if (resp.ok) {{
            saved = true;
            break;
          }}
        }} catch (e) {{}}
      }}

      // Cập nhật vào bộ nhớ frontend
      if (!currentData.truck_tracker) currentData.truck_tracker = {{}};
      if (!currentData.truck_tracker.xe_den_vndl) currentData.truck_tracker.xe_den_vndl = [];
      const xeList = currentData.truck_tracker.xe_den_vndl;
      
      const cleanPlate = (plate || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
      let found = xeList.find(x => {{
        const xPlate = (x['Biển số xe'] || x['Biển số'] || x['Số xe'] || '').replace(/[^a-zA-Z0-9]/g, '').toLowerCase();
        return cleanPlate && xPlate && (cleanPlate.includes(xPlate) || xPlate.includes(cleanPlate));
      }});

      const totalPlt = (parseInt(inPallet || 0) || 0) + (parseInt(outPallet || 0) || 0) || parseInt(curPallet || 0) || 0;

      if (found) {{
        if (inType || inPallet || inWh || inOp) {{
          found['IN Type'] = inType;
          found['IN Pallet'] = inPallet;
          found['IN From WH'] = inWh;
          found['IN Operater'] = inOp;
        }}
        if (outType || outPallet || outWh || outOp) {{
          found['OUT Type'] = outType;
          found['OUT Pallet'] = outPallet;
          found['OUT To WH'] = outWh;
          found['OUT Operater'] = outOp;
        }}
        found['Shipping ID'] = shippingId;
        found['Pallet thực tế'] = totalPlt;
        found['IN Pallet'] = inPallet || totalPlt;
        found['Pallet'] = totalPlt;
        found['Leadtime'] = cardLeadtime;
      }} else {{
        const nowTs = targetDate + ' ' + new Date().toLocaleTimeString('vi-VN');
        xeList.push({{
          'Dấu thời gian': nowTs,
          'Biển số xe': plate,
          'Name': driverName,
          'Số ĐT': driverPhone,
          'Kho xuất phát': fromWh,
          'IN Type': inType,
          'IN Pallet': inPallet || totalPlt,
          'IN From WH': inWh,
          'IN Operater': inOp,
          'OUT Type': outType,
          'OUT Pallet': outPallet,
          'OUT To WH': outWh,
          'OUT Operater': outOp,
          'Shipping ID': shippingId,
          'Pallet thực tế': totalPlt,
          'Pallet': totalPlt,
          'Leadtime': cardLeadtime,
          ngay: targetDate
        }});
      }}

      // Lưu shipping ID vào cache localStorage nếu có
      if (shippingId) {{
        try {{
          const sMap = JSON.parse(localStorage.getItem('truck_shipping_ids_map') || '{{}}');
          sMap[plate] = shippingId;
          localStorage.setItem('truck_shipping_ids_map', JSON.stringify(sMap));
        }} catch (e) {{}}
      }}

      // Tự động chuyển Arrived sang 'Yes (Đã đến)' cho đơn hàng có Shipping ID / Ship Order / ASN khớp
      let matchedOrders = [];
      if (shippingId) {{
        const tokens = String(shippingId || '').replace(/,/g, ' ').replace(/;/g, ' ').split(' ').map(t => t.trim()).filter(t => t.length >= 4);
        const allBuckets = [];
        if (currentData.mti && currentData.mti.items) allBuckets.push({{ name: 'MTI', items: currentData.mti.items }});
        if (currentData.inbound && currentData.inbound.items) allBuckets.push({{ name: 'Inbound', items: currentData.inbound.items }});
        
        const todayStr = getTodayInfo().ddmmyyyy;

        tokens.forEach(tok => {{
          const tokClean = tok.toLowerCase();
          allBuckets.forEach(b => {{
            const prefix = (b.name === 'Inbound') ? 'ib_' : 'mti_';
            let savedOverrides = {{}};
            let savedDates = {{}};
            try {{
              savedOverrides = JSON.parse(localStorage.getItem(prefix + 'arrived_overrides') || '{{}}');
              savedDates = JSON.parse(localStorage.getItem(prefix + 'arrived_dates') || '{{}}');
            }} catch (e) {{}}

            b.items.forEach(r => {{
              const shipOrder = (r['Ship Order'] || '').trim();
              const asn = (r['ASN ID'] || r['ASN'] || '').trim();
              const soClean = shipOrder.toLowerCase();
              const asnClean = asn.toLowerCase();

              const isMatch = (soClean && (soClean === tokClean || soClean.includes(tokClean) || tokClean.includes(soClean))) ||
                              (asnClean && (asnClean === tokClean || asnClean.includes(tokClean) || tokClean.includes(asnClean)));

              if (isMatch) {{
                const targetKey = shipOrder || asn;
                r.Arrived = 'Yes';
                const put = Number(r.Putaway || 0);
                const qty = Number(r.Quantity || 0);
                const backlog = Number(r.Backlog || 0);
                if (put === 0) {{
                  r.Status = 'Arriving';
                  r['Receive Date'] = todayStr;
                }} else if (backlog <= 700 || put >= qty) {{
                  r.Status = 'Done';
                }} else {{
                  r.Status = 'In Progress';
                }}

                savedOverrides[targetKey] = 'Yes';
                savedDates[targetKey] = todayStr;
                matchedOrders.push(targetKey);

                for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
                  try {{
                    fetch(`${{base}}/api/set_arrived?ship_order=${{encodeURIComponent(targetKey)}}&arrived=Yes&bucket=${{b.name}}`);
                  }} catch (e) {{}}
                }}
              }}
            }});

            try {{
              localStorage.setItem(prefix + 'arrived_overrides', JSON.stringify(savedOverrides));
              localStorage.setItem(prefix + 'arrived_dates', JSON.stringify(savedDates));
            }} catch (e) {{}}
          }});
        }});
      }}

      // Cập nhật lại UI bảng chính và modal
      updateUI();
      renderTruckTrackerModal();

      let infoParts = [];
      if (inType || inPallet) infoParts.push(`📥 IN: ${{inType||'-'}} (${{inPallet||0}} plt) WH: ${{inWh||'-'}}`);
      if (outType || outPallet) infoParts.push(`📤 OUT: ${{outType||'-'}} (${{outPallet||0}} plt) WH: ${{outWh||'-'}}`);
      if (infoParts.length === 0) infoParts.push(`[${{curMode}} ${{curPallet}} Pallet]`);

      let msg = `✅ ĐÃ LƯU THÀNH CÔNG!\n\n• Xe: ${{plate || 'Không có'}}\n• ` + infoParts.join('\\n• ') + `\n• Operater: ${{curOp || inOp || outOp || 'Chưa chọn'}}\n• Shipping ID: ${{shippingId || 'Trống'}}`;
      if (matchedOrders.length > 0) {{
        const uniqueOrders = Array.from(new Set(matchedOrders));
        msg += `\n\n🚚 Đã tự động chuyển ARRIVED -> 'Yes (Đã đến)' cho đơn:\n` + uniqueOrders.join(', ');
      }}
      if (showAlert) {{
        alert(msg);
      }} else {{
        showToast(`💾 Đã lưu dữ liệu xe ${{plate}} vào SQL & Google Sheet!`);
      }}
    }}

    function setXeDenMode(mode) {{
      document.getElementById('manual_inout_mode').value = mode;
      const btnIn = document.getElementById('btn_in_toggle');
      const btnOut = document.getElementById('btn_out_toggle');
      const lbl = document.getElementById('label_warehouse');
      if (mode === 'IN') {{
        btnIn.style.background = '#16a34a';
        btnIn.style.border = '2px solid #4ade80';
        btnIn.style.color = '#fff';
        btnOut.style.background = 'rgba(255,255,255,0.05)';
        btnOut.style.border = '1px solid rgba(255,255,255,0.2)';
        btnOut.style.color = '#94a3b8';
        if (lbl) lbl.textContent = '🏭 From WH (IN):';
      }} else {{
        btnOut.style.background = '#b45309';
        btnOut.style.border = '2px solid #f59e0b';
        btnOut.style.color = '#fff';
        btnIn.style.background = 'rgba(255,255,255,0.05)';
        btnIn.style.border = '1px solid rgba(255,255,255,0.2)';
        btnIn.style.color = '#94a3b8';
        if (lbl) lbl.textContent = '🏭 To WH (OUT):';
      }}
    }}

    async function submitManualXeDen() {{
      const soXe = (document.getElementById('manual_so_xe').value || '').trim();
      const sdt = (document.getElementById('manual_sdt').value || '').trim();
      const taiXe = (document.getElementById('manual_tai_xe').value || '').trim();
      const pallet = (document.getElementById('manual_pallet').value || '0').trim();
      const khoDi = (document.getElementById('manual_kho_di').value || '').trim();
      const ghiChu = (document.getElementById('manual_ghi_chu').value || '').trim();
      const inOutMode = (document.getElementById('manual_inout_mode').value || 'IN').trim();
      const typeSel = (document.getElementById('manual_type').value || '').trim();
      const operater = (document.getElementById('manual_operater').value || '').trim();
      const warehouse = (document.getElementById('manual_warehouse').value || '').trim();
      const shippingId = (document.getElementById('manual_shipping_id').value || '').trim();

      if (!soXe && !sdt) {{
        alert('Vui lòng nhập Biển số xe hoặc Số điện thoại tài xế!');
        return;
      }}

      let targetDate = '';
      if (selectedDate && selectedDate !== 'ALL') {{
        targetDate = selectedDate;
      }} else {{
        const today = getTodayInfo();
        targetDate = today.ddmmyyyy;
      }}

      // Lưu vào SQLite bảng [xe đến VNDL]
      let saved = false;
      for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
        try {{
          const params = new URLSearchParams({{
            so_xe: soXe, sdt: sdt, tai_xe: taiXe,
            pallet: pallet, kho_di: khoDi, ghi_chu: ghiChu,
            ngay: targetDate, inout_mode: inOutMode,
            type_val: typeSel, operater: operater,
            warehouse: warehouse, shipping_id: shippingId
          }});
          const resp = await fetch(`${{base}}/api/add_xe_den_vndl?${{params.toString()}}`);
          if (resp.ok) {{
            saved = true;
            break;
          }}
        }} catch (e) {{}}
      }}

      // Cập nhật ngay vào bộ nhớ frontend
      if (!currentData.truck_tracker) currentData.truck_tracker = {{}};
      if (!currentData.truck_tracker.xe_den_vndl) currentData.truck_tracker.xe_den_vndl = [];
      const nowTs = targetDate + ' ' + new Date().toLocaleTimeString('vi-VN');
      currentData.truck_tracker.xe_den_vndl.push({{
        'Dấu thời gian': nowTs,
        'Biển số xe': soXe,
        'Name': taiXe,
        'Số ĐT': sdt,
        'Kho xuất phát': khoDi,
        'IN Type': inOutMode === 'IN' ? typeSel : '',
        'IN Pallet': inOutMode === 'IN' ? pallet : '',
        'IN From WH': inOutMode === 'IN' ? warehouse : '',
        'IN Operater': inOutMode === 'IN' ? operater : '',
        'OUT Type': inOutMode === 'OUT' ? typeSel : '',
        'OUT Pallet': inOutMode === 'OUT' ? pallet : '',
        'OUT To WH': inOutMode === 'OUT' ? warehouse : '',
        'OUT Operater': inOutMode === 'OUT' ? operater : '',
        'Shipping ID': shippingId,
        ngay: targetDate,
        gio_den: new Date().toLocaleTimeString('vi-VN')
      }});

      showToast(`Đã lưu xe ${{soXe || sdt}} [${{inOutMode}}] ${{pallet}} pallet vào [xe đến VNDL]!`);
      // Reset form
      document.getElementById('manual_so_xe').value = '';
      document.getElementById('manual_sdt').value = '';
      document.getElementById('manual_tai_xe').value = '';
      document.getElementById('manual_pallet').value = '';
      document.getElementById('manual_kho_di').value = '';
      document.getElementById('manual_ghi_chu').value = '';
      document.getElementById('manual_type').value = '';
      document.getElementById('manual_operater').value = '';
      document.getElementById('manual_warehouse').value = '';
      document.getElementById('manual_shipping_id').value = '';

      updateUI();
      filterTruckModal('all');
    }}

    async function saveTruckShippingId(tripIdx, rowId) {{
      const inputEl = document.getElementById(`truckShipIdInput_${{tripIdx}}`);
      if (!inputEl) return;
      const newShipId = inputEl.value.trim();
      const truckData = currentData.truck_tracker || {{ items: [] }};
      const allItems = truckData.items || [];
      const trip = allItems[tripIdx] || {{}};

      const tripDate = getTruckVal(trip, ['Ngày MTO', 'Dự kiến giao hàng', 'Ngày', 'Ngay', 'Date']) || '';
      const tripChuyen = getTruckVal(trip, ['Số chuyến', 'So chuyen', 'Chuyến', 'Chuyen']) || (tripIdx + 1);
      const tripSoXe = getTruckVal(trip, ['Số xe', 'So xe', 'Biển số', 'Bien so', 'License Plate']);
      const tripRowId = rowId || trip._rowid || trip.id || '';

      trip['Shipping ID'] = newShipId;
      trip['Shiping ID'] = newShipId;

      const tripKey = `${{tripDate}}_${{tripChuyen}}_${{tripSoXe}}`;
      try {{
        const savedMap = JSON.parse(localStorage.getItem('truck_shipping_ids_map') || '{{}}');
        savedMap[tripKey] = newShipId;
        savedMap[`${{tripKey}}_${{tripRowId}}`] = newShipId;
        localStorage.setItem('truck_shipping_ids_map', JSON.stringify(savedMap));
      }} catch (e) {{}}

      let foundMatchingMTI = false;
      if (newShipId) {{
        const mtiItems = (currentData.mti && currentData.mti.items) ? currentData.mti.items : (currentData.items || []);
        for (const r of mtiItems) {{
          const so = String(r['Ship Order'] || '').trim().toLowerCase();
          const asn = String(r['ASN'] || '').trim().toLowerCase();
          const target = newShipId.toLowerCase();
          if (so === target || asn === target || so.includes(target) || (target.length >= 6 && target.includes(so))) {{
            const toggleKey = r['Ship Order'] || r['ASN'] || '';
            if (toggleKey) {{
              await toggleArrived(toggleKey, 'Yes');
              foundMatchingMTI = true;
            }}
          }}
        }}
      }}

      let savedDb = false;
      for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
        try {{
          const url = `${{base}}/api/set_truck_shipping_id?rowid=${{encodeURIComponent(tripRowId)}}&date=${{encodeURIComponent(tripDate)}}&chuyen=${{encodeURIComponent(tripChuyen)}}&so_xe=${{encodeURIComponent(tripSoXe)}}&shipping_id=${{encodeURIComponent(newShipId)}}`;
          const resp = await fetch(url);
          if (resp.ok) {{
            savedDb = true;
            break;
          }}
        }} catch (e) {{}}
      }}

      if (foundMatchingMTI) {{
        showToast(`✅ Đã lưu Shipping ID: ${{newShipId}} vào SQL & chuyển đơn MTI sang 'Yes (Đã đến)'!`);
      }} else if (savedDb) {{
        showToast(`✅ Đã lưu Shipping ID: ${{newShipId || '(trống)'}} thành công vào SQL!`);
      }} else {{
        showToast(`💾 Đã lưu Shipping ID: ${{newShipId || '(trống)'}} vào bộ nhớ trình duyệt!`);
      }}
      updateUI();
    }}

    function copyTruckTrackerText() {{
      const dateStr = document.getElementById('truckDateDisplay')?.textContent || '';
      const trips = (window._currentTruckTrips && window._currentTruckTrips.length > 0)
        ? window._currentTruckTrips
        : ((currentData.truck_tracker && currentData.truck_tracker.items) ? currentData.truck_tracker.items : []);
      
      let savedShipIds = {{}};
      try {{ savedShipIds = JSON.parse(localStorage.getItem('truck_shipping_ids_map') || '{{}}'); }} catch(e) {{}}

      let text = `🟦🟦🟦 KHO VNDL 🟦🟦🟦\\n🚚 DAILY TRUCK TRACKER\\n___________________________\\n📌 NGÀY: ${{dateStr}}\\n📦 TỔNG CHUYẾN: ${{trips.length}}\\n___________________________\\n`;

      trips.forEach((r, idx) => {{
        const chuyenVal = getTruckVal(r, ['Số chuyến', 'So chuyen', 'Chuyến', 'Chuyen']);
        const chuyenNum = chuyenVal || (idx + 1);
        const soXe = getTruckVal(r, ['Số xe', 'So xe', 'Biển số', 'Bien so', 'License Plate']);
        const tripRowId = r._rowid || r.id || (idx + 1);
        const tripKey = `${{getTruckVal(r, ['Ngày MTO', 'Dự kiến giao hàng', 'Ngày'])}}_${{chuyenNum}}_${{soXe}}_${{tripRowId}}`;
        const shipId = getTruckVal(r, ['Shipping ID', 'Shiping ID', 'ShippingId', 'Ship Order']) || (savedShipIds[tripKey] || '');

        text += `\\n🚚 CHUYẾN ${{chuyenNum}}\\n`;
        text += `🔴 Kho đi            : ${{getTruckVal(r, ['Kho đi', 'Kho di', 'From', 'Kho xuất'])}}\\n`;
        text += `🟠 Pick thêm        : ${{getTruckVal(r, ['Pick thêm', 'Pick them', 'Pick'])}}\\n`;
        text += `🟡 Kho đến          : ${{getTruckVal(r, ['Kho đến', 'Kho den', 'To', 'Kho nhập'])}}\\n`;
        text += `⚪ Drop thêm        : ${{getTruckVal(r, ['Drop thêm', 'Drop them', 'Drop'])}}\\n`;
        text += `🔷 Số xe             : ${{soXe}}\\n`;
        text += `🚐 Loại xe           : ${{getTruckVal(r, ['Loại xe', 'Loai xe', 'Truck Type', 'Type'])}}\\n`;
        text += `🏷️ Vendor            : ${{getTruckVal(r, ['Vendor', 'Nhà xe', 'Nha xe', 'Đơn vị vận chuyển'])}}\\n`;
        text += `👨‍✈️ Tài xế            : ${{getTruckVal(r, ['Thông tin tài xế', 'Tài xế', 'Tai xe', 'Driver', 'Tài xế / SĐT'])}}\\n`;
        text += `⏰ Thời gian lấy hàng : ${{getTruckVal(r, ['Thời gian lấy hàng', 'Thoi gian lay hang', 'Pickup Time', 'Thời gian', 'Time'])}}\\n`;
        text += `📝 Note              : ${{getTruckVal(r, ['Note', 'Ghi chú', 'Ghi chu', 'Notes'])}}\\n`;
        if (shipId) {{
          text += `📦 Shipping ID       : ${{shipId}}\\n`;
        }}
      }});

      if (navigator.clipboard && navigator.clipboard.writeText) {{
        navigator.clipboard.writeText(text).then(() => {{
          alert('Đã sao chép nội dung Daily Truck Tracker vào clipboard!');
        }}).catch(() => {{
          prompt('Nhấn Ctrl+C để sao chép:', text);
        }});
      }} else {{
        prompt('Nhấn Ctrl+C để sao chép:', text);
      }}
    }}

    function handlePipelineStatClick(idx) {{
      let targetFilter = 'ALL';
      if (currentBucket === 'RTS') {{
        if (idx === 1) targetFilter = 'ALL';
        else if (idx === 2) targetFilter = 'Allocated';
        else if (idx === 3) targetFilter = 'Pending Pick';
        else if (idx === 4) targetFilter = 'Outbound';
        else targetFilter = 'ALL';
      }} else if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
        if (idx === 1) targetFilter = 'ALL';
        else if (idx === 2) targetFilter = 'RACK';
        else if (idx === 3) targetFilter = 'PALLET';
        else targetFilter = 'ALL';
      }} else {{
        // Inbound & MTI
        if (idx === 1) targetFilter = 'ALL';
        else if (idx === 2) targetFilter = 'Pending';
        else if (idx === 3) targetFilter = 'Arriving';
        else if (idx === 4) targetFilter = 'In Progress';
        else if (idx === 5) targetFilter = 'Done';
      }}
      setFilter(targetFilter);

      // Visual highlight for the clicked stat item
      const statItems = document.querySelectorAll('.pipeline-stats .stat-item');
      statItems.forEach((el, i) => {{
        if (i + 1 === idx) {{
          el.style.borderColor = 'rgba(56, 189, 248, 0.7)';
          el.style.background = 'rgba(56, 189, 248, 0.12)';
        }} else {{
          el.style.borderColor = '';
          el.style.background = '';
        }}
      }});
    }}

    async function fetchLiveSQLData(force = false) {{
      for (const base of ['http://localhost:8501', 'http://localhost:8500', '']) {{
        try {{
          const resp = await fetch(`${{base}}/api/data`, {{ cache: 'no-store' }});
          if (resp.ok) {{
            const freshData = await resp.json();
            const newHash = JSON.stringify(freshData);
            if (force || newHash !== lastDataHash) {{
              lastDataHash = newHash;
              currentData = freshData;
              populateDateMonthSelectors();
              updateUI();
              if (force) showToast('⚡ Đã làm mới dữ liệu từ SQL!');
            }}
            return;
          }}
        }} catch (e) {{}}
      }}
    }}

    function exportCSV() {{
      const items = getDateMonthFilteredItems();
      if (items.length === 0) return alert('Không có dữ liệu để xuất!');

      let headers = [];
      if (currentBucket === 'RTS') {{
        if (rtsViewMode === 'items') {{
          headers = ['Order Number', 'Request ID', 'Mã NCC', 'SKU', 'SKU Name', 'QTY', 'Create Time', 'Ship Date', 'Status', 'Địa chỉ trả hàng', 'Địa chỉ chuyển', 'Thông tin xe', 'Mail'];
        }} else {{
          headers = ['Order Number', 'Request ID', 'Mã NCC', 'Total SKU', 'Total QTY', 'Create Time', 'Ship Date', 'Status', 'Địa chỉ trả hàng', 'Địa chỉ chuyển', 'Thông tin xe', 'Mail'];
        }}
      }} else if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
        headers = ['Type', 'Zone', 'Empty Qty', 'Used Qty', 'Location Qty', 'Max_capacity(m³)', 'Empty_Capacity(m³)', 'Volume Used(m³)', 'Qty Onhands', 'AVG Items available', 'Occupancy_Rate'];
      }} else if (currentBucket === 'Inbound') {{
        headers = ['PO ID', 'ASN ID', 'Receive Date', 'Expected Delivery', 'Total SKU', 'Quantity', 'Putaway', 'Backlog', 'Status', 'Arrived'];
      }} else {{
        headers = ['Ship Order', 'ASN', 'Receive Date', 'Total SKU', 'Quantity', 'Putaway', 'Backlog', 'Status', 'Arrived'];
      }}

      let csv = '\\uFEFF' + headers.join(',') + '\\n';
      items.forEach(r => {{
        if (currentBucket === 'RTS') {{
          const on = r['Order Number'] || '';
          const savedTransfer = JSON.parse(localStorage.getItem('rts_transfer_address') || localStorage.getItem('rts_return_address') || '{{}}');
          const retAddr = r['Return Address'] || r['Địa chỉ trả hàng'] || '';
          let trAddr = '';
          if (savedTransfer[on] !== undefined) {{
            const cand = String(savedTransfer[on]).trim();
            if (['VNNL', 'VNA', 'VNVL', 'VNG', 'VNSL'].includes(cand) || cand.length <= 15) trAddr = cand;
          }} else {{
            trAddr = r['Địa chỉ chuyển'] || r.transfer_address || '';
          }}
          let row = [];
          if (rtsViewMode === 'items') {{
            row = [
              `"${{r['Order Number'] || ''}}"`,
              `"${{r['Request ID'] || ''}}"`,
              `"${{r['Supplier'] || ''}}"`,
              `"${{r['SKU'] || ''}}"`,
              `"${{r['SKU Name'] || ''}}"`,
              r['QTY'] || 0,
              `"${{r['Create Time'] || ''}}"`,
              `"${{r['Ship Date'] || ''}}"`,
              `"${{r['Status'] || ''}}"`,
              `"${{retAddr}}"`,
              `"${{trAddr}}"`,
              `"${{r['Thông tin xe'] || ''}}"`,
              `"${{r['Mail'] || 'Chưa mail'}}"`
            ];
          }} else {{
            row = [
              `"${{r['Order Number'] || ''}}"`,
              `"${{r['Request ID'] || ''}}"`,
              `"${{r['Supplier'] || ''}}"`,
              r['Total SKU'] || 0,
              r['Total QTY'] || 0,
              `"${{r['Create Time'] || ''}}"`,
              `"${{r['Ship Date'] || ''}}"`,
              `"${{r['Status'] || ''}}"`,
              `"${{retAddr}}"`,
              `"${{trAddr}}"`,
              `"${{r['Thông tin xe'] || ''}}"`,
              `"${{r['Mail'] || 'Chưa mail'}}"`
            ];
          }}
          csv += row.join(',') + '\\n';
        }} else if (currentBucket === 'Overview' || currentBucket === 'Over view VNDL') {{
          const row = [
            `"${{r['Type'] || ''}}"`,
            `"${{r['Zone'] || ''}}"`,
            r['Empty Qty'] || 0,
            r['Used Qty'] || 0,
            r['Location Qty'] || 0,
            r['Max_capacity(m³)'] || 0,
            r['Empty_Capacity(m³)'] || 0,
            r['Volume Used(m³)'] || 0,
            r['Qty Onhands'] || 0,
            r['AVG Items available'] || 0,
            `"${{r['Occupancy_Rate'] || ''}}%"`
          ];
          csv += row.join(',') + '\\n';
        }} else if (currentBucket === 'Inbound') {{
          const row = [
            `"${{r['PO ID'] || ''}}"`,
            `"${{r['ASN ID'] || ''}}"`,
            `"${{r['Receive Date'] || ''}}"`,
            `"${{r['Expected Delivery'] || ''}}"`,
            r['Total SKU'] || 0,
            r['Quantity'] || 0,
            r['Putaway'] || 0,
            r['Backlog'] || 0,
            `"${{r['Status'] || ''}}"`,
            `"${{r['Arrived'] || ''}}"`
          ];
          csv += row.join(',') + '\\n';
        }} else {{
          const val1 = r['Ship Order'] || r['PO ID'] || '';
          const val2 = r['ASN'] || r['ASN ID'] || '';
          const row = [
            `"${{val1}}"`,
            `"${{val2}}"`,
            `"${{r['Receive Date'] || ''}}"`,
            r['Total SKU'] || 0,
            r['Quantity'] || 0,
            r['Putaway'] || 0,
            r['Backlog'] || 0,
            `"${{r['Status'] || ''}}"`,
            `"${{r['Arrived'] || ''}}"`
          ];
          csv += row.join(',') + '\\n';
        }}
      }});
      const blob = new Blob([csv], {{ type: 'text/csv;charset=utf-8;' }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `Inbound_${{currentBucket}}_Summary_${{new Date().toISOString().slice(0,10)}}.csv`;
      a.click();
    }}

    initCalendarState();
    renderFilterTabs();
    renderTableHead();
    populateDateMonthSelectors();
    const _mtiBox = document.getElementById('dataMtiSearchBox');
    if (_mtiBox) _mtiBox.style.display = (currentBucket === 'MTI') ? 'flex' : 'none';
    updateUI();
    document.addEventListener('click', function(e) {{
      const pop = document.getElementById('multiDatePopover');
      const trig = document.getElementById('multiDatePickerTrigger');
      if (pop && pop.classList.contains('show')) {{
        if (!pop.contains(e.target) && !trig.contains(e.target)) {{
          pop.classList.remove('show');
        }}
      }}
    }});
    if (window.location.pathname.includes('truck_tracker.html')) {{
      setTimeout(openTruckTrackerModal, 150);
    }}
    setInterval(() => fetchLiveSQLData(false), 10000);
  </script>
</body>
</html>
"""
    return html


def save_html_file(target_path=HTML_FILE, data=None, initial_bucket='MTI'):
    """Lưu mã nguồn Dashboard ra file HTML"""
    content = generate_html_content(data=data, initial_bucket=initial_bucket)
    dir_name = os.path.dirname(target_path)
    if dir_name and not os.path.exists(dir_name):
        os.makedirs(dir_name, exist_ok=True)
    with open(target_path, 'w', encoding='utf-8') as f:
        f.write(content)
    return target_path


def save_all_html_copies():
    """Đồng bộ tất cả các file index.html và rts.html trên máy"""
    data = query_inbound_backlog()
    for p in [
        HTML_FILE,
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\index.html',
        r'C:\Users\srvn0514\Desktop\SQL\index.html',
        r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\index.html',
        r'c:\Users\srvn0514\Desktop\mới\index.html'
    ]:
        try:
            save_html_file(p, data=data, initial_bucket='MTI')
        except Exception as e:
            print(f"[Dashboard HTML] Lỗi lưu {p}: {e}")

    # Tạo riêng file rts.html mở mặc định tab RTS
    for p_rts in [
        RTS_HTML_FILE,
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\rts.html',
        r'C:\Users\srvn0514\Desktop\SQL\rts.html',
        r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\rts.html',
        r'c:\Users\srvn0514\Desktop\mới\rts.html'
    ]:
        try:
            save_html_file(p_rts, data=data, initial_bucket='RTS')
        except Exception as e:
            print(f"[Dashboard HTML] Lỗi lưu {p_rts}: {e}")

    # Tạo riêng file overview_vndl.html mở mặc định tab Over view VNDL
    for p_ov in [
        OVERVIEW_HTML_FILE,
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\overview_vndl.html',
        r'C:\Users\srvn0514\Desktop\SQL\overview_vndl.html',
        r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\overview_vndl.html',
        r'c:\Users\srvn0514\Desktop\mới\overview_vndl.html'
    ]:
        try:
            save_html_file(p_ov, data=data, initial_bucket='Overview')
        except Exception as e:
            print(f"[Dashboard HTML] Lỗi lưu {p_ov}: {e}")

    # Tạo riêng file truck_tracker.html mở mặc định popup Daily Truck Tracker
    for p_trk in [
        TRUCK_HTML_FILE,
        r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard\truck_tracker.html',
        r'C:\Users\srvn0514\Desktop\SQL\truck_tracker.html',
        r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL\truck_tracker.html',
        r'c:\Users\srvn0514\Desktop\mới\truck_tracker.html'
    ]:
        try:
            save_html_file(p_trk, data=data, initial_bucket='MTI')
        except Exception as e:
            print(f"[Dashboard HTML] Lỗi lưu {p_trk}: {e}")


class DashboardHTTPRequestHandler(SimpleHTTPRequestHandler):
    """Xử lý các request HTTP cục bộ và cập nhật ngược về SQLite"""

    def end_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', '*')
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.end_headers()

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/api/set_arrived':
            try:
                qs = parse_qs(parsed.query)
                ship_order = qs.get('ship_order', [''])[0]
                arrived = qs.get('arrived', ['No'])[0]
                bucket = qs.get('bucket', ['MTI'])[0]
                if ship_order:
                    if bucket == 'Inbound':
                        mod_ib = _get_compute_ib_module()
                        mod_ib.set_ib_order_arrived(ship_order, arrived, DB_PATH)
                    else:
                        mod = _get_compute_module()
                        mod.set_order_arrived(ship_order, arrived, DB_PATH)
                    save_all_html_copies()
                body = json.dumps({"ok": True, "ship_order": ship_order, "arrived": arrived, "bucket": bucket}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path == '/api/set_rts_truck_info':
            try:
                qs = parse_qs(parsed.query)
                order_number = qs.get('order_number', [''])[0]
                truck_info = qs.get('truck_info', [''])[0]
                if order_number:
                    mod_rts = _get_compute_rts_module()
                    if hasattr(mod_rts, 'set_rts_truck_info'):
                        mod_rts.set_rts_truck_info(order_number, truck_info, DB_PATH)
                    save_all_html_copies()
                body = json.dumps({"ok": True, "order_number": order_number, "truck_info": truck_info}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path == '/api/set_rts_status':
            try:
                qs = parse_qs(parsed.query)
                order_number = qs.get('order_number', [''])[0]
                sku = qs.get('sku', [''])[0]
                status = qs.get('status', ['Allocated'])[0]
                if order_number:
                    mod_rts = _get_compute_rts_module()
                    mod_rts.set_rts_status(order_number, sku, status, DB_PATH)
                    save_all_html_copies()
                body = json.dumps({"ok": True, "order_number": order_number, "sku": sku, "status": status}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path == '/api/set_rts_truck_info':
            try:
                qs = parse_qs(parsed.query)
                order_number = qs.get('order_number', [''])[0]
                truck_info = qs.get('truck_info', [''])[0]
                if order_number:
                    mod_rts = _get_compute_rts_module()
                    if hasattr(mod_rts, 'set_rts_truck_info'):
                        mod_rts.set_rts_truck_info(order_number, truck_info, DB_PATH)
                    save_all_html_copies()
                body = json.dumps({"ok": True, "order_number": order_number, "truck_info": truck_info}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path == '/api/set_rts_mail_status':
            try:
                qs = parse_qs(parsed.query)
                order_number = qs.get('order_number', [''])[0]
                mail_status = qs.get('mail_status', ['Chưa mail'])[0]
                if order_number:
                    mod_rts = _get_compute_rts_module()
                    if hasattr(mod_rts, 'set_rts_mail_status'):
                        mod_rts.set_rts_mail_status(order_number, mail_status, DB_PATH)
                    save_all_html_copies()
                body = json.dumps({"ok": True, "order_number": order_number, "mail_status": mail_status}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path in ('/api/set_rts_address', '/api/set_rts_transfer_address'):
            try:
                qs = parse_qs(parsed.query)
                order_number = qs.get('order_number', [''])[0]
                address = qs.get('address', qs.get('transfer_address', ['']))[0]
                if order_number:
                    mod_rts = _get_compute_rts_module()
                    if hasattr(mod_rts, 'set_rts_transfer_address'):
                        mod_rts.set_rts_transfer_address(order_number, address, DB_PATH)
                    elif hasattr(mod_rts, 'set_rts_return_address'):
                        mod_rts.set_rts_return_address(order_number, address, DB_PATH)
                    save_all_html_copies()
                body = json.dumps({"ok": True, "order_number": order_number, "address": address}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path == '/api/set_truck_shipping_id':
            try:
                qs = parse_qs(parsed.query)
                row_id = qs.get('rowid', qs.get('id', ['']))[0].strip()
                trip_date = qs.get('date', [''])[0].strip()
                chuyen = qs.get('chuyen', [''])[0].strip()
                so_xe = qs.get('so_xe', [''])[0].strip()
                shipping_id = qs.get('shipping_id', [''])[0].strip()

                conn = get_db_connection()
                cur = conn.cursor()
                cur.execute("""
                    CREATE TABLE IF NOT EXISTS _truck_shipping_ids_overrides (
                        trip_key TEXT PRIMARY KEY,
                        date_val TEXT,
                        chuyen TEXT,
                        so_xe TEXT,
                        shipping_id TEXT,
                        updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                    )
                """)
                trip_key = f"{trip_date}_{chuyen}_{so_xe}".strip().lower()
                cur.execute("""
                    INSERT INTO _truck_shipping_ids_overrides (trip_key, date_val, chuyen, so_xe, shipping_id, updated_at)
                    VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(trip_key) DO UPDATE SET
                        shipping_id = excluded.shipping_id,
                        updated_at = CURRENT_TIMESTAMP
                """, (trip_key, trip_date, chuyen, so_xe, shipping_id))
                conn.commit()

                cur.execute("PRAGMA table_info([Daily truck tracker])")
                cols = [c[1] for c in cur.fetchall()]
                target_col = 'Shipping ID'
                if 'Shipping ID' not in cols:
                    if 'Shiping ID' in cols:
                        target_col = 'Shiping ID'
                    else:
                        cur.execute("ALTER TABLE [Daily truck tracker] ADD COLUMN [Shipping ID] TEXT")
                        conn.commit()
                        target_col = 'Shipping ID'

                updated = False
                for date_col in ['Ngày MTO', 'Ngay MTO', 'Dự kiến giao hàng', 'Ngày', 'Ngay']:
                    if date_col in cols:
                        for chuyen_col in ['Số chuyến', 'So chuyen', 'Chuyến', 'Chuyen']:
                            if chuyen_col in cols and trip_date and chuyen:
                                cur.execute(f"UPDATE [Daily truck tracker] SET [{target_col}] = ? WHERE [{date_col}] = ? AND [{chuyen_col}] = ?", (shipping_id, trip_date, chuyen))
                                conn.commit()
                                if cur.rowcount > 0:
                                    updated = True
                                    break
                        if updated:
                            break

                if not updated and so_xe:
                    for sx_col in ['Số xe', 'So xe', 'Biển số']:
                        if sx_col in cols:
                            cur.execute(f"UPDATE [Daily truck tracker] SET [{target_col}] = ? WHERE [{sx_col}] = ?", (shipping_id, so_xe))
                            conn.commit()
                            if cur.rowcount > 0:
                                updated = True
                                break

                if not updated and row_id and str(row_id).isdigit():
                    cur.execute(f"UPDATE [Daily truck tracker] SET [{target_col}] = ? WHERE rowid = ?", (shipping_id, int(row_id)))
                    conn.commit()

                conn.close()

                # Nếu nhập Shipping ID, tự động cập nhật đơn hàng MTI sang Yes (Đã đến)
                if shipping_id:
                    mod = _get_compute_module()
                    if mod and hasattr(mod, 'set_order_arrived'):
                        mod.set_order_arrived(shipping_id, 'Yes', DB_PATH)

                save_all_html_copies()
                body = json.dumps({"ok": True, "shipping_id": shipping_id, "rowid": row_id}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path == '/api/add_xe_den_vndl':
            try:
                qs = parse_qs(parsed.query)
                so_xe = qs.get('so_xe', [''])[0].strip()
                sdt = qs.get('sdt', [''])[0].strip()
                tai_xe = qs.get('tai_xe', [''])[0].strip()
                pallet_raw = qs.get('pallet', ['0'])[0].strip()
                pallet_num = int(pallet_raw) if pallet_raw.isdigit() else 0
                ngay = qs.get('ngay', [''])[0].strip() or time.strftime('%d/%m/%Y')
                kho_di = qs.get('kho_di', [''])[0].strip()
                ghi_chu = qs.get('ghi_chu', [''])[0].strip()
                inout_mode = qs.get('inout_mode', ['IN'])[0].strip().upper() or 'IN'
                type_val = qs.get('type_val', [''])[0].strip()
                operater = qs.get('operater', [''])[0].strip()
                warehouse = qs.get('warehouse', [''])[0].strip()
                shipping_id_val = qs.get('shipping_id', [''])[0].strip()
                gio_den = time.strftime('%H:%M:%S')
                ts_val = f'{ngay} {gio_den}'.strip()

                # Nhận dữ liệu cả 2 chiều IN và OUT
                in_type = qs.get('in_type', [''])[0].strip() or (type_val if inout_mode == 'IN' else '')
                in_pallet = qs.get('in_pallet', [''])[0].strip() or (str(pallet_num) if inout_mode == 'IN' and pallet_num > 0 else '')
                in_from_wh = qs.get('in_wh', [''])[0].strip() or (warehouse if inout_mode == 'IN' else '')
                in_operater = qs.get('in_op', [''])[0].strip() or (operater if inout_mode == 'IN' else '')
                in_ldd = qs.get('in_ldd', [''])[0].strip() or (ghi_chu if inout_mode == 'IN' else '')

                out_type = qs.get('out_type', [''])[0].strip() or (type_val if inout_mode == 'OUT' else '')
                out_pallet = qs.get('out_pallet', [''])[0].strip() or (str(pallet_num) if inout_mode == 'OUT' and pallet_num > 0 else '')
                out_to_wh = qs.get('out_wh', [''])[0].strip() or (warehouse if inout_mode == 'OUT' else '')
                out_operater = qs.get('out_op', [''])[0].strip() or (operater if inout_mode == 'OUT' else '')
                out_ldd = qs.get('out_ldd', [''])[0].strip() or (ghi_chu if inout_mode == 'OUT' else '')

                has_in = bool(in_type or in_pallet or in_from_wh or in_operater)
                has_out = bool(out_type or out_pallet or out_to_wh or out_operater)

                # Tính toán Leadtime
                leadtime_val = qs.get('leadtime', [''])[0].strip() or 'On Time'
                if not qs.get('leadtime'):
                    today_d = time.strftime('%d/%m/%Y')
                    try:
                        p_ngay = ngay.split('/')
                        p_today = today_d.split('/')
                        if len(p_ngay) == 3 and len(p_today) == 3:
                            sk_ngay = f"{p_ngay[2]}{p_ngay[1]}{p_ngay[0]}"
                            sk_today = f"{p_today[2]}{p_today[1]}{p_today[0]}"
                            if sk_ngay > sk_today:
                                leadtime_val = 'Late'
                    except Exception:
                        pass

                # Lưu vào SQLite
                conn = get_db_connection()
                with conn:
                    cur = conn.cursor()
                    try:
                        cur.execute("ALTER TABLE [xe đến VNDL] ADD COLUMN [Leadtime] TEXT DEFAULT 'On Time'")
                    except Exception:
                        pass
                    cols_in_tbl = [c[1] for c in cur.execute('PRAGMA table_info([xe đến VNDL])').fetchall()]
                    if 'Biển số xe' in cols_in_tbl:
                        # Tìm dòng xe khớp biển số (chuẩn hóa bỏ dấu/gạch)
                        clean_so_xe = re.sub(r'[^a-zA-Z0-9]', '', so_xe).lower()
                        all_rows_xd = cur.execute("SELECT rowid, [Biển số xe], [Dấu thời gian] FROM [xe đến VNDL]").fetchall()
                        existing_row = None
                        for rxd in reversed(all_rows_xd):
                            p_clean = re.sub(r'[^a-zA-Z0-9]', '', str(rxd[1] or '')).lower()
                            if p_clean and (p_clean == clean_so_xe or clean_so_xe in p_clean or p_clean in clean_so_xe):
                                if ngay in str(rxd[2] or ''):
                                    existing_row = (rxd[0],)
                                    break
                                elif existing_row is None:
                                    existing_row = (rxd[0],)
                        
                        has_ship_col = 'Shipping ID' in cols_in_tbl
                        has_lead_col = 'Leadtime' in cols_in_tbl

                        if existing_row:
                            # UPDATE dòng đã có - cập nhật cả IN và OUT nếu có dữ liệu
                            set_parts = []
                            set_vals = []
                            if has_in:
                                set_parts.extend(["[IN Type] = ?", "[IN Pallet] = ?", "[IN From WH] = ?", "[IN Operater] = ?", "[IN LDD] = ?"])
                                set_vals.extend([in_type, in_pallet, in_from_wh, in_operater, in_ldd])
                            if has_out:
                                set_parts.extend(["[OUT Type] = ?", "[OUT Pallet] = ?", "[OUT To WH] = ?", "[OUT Operater] = ?", "[OUT LDD] = ?"])
                                set_vals.extend([out_type, out_pallet, out_to_wh, out_operater, out_ldd])
                            
                            if has_ship_col and shipping_id_val:
                                set_parts.append("[Shipping ID] = ?")
                                set_vals.append(shipping_id_val)
                            
                            if has_lead_col:
                                set_parts.append("[Leadtime] = ?")
                                set_vals.append(leadtime_val)
                            
                            if set_parts:
                                set_sql = ", ".join(set_parts)
                                set_vals.append(existing_row[0])
                                cur.execute(f"UPDATE [xe đến VNDL] SET {set_sql} WHERE rowid = ?", set_vals)
                                print(f"[Xe đến VNDL] Updated rowid {existing_row[0]}: {so_xe} IN: {in_type} {in_pallet}plt, OUT: {out_type} {out_pallet}plt, Leadtime: {leadtime_val}")
                        else:
                            # INSERT dòng mới với cả 2 chiều IN và OUT
                            icols = ['[Dấu thời gian]','[Biển số xe]','[Name]','[Số ĐT]','[Kho xuất phát]',
                                     '[IN Type]','[IN Pallet]','[IN From WH]','[IN Operater]','[IN LDD]',
                                     '[OUT Type]','[OUT Pallet]','[OUT To WH]','[OUT Operater]','[OUT LDD]']
                            ivals = [ts_val, so_xe, tai_xe, sdt, kho_di,
                                     in_type, in_pallet, in_from_wh, in_operater, in_ldd,
                                     out_type, out_pallet, out_to_wh, out_operater, out_ldd]
                            if has_ship_col:
                                icols.append('[Shipping ID]')
                                ivals.append(shipping_id_val)
                            if has_lead_col:
                                icols.append('[Leadtime]')
                                ivals.append(leadtime_val)
                            phs = ', '.join(['?'] * len(icols))
                            cur.execute(f'INSERT INTO [xe đến VNDL] ({", ".join(icols)}) VALUES ({phs})', ivals)
                            print(f"[Xe đến VNDL] Inserted new: {so_xe} IN: {in_type} {in_pallet}plt, OUT: {out_type} {out_pallet}plt, Leadtime: {leadtime_val}")
                    else:
                        cur.execute('INSERT INTO [xe đến VNDL] (ngay, gio_den, so_xe, sdt, tai_xe, pallet, kho_di, ghi_chu) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
                                    (ngay, gio_den, so_xe, sdt, tai_xe, pallet_num, kho_di, ghi_chu))
                conn.close()

                # Tự động cập nhật Arrived = Yes trong SQLite nếu có shipping_id
                if shipping_id_val:
                    tokens = [t.strip() for t in re.split(r'[\s,;\t]+', shipping_id_val) if len(t.strip()) >= 4]
                    for tok in tokens:
                        try:
                            mod_mti = _get_compute_module()
                            mod_mti.set_order_arrived(tok, 'Yes', DB_PATH)
                        except Exception:
                            pass
                        try:
                            mod_ib = _get_compute_ib_module()
                            mod_ib.set_ib_order_arrived(tok, 'Yes', DB_PATH)
                        except Exception:
                            pass
                    print(f"[Arrived] Auto-set Arrived=Yes for shipping tokens: {tokens}")

                # Đẩy lên Google Sheet (chạy background để không block response)
                def _gs_write_xe_den(so_xe=so_xe, has_in=has_in, has_out=has_out,
                                     in_type=in_type, in_pallet=in_pallet, in_from_wh=in_from_wh,
                                     in_operater=in_operater, in_ldd=in_ldd,
                                     out_type=out_type, out_pallet=out_pallet, out_to_wh=out_to_wh,
                                     out_operater=out_operater, out_ldd=out_ldd,
                                     shipping_id_val=shipping_id_val, leadtime_val=leadtime_val,
                                     ts_val=ts_val, tai_xe=tai_xe, sdt=sdt, kho_di=kho_di):
                    try:
                        import gspread
                        from google.oauth2.service_account import Credentials
                        c_path = os.path.join(os.path.dirname(DB_PATH), 'credentials.json')
                        if not os.path.exists(c_path):
                            c_path = os.path.join(BASE_DIR, 'credentials.json')
                        if not os.path.exists(c_path):
                            print('[GSheet] Không tìm thấy credentials.json')
                            return
                        sc = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
                        gclient = gspread.authorize(Credentials.from_service_account_file(c_path, scopes=sc))
                        gws = gclient.open_by_key('1vOUEjQGjmbYn0jEYJZZy2AdZchetBiIdUFIj47NxMAI').worksheet('Trang tính1')
                        all_vals = gws.get_all_values()
                        target_row_idx = None
                        clean_so_xe = re.sub(r'[^a-zA-Z0-9]', '', so_xe).lower()
                        for r_idx in range(len(all_vals), 2, -1):
                            r_val = all_vals[r_idx - 1]
                            if len(r_val) > 1:
                                row_plate = re.sub(r'[^a-zA-Z0-9]', '', str(r_val[1])).lower()
                                if row_plate and (row_plate == clean_so_xe or clean_so_xe in row_plate or row_plate in clean_so_xe):
                                    target_row_idx = r_idx
                                    break

                        if target_row_idx:
                            curr_row_cells = all_vals[target_row_idx - 1]
                            curr_h_u = (curr_row_cells[7:21] + [''] * 14)[:14]
                            if has_in:
                                curr_h_u[0] = in_type or curr_h_u[0]
                                curr_h_u[1] = in_pallet or curr_h_u[1]
                                curr_h_u[2] = in_from_wh or curr_h_u[2]
                                curr_h_u[4] = in_operater or curr_h_u[4]
                                curr_h_u[5] = in_ldd or curr_h_u[5]
                            if has_out:
                                curr_h_u[6] = out_type or curr_h_u[6]
                                curr_h_u[7] = out_pallet or curr_h_u[7]
                                curr_h_u[8] = out_to_wh or curr_h_u[8]
                                curr_h_u[10] = out_operater or curr_h_u[10]
                                curr_h_u[11] = out_ldd or curr_h_u[11]
                            if shipping_id_val:
                                curr_h_u[12] = shipping_id_val
                            curr_h_u[13] = leadtime_val
                            gws.update([curr_h_u], f'H{target_row_idx}:U{target_row_idx}')
                            print(f'[GSheet] Cập nhật dòng {target_row_idx}: {so_xe} IN:{curr_h_u[0]} {curr_h_u[1]}plt, OUT:{curr_h_u[6]} {curr_h_u[7]}plt, Lead:{leadtime_val}')
                        else:
                            row_to_append = [
                                ts_val, so_xe, tai_xe, sdt, kho_di, '8 Tấn', 'YES',
                                in_type, in_pallet, in_from_wh, '', in_operater, in_ldd,
                                out_type, out_pallet, out_to_wh, '', out_operater, out_ldd,
                                shipping_id_val, leadtime_val
                            ]
                            gws.append_row(row_to_append)
                            print(f'[GSheet] Append mới: {so_xe} IN:{in_type} {in_pallet}plt, Leadtime:{leadtime_val}')
                    except Exception as e_gs:
                        print(f'[GSheet] Lỗi xe đến: {e_gs}')

                threading.Thread(target=_gs_write_xe_den, daemon=True).start()
                threading.Thread(target=save_all_html_copies, daemon=True).start()
                body = json.dumps({'ok': True, 'so_xe': so_xe, 'pallet': pallet_num, 'inout': inout_mode, 'leadtime': leadtime_val}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({'ok': False, 'error': str(e)}).encode('utf-8'))

        elif parsed.path == '/api/set_rts_truck_info':
            try:
                qs = parse_qs(parsed.query)
                order_number = qs.get('order_number', [''])[0]
                truck_info = qs.get('truck_info', [''])[0]
                if order_number:
                    mod_rts = _get_compute_rts_module()
                    if hasattr(mod_rts, 'set_rts_truck_info'):
                        mod_rts.set_rts_truck_info(order_number, truck_info, DB_PATH)
                    save_all_html_copies()
                body = json.dumps({"ok": True, "order_number": order_number, "truck_info": truck_info}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path == '/api/set_rts_mail_status':
            try:
                qs = parse_qs(parsed.query)
                order_number = qs.get('order_number', [''])[0]
                mail_status = qs.get('mail_status', ['Chưa mail'])[0]
                if order_number:
                    mod_rts = _get_compute_rts_module()
                    if hasattr(mod_rts, 'set_rts_mail_status'):
                        mod_rts.set_rts_mail_status(order_number, mail_status, DB_PATH)
                    save_all_html_copies()
                body = json.dumps({"ok": True, "order_number": order_number, "mail_status": mail_status}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path in ('/api/set_rts_address', '/api/set_rts_transfer_address'):
            try:
                qs = parse_qs(parsed.query)
                order_number = qs.get('order_number', [''])[0]
                address = qs.get('address', qs.get('transfer_address', ['']))[0]
                if order_number:
                    mod_rts = _get_compute_rts_module()
                    if hasattr(mod_rts, 'set_rts_transfer_address'):
                        mod_rts.set_rts_transfer_address(order_number, address, DB_PATH)
                    elif hasattr(mod_rts, 'set_rts_return_address'):
                        mod_rts.set_rts_return_address(order_number, address, DB_PATH)
                    save_all_html_copies()
                body = json.dumps({"ok": True, "order_number": order_number, "address": address}).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(json.dumps({"ok": False, "error": str(e)}).encode('utf-8'))

        elif parsed.path == '/api/data':
            try:
                data = query_inbound_backlog()
                body = json.dumps(data, ensure_ascii=False).encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'application/json; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(f'{{"error": "{e}"}}'.encode('utf-8'))

        elif parsed.path in ['/', '/index.html']:
            try:
                content = generate_html_content(initial_bucket='MTI')
                body = content.encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(f"Lỗi tạo Dashboard: {e}".encode('utf-8'))

        elif parsed.path == '/rts.html':
            try:
                content = generate_html_content(initial_bucket='RTS')
                body = content.encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(f"Lỗi tạo RTS Dashboard: {e}".encode('utf-8'))

        elif parsed.path in ['/overview_vndl.html', '/overview.html']:
            try:
                content = generate_html_content(initial_bucket='Overview')
                body = content.encode('utf-8')
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', str(len(body)))
                self.end_headers()
                self.wfile.write(body)
            except Exception as e:
                self.send_response(500)
                self.end_headers()
                self.wfile.write(f"Lỗi tạo Overview Dashboard: {e}".encode('utf-8'))

        else:
            super().do_GET()

    def log_message(self, format, *args):
        pass


def run_server(port=PORT, open_browser=True):
    """Khởi chạy máy chủ Web cục bộ phục vụ Dashboard Realtime"""
    save_all_html_copies()

    httpd = None
    active_port = port
    for candidate_port in [8501, port, 8502]:
        try:
            httpd = HTTPServer(('', candidate_port), DashboardHTTPRequestHandler)
            active_port = candidate_port
            break
        except OSError:
            continue

    if httpd is None:
        if open_browser:
            webbrowser.open(f"http://localhost:{port}")
        return

    url = f"http://localhost:{active_port}"
    print(f"DASHBOARD ĐANG CHẠY: {url}")

    if open_browser:
        webbrowser.open(url)

    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        httpd.server_close()


if __name__ == '__main__':
    open_in_browser = '--open' in sys.argv or len(sys.argv) == 1
    if '--static-only' in sys.argv:
        save_all_html_copies()
        print(f"Đã xuất Dashboard ra file: {HTML_FILE}")
    else:
        run_server(open_browser=open_in_browser)
