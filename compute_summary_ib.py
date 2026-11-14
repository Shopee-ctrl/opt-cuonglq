import os
import sys
import time
import sqlite3
from collections import defaultdict
from typing import Optional
import pandas as pd

if hasattr(sys, 'stdout') and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

DB_PATH = r'C:\Users\srvn0514\Desktop\SQL\Data MTI.db'


def _extract_date_ddmmyyyy(val: str) -> str:
    """Chuyển chuỗi ngày giờ (2026-10-19 21:14:19 hoặc 19/10/2026) về dạng DD/MM/YYYY chuẩn."""
    s = str(val or '').strip()
    if not s or s in ('None', 'nan', 'NaN', '-'):
        return ''
    if ' ' in s:
        s = s.split(' ')[0]
    if '-' in s:
        parts = s.split('-')
        if len(parts) == 3:
            if len(parts[0]) == 4:
                return f"{parts[2].zfill(2)}/{parts[1].zfill(2)}/{parts[0]}"
            else:
                return f"{parts[0].zfill(2)}/{parts[1].zfill(2)}/{parts[2]}"
    if '/' in s:
        parts = s.split('/')
        if len(parts) == 3:
            if len(parts[2]) == 4:
                return f"{parts[0].zfill(2)}/{parts[1].zfill(2)}/{parts[2]}"
            elif len(parts[0]) == 4:
                return f"{parts[2].zfill(2)}/{parts[1].zfill(2)}/{parts[0]}"
    return s


def _date_to_sortkey(s: str) -> str:
    s = str(s or '').strip()
    if '/' in s:
        parts = s.split('/')
        if len(parts) == 3:
            if len(parts[0]) == 4:
                return f"{parts[0]}{parts[1].zfill(2)}{parts[2].zfill(2)}"
            return f"{parts[2]}{parts[1].zfill(2)}{parts[0].zfill(2)}"
    if '-' in s:
        parts = s.split(' ')[0].split('-')
        if len(parts) == 3:
            if len(parts[0]) == 4:
                return f"{parts[0]}{parts[1].zfill(2)}{parts[2].zfill(2)}"
    return s


def compute_and_save_summary_ib(db_path: Optional[str] = None) -> pd.DataFrame:
    """
    Tổng hợp dữ liệu Inbound đối chiếu [data IB] và [Put IB]:
    - Nhóm theo ASN ID (và PO ID) từ [data IB]
    - Đối chiếu Putaway từ [Put IB] theo Inbound ID == ASN ID
    - Tính Quantity, Putaway, Backlog, Total SKU, Status, Arrived
    - Ghi vào bảng [Bao_Cao_Tong_Hop (IB)]
    """
    t0 = time.time()
    db_file = db_path or DB_PATH
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    # Tạo bảng lưu trạng thái tích Arrived riêng cho Inbound
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS _ib_order_arrived_state (
            asn_id TEXT PRIMARY KEY,
            arrived TEXT NOT NULL,
            arrived_date TEXT DEFAULT '',
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()

    cursor.execute("SELECT asn_id, arrived, COALESCE(arrived_date, '') FROM _ib_order_arrived_state")
    saved_arrived = {}
    saved_arrived_date = {}
    for r in cursor.fetchall():
        aid = str(r[0]).strip()
        saved_arrived[aid] = str(r[1]).strip()
        saved_arrived_date[aid] = str(r[2]).strip()

    # Đọc dữ liệu [data IB]
    cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='data IB'")
    if cursor.fetchone()[0] == 0:
        conn.close()
        return pd.DataFrame()

    df_ib = pd.read_sql_query("SELECT * FROM [data IB]", conn)
    if len(df_ib) == 0:
        conn.close()
        return pd.DataFrame()

    # Đọc dữ liệu [Put IB]
    cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='Put IB'")
    has_put_ib = (cursor.fetchone()[0] > 0)
    df_put_ib = pd.read_sql_query("SELECT * FROM [Put IB]", conn) if has_put_ib else pd.DataFrame()

    # Tính tổng Putaway từ [Put IB] theo Inbound ID
    put_by_inbound = defaultdict(int)
    if len(df_put_ib) > 0 and 'Inbound ID' in df_put_ib.columns and 'Quantity On Rack' in df_put_ib.columns:
        for _, r in df_put_ib.iterrows():
            inbound_id = str(r['Inbound ID']).strip()
            if not inbound_id or inbound_id in ('None', 'nan', ''):
                continue
            try:
                q = int(float(str(r['Quantity On Rack']).replace(',', '').strip()))
            except Exception:
                q = 0
            put_by_inbound[inbound_id] += q

    # Nhóm theo Inbound ID (ASN ID) trong [data IB]
    asns = {}
    for _, r in df_ib.iterrows():
        asn = str(r.get('Inbound ID', '') or r.get('ASN ID', '')).strip()
        if not asn or asn in ('None', 'nan', ''):
            continue
        po = str(r.get('PO ID', '')).strip()
        sku = str(r.get('SKU', '') or r.get('SKU ID', '')).strip()
        
        qty_val = r.get('Quantity', None)
        if qty_val is None or str(qty_val).strip() == '':
            qty_val = r.get('Expected Quantity', 0)
        try:
            exp_q = int(float(str(qty_val).replace(',', '').strip()))
        except Exception:
            exp_q = 0

        # Expected Delivery từ cột Expected Delivery Time
        exp_deliv_raw = str(r.get('Expected Delivery Time', '') or r.get('Expected Delivery', '')).strip()
        exp_deliv_date = _extract_date_ddmmyyyy(exp_deliv_raw)

        # Lấy ngày của đơn: ưu tiên Create time, Actual Arrival Time, hoặc Expected Delivery Time
        ctime = str(r.get('Create time', '')).strip()
        atime = str(r.get('Actual Arrival Time', '')).strip()
        date_str = _extract_date_ddmmyyyy(ctime) or _extract_date_ddmmyyyy(atime) or exp_deliv_date

        if asn not in asns:
            asns[asn] = {
                'asn': asn,
                'po_id': po,
                'receive_date': date_str,
                'expected_delivery': exp_deliv_date,
                'skus': set(),
                'quantity': 0,
                'putaway': put_by_inbound.get(asn, 0)
            }
        asns[asn]['quantity'] += exp_q
        if sku:
            asns[asn]['skus'].add(sku)
        if not asns[asn]['po_id'] and po:
            asns[asn]['po_id'] = po
        if not asns[asn]['receive_date'] and date_str:
            asns[asn]['receive_date'] = date_str
        if not asns[asn]['expected_delivery'] and exp_deliv_date:
            asns[asn]['expected_delivery'] = exp_deliv_date

    today_ddmmyyyy = time.strftime("%d/%m/%Y")
    rows = []

    for asn, o in asns.items():
        qty = o['quantity']
        put = o['putaway']
        backlog = max(0, qty - put)

        if saved_arrived.get(asn) == 'Done':
            arrived_val = 'Done'
            status = 'Done'
        elif put == 0:
            arrived_val = saved_arrived.get(asn, 'No')
            status = 'Arriving' if arrived_val == 'Yes' else 'Pending'
            if asn not in saved_arrived:
                conn.execute("""
                    INSERT OR IGNORE INTO _ib_order_arrived_state (asn_id, arrived, arrived_date, updated_at)
                    VALUES (?, 'No', '', CURRENT_TIMESTAMP)
                """, (asn,))
        else:
            arrived_val = 'Yes'
            if saved_arrived.get(asn) not in ('Yes', 'Done'):
                saved_arrived[asn] = 'Yes'
                saved_arrived_date[asn] = today_ddmmyyyy
                conn.execute("""
                    INSERT INTO _ib_order_arrived_state (asn_id, arrived, arrived_date, updated_at)
                    VALUES (?, 'Yes', ?, CURRENT_TIMESTAMP)
                    ON CONFLICT(asn_id) DO UPDATE SET arrived='Yes', arrived_date=excluded.arrived_date, updated_at=CURRENT_TIMESTAMP
                """, (asn, today_ddmmyyyy))
            if backlog <= 700 or put >= qty:
                status = 'Done'
                arrived_val = 'Done'
                if saved_arrived.get(asn) != 'Done':
                    saved_arrived[asn] = 'Done'
                    conn.execute("""
                        INSERT INTO _ib_order_arrived_state (asn_id, arrived, arrived_date, updated_at)
                        VALUES (?, 'Done', ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(asn_id) DO UPDATE SET arrived='Done', updated_at=CURRENT_TIMESTAMP
                    """, (asn, today_ddmmyyyy))
            else:
                status = 'In Progress'

        rec_date = o['receive_date']
        if arrived_val in ('Yes', 'Done') and asn in saved_arrived_date and saved_arrived_date[asn]:
            rec_date = saved_arrived_date[asn]

        # Nếu status là Done thì Arrived cũng là Done và ngược lại, backlog về = 0
        if status == 'Done' or arrived_val == 'Done':
            status = 'Done'
            arrived_val = 'Done'
            backlog = 0

        rows.append({
            'PO ID': o['po_id'],
            'ASN ID': asn,
            'Receive Date': rec_date or today_ddmmyyyy,
            'Expected Delivery': o['expected_delivery'],
            'Total SKU': len(o['skus']),
            'Quantity': qty,
            'Putaway': put,
            'Backlog': backlog,
            'Status': status,
            'Arrived': arrived_val
        })

    df_res = pd.DataFrame(rows)
    if len(df_res) > 0:
        status_order = {'In Progress': 1, 'Arriving': 2, 'Pending': 3, 'Done': 4}
        df_res['sort_status'] = df_res['Status'].map(status_order).fillna(5)
        df_res = df_res.sort_values(
            by=['sort_status', 'Receive Date', 'Quantity'],
            ascending=[True, False, False]
        ).drop(columns=['sort_status'])

    # Lưu vào bảng [Bao_Cao_Tong_Hop (IB)]
    df_res.to_sql('Bao_Cao_Tong_Hop (IB)', conn, if_exists='replace', index=False)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_bcth_ib_asn ON [Bao_Cao_Tong_Hop (IB)] ([ASN ID])")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_bcth_ib_status ON [Bao_Cao_Tong_Hop (IB)] ([Status])")
    conn.commit()
    conn.close()

    dur = time.time() - t0
    print(f"Computed & saved {len(df_res)} orders to [Bao_Cao_Tong_Hop (IB)] in {dur:.2f}s")
    return df_res


def set_ib_order_arrived(asn_id: str, arrived_val: str, db_path: Optional[str] = None):
    """Cập nhật trạng thái Arrived (Yes/No) cho đơn Inbound."""
    db_file = db_path or DB_PATH
    conn = sqlite3.connect(db_file)
    today_ddmmyyyy = time.strftime("%d/%m/%Y")
    arrived_date = today_ddmmyyyy if arrived_val == 'Yes' else ''

    with conn:
        conn.execute("""
            INSERT INTO _ib_order_arrived_state (asn_id, arrived, arrived_date, updated_at)
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(asn_id) DO UPDATE SET 
                arrived = excluded.arrived,
                arrived_date = excluded.arrived_date,
                updated_at = CURRENT_TIMESTAMP
        """, (asn_id, arrived_val, arrived_date))
    conn.close()
    compute_and_save_summary_ib(db_file)


if __name__ == '__main__':
    df = compute_and_save_summary_ib()
    print(df.head(10))
