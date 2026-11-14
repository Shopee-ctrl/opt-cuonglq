import os
import sys
import time
import json
import sqlite3
import pandas as pd
from collections import defaultdict

DB_PATH = r'C:\Users\srvn0514\Desktop\SQL\Data MTI.db'


def _date_to_sortkey(s: str) -> str:
    s = str(s or '').strip()
    if '/' in s:
        parts = s.split('/')
        if len(parts) == 3:
            if len(parts[0]) == 4:
                return f"{parts[0]}{parts[1].zfill(2)}{parts[2].zfill(2)}"
            y = ('20' + parts[2]) if len(parts[2]) == 2 else parts[2]
            return f"{y}{parts[1].zfill(2)}{parts[0].zfill(2)}"
    if '-' in s:
        parts = s[:10].split('-')
        if len(parts) == 3:
            if len(parts[0]) == 4:
                return f"{parts[0]}{parts[1].zfill(2)}{parts[2].zfill(2)}"
            return f"{parts[2]}{parts[1].zfill(2)}{parts[0].zfill(2)}"
    return s


def ensure_arrived_table(conn: sqlite3.Connection):
    conn.execute("""
        CREATE TABLE IF NOT EXISTS _order_arrived_state (
            ship_order TEXT PRIMARY KEY,
            arrived TEXT,
            arrived_date TEXT,
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    cursor = conn.cursor()
    cursor.execute("PRAGMA table_info(_order_arrived_state)")
    cols = [c[1] for c in cursor.fetchall()]
    if 'arrived_date' not in cols:
        conn.execute("ALTER TABLE _order_arrived_state ADD COLUMN arrived_date TEXT")
    conn.commit()


def set_order_arrived(ship_order: str, arrived: str, db_path: str = DB_PATH):
    """Cập nhật trạng thái Arrived (Yes/No/Done) từ HTML hoặc App ngược về SQL"""
    ship_order = str(ship_order).strip()
    arr_raw = str(arrived).strip().lower()
    if arr_raw == 'done':
        arrived = 'Done'
    elif arr_raw == 'yes':
        arrived = 'Yes'
    else:
        arrived = 'No'
    today_ddmmyyyy = time.strftime("%d/%m/%Y")
    today_key = time.strftime("%Y%m%d")

    conn = sqlite3.connect(db_path)
    ensure_arrived_table(conn)
    cursor = conn.cursor()

    # Lấy ngày gốc từ data_mti
    orig_date = ''
    try:
        cursor.execute("SELECT MIN(date) FROM data_mti WHERE trim(ship_order) = ?", (ship_order,))
        r_orig = cursor.fetchone()
        if r_orig and r_orig[0]:
            orig_date = str(r_orig[0]).strip()
    except Exception:
        pass

    if arrived in ('Yes', 'Done'):
        # Khi bấm chọn Arrived = Yes hoặc Done vào ngày nào (ngày D), Receive Date nhảy sang đúng ngày D đó luôn
        effective_date = today_ddmmyyyy
        arrived_date_to_save = today_ddmmyyyy
    else:
        effective_date = orig_date
        arrived_date_to_save = ''

    conn.execute("""
        INSERT INTO _order_arrived_state (ship_order, arrived, arrived_date, updated_at)
        VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        ON CONFLICT(ship_order) DO UPDATE SET
            arrived = excluded.arrived,
            arrived_date = excluded.arrived_date,
            updated_at = CURRENT_TIMESTAMP
    """, (ship_order, arrived, arrived_date_to_save))
    
    # Cập nhật trực tiếp bảng Bao_Cao_Tong_Hop nếu đang tồn tại
    cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='Bao_Cao_Tong_Hop'")
    if cursor.fetchone()[0] > 0:
        cursor.execute("SELECT [Quantity], [Putaway], [Backlog], [Receive Date] FROM Bao_Cao_Tong_Hop WHERE [Ship Order] = ?", (ship_order,))
        row = cursor.fetchone()
        if row:
            qty, put, backlog = int(row[0] or 0), int(row[1] or 0), int(row[2] or 0)
            curr_rec_date = str(row[3] or '').strip()
            if arrived == 'Done':
                new_status = 'Done'
            elif put == 0:
                new_status = 'Arriving' if arrived == 'Yes' else 'Pending'
            elif backlog <= 700 or put >= qty:
                new_status = 'Done'
            else:
                new_status = 'In Progress'

            new_rec_date = today_ddmmyyyy if arrived in ('Yes', 'Done') else (orig_date or curr_rec_date)

            conn.execute("""
                UPDATE Bao_Cao_Tong_Hop
                SET [Arrived] = ?, [Status] = ?, [Receive Date] = ?
                WHERE [Ship Order] = ?
            """, (arrived, new_status, new_rec_date, ship_order))
    conn.commit()
    conn.close()
    return True
    conn.commit()
    conn.close()
    return True


def compute_and_save_summary(db_path=DB_PATH):
    t0 = time.time()
    conn = sqlite3.connect(db_path)
    ensure_arrived_table(conn)
    
    cursor = conn.cursor()
    cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='data_mti'")
    if cursor.fetchone()[0] == 0:
        conn.close()
        return pd.DataFrame()
        
    df_data = pd.read_sql_query("SELECT ship_order, asn, sku, qty, date FROM data_mti WHERE ship_order IS NOT NULL AND trim(ship_order) != ''", conn)
    
    cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='Put MTI'")
    if cursor.fetchone()[0] == 0:
        conn.close()
        return pd.DataFrame()
        
    df_put = pd.read_sql_query("SELECT [ASN ID] as asn, [SKU ID] as sku, [Quantity] as qty, [Putaway Time] as put_time FROM [Put MTI]", conn)

    # Đọc trạng thái Arrived & ngày đã lưu của người dùng
    cursor.execute("SELECT ship_order, arrived, COALESCE(arrived_date, '') FROM _order_arrived_state")
    saved_arrived = {}
    saved_arrived_date = {}
    for r in cursor.fetchall():
        so_k = str(r[0]).strip()
        saved_arrived[so_k] = str(r[1]).strip()
        if r[2]:
            saved_arrived_date[so_k] = str(r[2]).strip()

    df_data['sku'] = df_data['sku'].astype(str).str.strip()
    df_data['qty'] = pd.to_numeric(df_data['qty'], errors='coerce').fillna(0).astype(int)
    df_data['asn'] = df_data['asn'].astype(str).str.strip()
    df_data['ship_order'] = df_data['ship_order'].astype(str).str.strip()
    df_data['date'] = df_data['date'].astype(str).str.strip()

    df_put['sku'] = df_put['sku'].astype(str).str.strip()
    df_put['asn'] = df_put['asn'].astype(str).str.strip()
    df_put['qty'] = pd.to_numeric(df_put['qty'], errors='coerce').fillna(0).astype(int)

    # Tính ngày Putaway thực tế theo từng ASN từ cột [Putaway Time]
    asn_put_dates = {}
    for _, r_p in df_put.iterrows():
        a = str(r_p['asn']).strip()
        pt = str(r_p.get('put_time', '') or '').strip()
        if a and pt and len(pt) >= 10:
            d_part = pt[:10]
            if '-' in d_part:
                p_parts = d_part.split('-')
                if len(p_parts) == 3 and len(p_parts[0]) == 4:
                    ddmmyyyy = f"{p_parts[2]}/{p_parts[1]}/{p_parts[0]}"
                else:
                    ddmmyyyy = d_part
            else:
                ddmmyyyy = d_part
            if a not in asn_put_dates or pt > asn_put_dates[a][0]:
                asn_put_dates[a] = (pt, ddmmyyyy)

    # 1. Map each order to its ASNs and SKUs
    orders = {}
    asn_to_orders = defaultdict(list)
    for so, group in df_data.groupby('ship_order'):
        asns = set(a for a in group['asn'] if a)
        skus = set(group['sku'])
        valid_asns = [a for a in group['asn'].astype(str).str.strip() if a]
        valid_dates = [d for d in group['date'].astype(str).str.strip() if d]
        orders[so] = {
            'ship_order': so,
            'primary_asn': valid_asns[0] if valid_asns else '',
            'receive_date': valid_dates[0] if valid_dates else '',
            'total_sku': len(group),
            'distinct_sku': len(skus),
            'quantity': int(group['qty'].sum()),
            'putaway': 0,
            'asns': asns,
            'skus': skus
        }
        for a in asns:
            asn_to_orders[a].append(so)

    # 2. Group Put MTI by ASN ID
    put_by_asn = defaultdict(list)
    for r in df_put.itertuples():
        put_by_asn[r.asn].append((r.sku, r.qty))

    # 3. For each Put ASN, find its owning ship order
    for put_asn, records in put_by_asn.items():
        total_put_qty = sum(q for s, q in records)
        
        matched_so = None
        if put_asn in asn_to_orders:
            cand = asn_to_orders[put_asn]
            if len(cand) == 1:
                matched_so = cand[0]
            else:
                best_cnt = -1
                for c in cand:
                    cnt = sum(1 for s, q in records if s in orders[c]['skus'])
                    if cnt > best_cnt:
                        best_cnt = cnt
                        matched_so = c
        else:
            # Put ASN is not directly in data_mti (e.g. pallet ASN MIVNDL2611050041)
            order_match_counts = defaultdict(int)
            for s, q in records:
                for so, o in orders.items():
                    if s in o['skus']:
                        order_match_counts[so] += 1
            if order_match_counts:
                best_so, best_skus = max(order_match_counts.items(), key=lambda x: x[1])
                distinct_put_skus = len(set(s for s, q in records))
                if distinct_put_skus > 0 and (best_skus / distinct_put_skus) >= 0.7:
                    matched_so = best_so
                    
        if matched_so:
            orders[matched_so]['putaway'] += total_put_qty
            if put_asn in asn_put_dates:
                if 'put_date' not in orders[matched_so] or asn_put_dates[put_asn][0] > orders[matched_so].get('_max_pt', ''):
                    orders[matched_so]['_max_pt'] = asn_put_dates[put_asn][0]
                    orders[matched_so]['put_date'] = asn_put_dates[put_asn][1]

    # Kiểm tra thêm ngày Putaway từ các ASN trực tiếp của đơn hàng
    for so, o in orders.items():
        if 'put_date' not in o:
            for a in o['asns']:
                if a in asn_put_dates:
                    if 'put_date' not in o or asn_put_dates[a][0] > o.get('_max_pt', ''):
                        o['_max_pt'] = asn_put_dates[a][0]
                        o['put_date'] = asn_put_dates[a][1]

    today_ddmmyyyy = time.strftime("%d/%m/%Y")
    today_key = time.strftime("%Y%m%d")

    rows = []
    for so, o in orders.items():
        qty = o['quantity']
        put = o['putaway']
        backlog = max(0, qty - put)
        
        if saved_arrived.get(so) == 'Done':
            arrived_val = 'Done'
            status = 'Done'
        elif put == 0:
            arrived_val = saved_arrived.get(so, 'No')
            status = 'Arriving' if arrived_val == 'Yes' else 'Pending'
            if so not in saved_arrived:
                conn.execute("""
                    INSERT OR IGNORE INTO _order_arrived_state (ship_order, arrived, arrived_date, updated_at)
                    VALUES (?, 'No', '', CURRENT_TIMESTAMP)
                """, (so,))
        else:
            # Bất kỳ đơn nào đã có số Put (Putaway > 0) -> Chắc chắn xe đã đến (Arrived = 'Yes')
            arrived_val = 'Yes'
            if saved_arrived.get(so) not in ('Yes', 'Done'):
                saved_arrived[so] = 'Yes'
            if backlog <= 700 or put >= qty:
                status = 'Done'
                arrived_val = 'Done'
                if saved_arrived.get(so) != 'Done':
                    saved_arrived[so] = 'Done'
                    conn.execute("""
                        INSERT INTO _order_arrived_state (ship_order, arrived, arrived_date, updated_at)
                        VALUES (?, 'Done', ?, CURRENT_TIMESTAMP)
                        ON CONFLICT(ship_order) DO UPDATE SET arrived='Done', updated_at=CURRENT_TIMESTAMP
                    """, (so, saved_arrived_date.get(so, '')))
            else:
                status = 'In Progress'

        # Ngày Receive Date: Mặc định lấy chuẩn từ sheet data_mti
        rec_date = o['receive_date']
        # Chỉ khi nào người dùng bấm "Yes" xe tới (hoặc có lưu ngày xe đến cụ thể) thì mới cập nhật ngày
        if saved_arrived.get(so) == 'Yes' and saved_arrived_date.get(so):
            rec_date = saved_arrived_date[so]
            
        # Nếu status là Done thì backlog về = 0 luôn và Putaway = Quantity
        if status == 'Done' or arrived_val == 'Done':
            put = qty
            backlog = 0
            # Khi đã hoàn thành (Done), Receive Date chuyển sang ngày hoàn thành thực tế (ngày putaway hoặc ngày bấm Done)
            actual_done_date = o.get('put_date') or saved_arrived_date.get(so)
            if actual_done_date:
                rec_date = actual_done_date
            
        rows.append({
            'Ship Order': so,
            'ASN': o['primary_asn'],
            'Receive Date': rec_date,
            'Total SKU': o['total_sku'],
            'Quantity': qty,
            'Putaway': put,
            'Backlog': backlog,
            'Status': status,
            'Arrived': arrived_val
        })

    df_res = pd.DataFrame(rows)
    # Thứ tự sắp xếp ưu tiên hiển thị: In Progress -> Arriving -> Pending -> Done
    status_order = {'In Progress': 1, 'Arriving': 2, 'Pending': 3, 'Done': 4}
    df_res['sort_status'] = df_res['Status'].map(status_order).fillna(5)
    df_res = df_res.sort_values(
        by=['sort_status', 'Receive Date', 'Quantity'],
        ascending=[True, False, False]
    ).drop(columns=['sort_status'])

    # Ghi lại bảng Bao_Cao_Tong_Hop (MTI) vào SQL
    df_res.to_sql('Bao_Cao_Tong_Hop (MTI)', conn, if_exists='replace', index=False)
    conn.execute("CREATE INDEX IF NOT EXISTS idx_bcth_mti_so ON [Bao_Cao_Tong_Hop (MTI)] ([Ship Order])")
    conn.execute("CREATE INDEX IF NOT EXISTS idx_bcth_mti_status ON [Bao_Cao_Tong_Hop (MTI)] ([Status])")
    # Xóa bảng cũ thừa nếu có để không bị trùng 2 bảng trong danh sách chọn
    try:
        conn.execute("DROP TABLE IF EXISTS [Bao_Cao_Tong_Hop]")
    except Exception:
        pass
    conn.commit()
    conn.close()

    dur = time.time() - t0
    print(f"Computed & saved {len(df_res)} orders to [Bao_Cao_Tong_Hop (MTI)] in {dur:.2f}s")
    return df_res

if __name__ == '__main__':
    compute_and_save_summary()
