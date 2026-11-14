# -*- coding: utf-8 -*-
"""
TỔNG HỢP VÀ ĐỐI CHIẾU DỮ LIỆU RTS
- Nguồn: bảng [RTS] trong SQLite
- Tổng hợp theo Order Number, sum QTY (Allocated Quantity), đếm Total SKU, trạng thái Status
- Dữ liệu chi tiết từng dòng: Order Number, Supplier, SKU, SKU Name, QTY, Create Time, Ship Date, Status
- Cho phép chỉnh sửa Status trực tiếp trên web và lưu vào bảng _rts_status_overrides
"""

import os
import sys
import time
import sqlite3
from typing import Optional, Dict, Tuple
import pandas as pd

if hasattr(sys, 'stdout') and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

DB_PATH = r'C:\Users\srvn0514\Desktop\SQL\Data MTI.db'


def compute_and_save_summary_rts(db_path: Optional[str] = None) -> Tuple[pd.DataFrame, pd.DataFrame]:
    t0 = time.time()
    db_file = db_path or DB_PATH
    conn = sqlite3.connect(db_file)
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS _rts_status_overrides (
            order_number TEXT NOT NULL,
            sku TEXT NOT NULL,
            status TEXT NOT NULL,
            truck_info TEXT DEFAULT '',
            mail_status TEXT DEFAULT 'Chưa mail',
            return_address TEXT DEFAULT '',
            transfer_address TEXT DEFAULT '',
            updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            PRIMARY KEY (order_number, sku)
        )
    """)
    try:
        cursor.execute("ALTER TABLE _rts_status_overrides ADD COLUMN truck_info TEXT DEFAULT ''")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE _rts_status_overrides ADD COLUMN mail_status TEXT DEFAULT 'Chưa mail'")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE _rts_status_overrides ADD COLUMN return_address TEXT DEFAULT ''")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE _rts_status_overrides ADD COLUMN transfer_address TEXT DEFAULT ''")
    except Exception:
        pass
    conn.commit()

    cursor.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='RTS'")
    if cursor.fetchone()[0] == 0:
        conn.close()
        return pd.DataFrame(), pd.DataFrame()

    # Tự động cập nhật _rts_status_overrides nếu trong [RTS] đã có ngày Ship Date hoặc trạng thái Outbound/Shipped
    try:
        cursor.execute("""
            UPDATE _rts_status_overrides 
            SET status = (SELECT COALESCE(MAX([Status]), 'Outbound') FROM [RTS] WHERE [RTS].[Order Number] = _rts_status_overrides.order_number)
            WHERE order_number IN (
                SELECT [Order Number] FROM [RTS] 
                WHERE [Status] IN ('Outbound', 'Shipped') OR ([Ship Date] IS NOT NULL AND [Ship Date] != '' AND [Ship Date] != '-')
            )
        """)
        conn.commit()
    except Exception:
        pass

    df_rts = pd.read_sql_query("SELECT * FROM [RTS]", conn)
    if len(df_rts) == 0:
        conn.close()
        return pd.DataFrame(), pd.DataFrame()

    cursor.execute("PRAGMA table_info(_rts_status_overrides)")
    cols_override = [c[1] for c in cursor.fetchall()]
    if 'transfer_address' in cols_override:
        cursor.execute("SELECT order_number, sku, status, truck_info, mail_status, return_address, transfer_address FROM _rts_status_overrides")
    else:
        cursor.execute("SELECT order_number, sku, status, truck_info, mail_status, return_address FROM _rts_status_overrides")

    status_overrides = {}
    truck_overrides = {}
    mail_overrides = {}
    transfer_overrides = {}
    for r in cursor.fetchall():
        on_val = str(r[0] or '').strip()
        sku_val = str(r[1] or '').strip()
        status_overrides[(on_val, sku_val)] = str(r[2] or '').strip()
        if r[3] is not None and str(r[3]).strip():
            truck_overrides[on_val] = str(r[3]).strip()
        if r[4] is not None and str(r[4]).strip():
            mail_overrides[on_val] = str(r[4]).strip()
        
        tr_val = ''
        if len(r) > 6 and r[6] is not None and str(r[6]).strip():
            cand = str(r[6]).strip().upper()
            if cand in ['VNNL', 'VNA', 'VNVL', 'VNG', 'VNSL']:
                tr_val = cand
            elif str(r[6]).strip():
                tr_val = str(r[6]).strip()
        elif len(r) > 5 and r[5] is not None and str(r[5]).strip():
            cand = str(r[5]).strip().upper()
            if cand in ['VNNL', 'VNA', 'VNVL', 'VNG', 'VNSL']:
                tr_val = cand
        if tr_val:
            transfer_overrides[on_val] = tr_val

    items = []
    for _, r in df_rts.iterrows():
        on = str(r.get('Order Number', '')).strip()
        if not on or on in ('None', 'nan', ''):
            continue
        req_id = str(r.get('Request ID', '')).strip()
        if req_id in ('None', 'nan'):
            req_id = ''
        sku = str(r.get('SKU', '')).strip()
        s_name = str(r.get('SKU Name', '')).strip()
        supp = str(r.get('Supplier', '')).strip()
        qty_raw = r.get('Allocated Quantity', 0)
        try:
            qty = int(float(str(qty_raw).replace(',', '').strip()))
        except Exception:
            qty = 0
        ctime = str(r.get('Create Time', '')).strip()
        sdate = str(r.get('Ship Date', '')).strip()
        stat_sheet = str(r.get('Status', '')).strip() or 'Allocated'
        # Nếu bảng [RTS] đã cập nhật Ship Date hoặc Outbound/Shipped, bắt buộc ưu tiên trạng thái từ SQL/Sheet
        is_shipped = bool(sdate and sdate not in ('', 'None', 'nan', '-')) or stat_sheet.lower() in ('outbound', 'shipped')

        if is_shipped:
            stat = stat_sheet if stat_sheet.lower() in ('outbound', 'shipped') else 'Outbound'
        elif (on, sku) in status_overrides and status_overrides[(on, sku)]:
            stat = status_overrides[(on, sku)]
        elif (on, 'ALL') in status_overrides and status_overrides[(on, 'ALL')]:
            ov = status_overrides[(on, 'ALL')]
            if ov in ('Allocated', 'Pending Pick') and stat_sheet and stat_sheet != 'Allocated':
                stat = stat_sheet
            else:
                stat = ov
        else:
            stat = stat_sheet

        supp_name = str(r.get('Supplier Name Pic', r.get('Supplier Name', ''))).strip()
        tel = str(r.get('Telephone Number', r.get('Telephone', ''))).strip()
        orig_return_addr = str(r.get('Return Address', '')).strip()
        transfer_addr = transfer_overrides.get(on, '')
        ret_reason = str(r.get('Return Reason', '')).strip()
        ret_type = str(r.get('Return Type', '')).strip()
        deliv_method = str(r.get('Delivery Method', '')).strip()
        order_src = str(r.get('Order Source', '')).strip()
        whs_id = str(r.get('Whs ID', '')).strip()
        remark = str(r.get('Remark', '')).strip()
        num_pkg = str(r.get('Number of Packed Boxes', r.get('Number of Packages', ''))).strip()

        items.append({
            'Order Number': on,
            'Request ID': req_id,
            'Supplier': supp,
            'Supplier Name': supp_name,
            'Telephone': tel,
            'Return Address': orig_return_addr,
            'Địa chỉ trả hàng': orig_return_addr,
            'Địa chỉ chuyển': transfer_addr,
            'transfer_address': transfer_addr,
            'Return Reason': ret_reason,
            'Return Type': ret_type,
            'Delivery Method': deliv_method,
            'Order Source': order_src,
            'Whs ID': whs_id,
            'Remark': remark,
            'Number of Packages': num_pkg,
            'SKU': sku,
            'SKU Name': s_name,
            'QTY': qty,
            'Create Time': ctime,
            'Ship Date': sdate,
            'Status': stat,
            'Thông tin xe': truck_overrides.get(on, ''),
            'Mail': mail_overrides.get(on, 'Chưa mail')
        })

    df_items = pd.DataFrame(items)

    # Tổng hợp theo Order Number
    orders = []
    if len(df_items) > 0:
        for on, grp in df_items.groupby('Order Number'):
            first_row = grp.iloc[0]
            req_id = str(first_row.get('Request ID', '')).strip()
            supp = first_row.get('Supplier', '')
            supp_name = first_row.get('Supplier Name', '')
            tel = first_row.get('Telephone', '')
            ret_addr = first_row.get('Return Address', '')
            tr_addr = first_row.get('Địa chỉ chuyển', '')
            ret_reason = first_row.get('Return Reason', '')
            ret_type = first_row.get('Return Type', '')
            deliv_method = first_row.get('Delivery Method', '')
            order_src = first_row.get('Order Source', '')
            whs_id = first_row.get('Whs ID', '')
            remark = first_row.get('Remark', '')
            num_pkg = first_row.get('Number of Packages', '')
            total_sku = grp['SKU'].nunique()
            total_qty = int(grp['QTY'].sum())
            ctime = first_row.get('Create Time', '')
            sdate = first_row.get('Ship Date', '')
            statuses = list(grp['Status'].unique())
            if len(statuses) == 1:
                status_str = statuses[0]
            else:
                status_str = ", ".join(statuses)

            orders.append({
                'Order Number': on,
                'Request ID': req_id,
                'Supplier': supp,
                'Supplier Name': supp_name,
                'Telephone': tel,
                'Return Address': ret_addr,
                'Địa chỉ trả hàng': ret_addr,
                'Địa chỉ chuyển': tr_addr,
                'transfer_address': tr_addr,
                'Return Reason': ret_reason,
                'Return Type': ret_type,
                'Delivery Method': deliv_method,
                'Order Source': order_src,
                'Whs ID': whs_id,
                'Remark': remark,
                'Number of Packages': num_pkg,
                'Total SKU': total_sku,
                'Total QTY': total_qty,
                'Create Time': ctime,
                'Ship Date': sdate,
                'Status': status_str,
                'Thông tin xe': truck_overrides.get(on, ''),
                'Mail': mail_overrides.get(on, 'Chưa mail')
            })

    # Đảm bảo bảng [RTS] có đủ 3 cột: Địa chỉ chuyển, Thông tin xe, Mail
    try:
        cursor.execute("ALTER TABLE [RTS] ADD COLUMN [Địa chỉ chuyển] TEXT DEFAULT ''")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE [RTS] ADD COLUMN [Thông tin xe] TEXT DEFAULT ''")
    except Exception:
        pass
    try:
        cursor.execute("ALTER TABLE [RTS] ADD COLUMN [Mail] TEXT DEFAULT 'Chưa mail'")
    except Exception:
        pass

    all_overridden_orders = set(list(transfer_overrides.keys()) + list(truck_overrides.keys()) + list(mail_overrides.keys()))
    for on_key in all_overridden_orders:
        tr_val = transfer_overrides.get(on_key, '')
        tk_val = truck_overrides.get(on_key, '')
        ml_val = mail_overrides.get(on_key, 'Chưa mail')
        cursor.execute("""
            UPDATE [RTS] SET 
                [Địa chỉ chuyển] = ?,
                [Thông tin xe] = ?,
                [Mail] = ?
            WHERE [Order Number] = ?
        """, (tr_val, tk_val, ml_val, on_key))

    df_orders = pd.DataFrame(orders)

    # Ghi vào bảng tổng hợp SQL [Bao_Cao_Tong_Hop (RTS)]
    df_orders.to_sql('Bao_Cao_Tong_Hop (RTS)', conn, if_exists='replace', index=False)
    conn.commit()
    conn.close()

    dur = time.time() - t0
    print(f"Computed & saved {len(df_orders)} orders ({len(df_items)} items) to [Bao_Cao_Tong_Hop (RTS)] in {dur:.2f}s")
    return df_orders, df_items


def set_rts_status(order_number: str, sku: str, new_status: str, db_path: Optional[str] = None):
    """Cập nhật Status cho dòng RTS (lưu vào _rts_status_overrides)."""
    db_file = db_path or DB_PATH
    conn = sqlite3.connect(db_file)
    on_clean = str(order_number).strip()
    sku_clean = str(sku).strip() or 'ALL'

    with conn:
        conn.execute("""
            INSERT INTO _rts_status_overrides (order_number, sku, status, updated_at)
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(order_number, sku) DO UPDATE SET 
                status = excluded.status,
                updated_at = CURRENT_TIMESTAMP
        """, (on_clean, sku_clean, new_status))
        if sku_clean == 'ALL':
            conn.execute("UPDATE _rts_status_overrides SET status = ?, updated_at = CURRENT_TIMESTAMP WHERE order_number = ?", (new_status, on_clean))
            try:
                conn.execute("UPDATE [RTS] SET [Status] = ? WHERE [Order Number] = ?", (new_status, on_clean))
            except Exception:
                pass
    conn.close()
    compute_and_save_summary_rts(db_file)


def set_rts_truck_info(order_number: str, truck_info: str, db_path: Optional[str] = None):
    """Cập nhật Thông tin xe cho đơn RTS (lưu vào _rts_status_overrides và đồng bộ bảng RTS)."""
    db_file = db_path or DB_PATH
    conn = sqlite3.connect(db_file)
    on_clean = str(order_number).strip()
    t_clean = str(truck_info).strip()

    with conn:
        conn.execute("""
            INSERT INTO _rts_status_overrides (order_number, sku, status, truck_info, mail_status, updated_at)
            VALUES (?, 'ALL', (SELECT COALESCE(MAX([Status]), '') FROM [RTS] WHERE [Order Number] = ?), ?, 'Chưa mail', CURRENT_TIMESTAMP)
            ON CONFLICT(order_number, sku) DO UPDATE SET 
                truck_info = excluded.truck_info,
                updated_at = CURRENT_TIMESTAMP
        """, (on_clean, on_clean, t_clean))
        conn.execute("UPDATE _rts_status_overrides SET truck_info = ?, updated_at = CURRENT_TIMESTAMP WHERE order_number = ?", (t_clean, on_clean))
        try:
            conn.execute("ALTER TABLE [RTS] ADD COLUMN [Thông tin xe] TEXT DEFAULT ''")
        except Exception:
            pass
        try:
            conn.execute("UPDATE [RTS] SET [Thông tin xe] = ? WHERE [Order Number] = ?", (t_clean, on_clean))
        except Exception:
            pass
    conn.close()
    compute_and_save_summary_rts(db_file)


def set_rts_mail_status(order_number: str, mail_status: str, db_path: Optional[str] = None):
    """Cập nhật trạng thái Mail (Mail / Chưa mail) cho đơn RTS (lưu vào _rts_status_overrides và đồng bộ bảng RTS)."""
    db_file = db_path or DB_PATH
    conn = sqlite3.connect(db_file)
    on_clean = str(order_number).strip()
    m_clean = str(mail_status).strip() or 'Chưa mail'

    with conn:
        conn.execute("""
            INSERT INTO _rts_status_overrides (order_number, sku, status, truck_info, mail_status, updated_at)
            VALUES (?, 'ALL', (SELECT COALESCE(MAX([Status]), '') FROM [RTS] WHERE [Order Number] = ?), '', ?, CURRENT_TIMESTAMP)
            ON CONFLICT(order_number, sku) DO UPDATE SET 
                mail_status = excluded.mail_status,
                updated_at = CURRENT_TIMESTAMP
        """, (on_clean, on_clean, m_clean))
        conn.execute("UPDATE _rts_status_overrides SET mail_status = ?, updated_at = CURRENT_TIMESTAMP WHERE order_number = ?", (m_clean, on_clean))
        try:
            conn.execute("ALTER TABLE [RTS] ADD COLUMN [Mail] TEXT DEFAULT 'Chưa mail'")
        except Exception:
            pass
        try:
            conn.execute("UPDATE [RTS] SET [Mail] = ? WHERE [Order Number] = ?", (m_clean, on_clean))
        except Exception:
            pass
    conn.close()
    compute_and_save_summary_rts(db_file)


def set_rts_transfer_address(order_number: str, transfer_address: str, db_path: Optional[str] = None):
    """Cập nhật Địa chỉ chuyển (VNNL, VNVL, VNA, VNSL, VNG hoặc điền tay) cho đơn RTS."""
    db_file = db_path or DB_PATH
    conn = sqlite3.connect(db_file)
    on_clean = str(order_number).strip()
    addr_clean = str(transfer_address).strip()

    with conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS _rts_status_overrides (
                order_number TEXT NOT NULL,
                sku TEXT NOT NULL,
                status TEXT NOT NULL,
                truck_info TEXT DEFAULT '',
                mail_status TEXT DEFAULT 'Chưa mail',
                return_address TEXT DEFAULT '',
                transfer_address TEXT DEFAULT '',
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                PRIMARY KEY (order_number, sku)
            )
        """)
        try:
            conn.execute("ALTER TABLE _rts_status_overrides ADD COLUMN transfer_address TEXT DEFAULT ''")
        except Exception:
            pass

        conn.execute("""
            INSERT INTO _rts_status_overrides (order_number, sku, status, truck_info, mail_status, return_address, transfer_address, updated_at)
            VALUES (?, 'ALL', (SELECT COALESCE(MAX([Status]), '') FROM [RTS] WHERE [Order Number] = ?), '', 'Chưa mail', ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(order_number, sku) DO UPDATE SET 
                return_address = excluded.return_address,
                transfer_address = excluded.transfer_address,
                updated_at = CURRENT_TIMESTAMP
        """, (on_clean, on_clean, addr_clean, addr_clean))
        conn.execute("UPDATE _rts_status_overrides SET transfer_address = ?, return_address = ?, updated_at = CURRENT_TIMESTAMP WHERE order_number = ?", (addr_clean, addr_clean, on_clean))
        try:
            conn.execute("ALTER TABLE [RTS] ADD COLUMN [Địa chỉ chuyển] TEXT DEFAULT ''")
        except Exception:
            pass
        try:
            conn.execute("UPDATE [RTS] SET [Địa chỉ chuyển] = ? WHERE [Order Number] = ?", (addr_clean, on_clean))
        except Exception:
            pass
    conn.close()
    compute_and_save_summary_rts(db_file)


def set_rts_return_address(order_number: str, return_address: str, db_path: Optional[str] = None):
    """Alias cho set_rts_transfer_address để tương thích ngược."""
    set_rts_transfer_address(order_number, return_address, db_path)


if __name__ == '__main__':
    o, it = compute_and_save_summary_rts()
    print("Orders:", len(o), "Items:", len(it))
