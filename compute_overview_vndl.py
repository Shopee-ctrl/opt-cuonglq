# -*- coding: utf-8 -*-
"""
TỔNG HỢP VÀ PHÂN TÍCH CAPACITY KHO VNDL (OVER VIEW VNDL)
- Nguồn dữ liệu:
  + Bảng [Cell] trong SQLite: đếm số vị trí theo [zone_id] và [cell_occupied]
    * Location Qty: số dòng có cell_occupied IN ('Y', 'N')
    * Used Qty: số dòng có cell_occupied = 'Y'
    * Empty Qty: số dòng có cell_occupied = 'N'
    * Max_capacity(m³): dung tích tối đa theo loại kệ (A1: * 0.045144, XS: * 0.012, B/C: * 1.344, BT: * 2.064)
  + Bảng [Tổng hợp ONH (VNDL)] (hoặc [ONH]):
    * Volume Used(m³): tổng CBM theo Zone ID
    * Qty Onhands: tổng On-hand Qty theo Zone ID
    * Empty_Capacity(m³) = Max_capacity - Volume Used
    * AVG Items available = Qty Onhands / Used Qty
- Kết quả:
  + Ghi bảng [Over view VNDL] vào SQLite
  + Cung cấp dữ liệu trực quan cho HTML Dashboard (KPI, Biểu đồ ngang, Biểu đồ dung tích)
"""

import os
import sys
import time
import sqlite3
from typing import Optional, Dict, Any, List
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

DB_PATH = r'C:\Users\srvn0514\Desktop\SQL\Data MTI.db'


def compute_and_save_overview_vndl(db_path: Optional[str] = None) -> Dict[str, Any]:
    """
    Tính toán báo cáo CAPACITY VNDL và lưu vào SQLite bảng [Over view VNDL].
    Trả về dictionary cấu trúc dữ liệu đầy đủ cho Dashboard & Charts.
    """
    t0 = time.time()
    db_file = db_path or DB_PATH
    conn = sqlite3.connect(db_file)
    cur = conn.cursor()

    # Kiểm tra bảng Cell
    cur.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='Cell'")
    if cur.fetchone()[0] == 0:
        conn.close()
        return {'items': [], 'total': {}, 'chart_data': []}

    # Đọc Cell
    df_cell = pd.read_sql_query("SELECT zone_id, cell_occupied, [max_capacity(cu.cm.)] FROM [Cell]", conn)

    # Đọc Tổng hợp ONH (hoặc ONH)
    onh_table = 'Tổng hợp ONH (VNDL)'
    cur.execute(f"SELECT count(*) FROM sqlite_master WHERE type='table' AND name='{onh_table}'")
    if cur.fetchone()[0] == 0:
        onh_table = 'ONH'
        cur.execute("SELECT count(*) FROM sqlite_master WHERE type='table' AND name='ONH'")
        if cur.fetchone()[0] == 0:
            df_onh = pd.DataFrame()
        else:
            df_onh = pd.read_sql_query("SELECT * FROM [ONH]", conn)
    else:
        df_onh = pd.read_sql_query(f"SELECT * FROM [{onh_table}]", conn)

    onh_by_zone = {}
    if not df_onh.empty:
        # Chuẩn hóa cột Zone id
        zone_col = None
        for c in df_onh.columns:
            if c.lower() in ('zone id', 'zone_id', 'zone'):
                zone_col = c
                break
        
        cbm_col = 'CBM' if 'CBM' in df_onh.columns else None
        qty_col = None
        for c in df_onh.columns:
            if c.lower() in ('on-hand qty', 'on_hand_qty', 'on hand qty', 'qty'):
                qty_col = c
                break

        if zone_col:
            df_onh['cbm_num'] = pd.to_numeric(df_onh[cbm_col], errors='coerce').fillna(0) if cbm_col else 0
            df_onh['qty_num'] = pd.to_numeric(df_onh[qty_col], errors='coerce').fillna(0) if qty_col else 0
            df_onh['vol_item_cbm'] = df_onh['cbm_num'] * df_onh['qty_num']

            grp = df_onh.groupby(zone_col).agg({
                'cbm_num': 'sum',
                'vol_item_cbm': 'sum',
                'qty_num': 'sum'
            })
            onh_by_zone = grp.to_dict('index')

    # Tự động nhận diện zone Rack Storage là AT1 (mới) hoặc AT (cũ)
    rack_storage_zone = 'AT1' if 'AT1' in df_cell['zone_id'].values else ('AT' if 'AT' in df_cell['zone_id'].values else 'AT1')

    # Cấu hình 6 Zone chính chuẩn theo kho VNDL
    core_zone_configs = [
        {'type': 'PICKING RACK', 'zone': 'A1', 'cap_coef': 0.045144},
        {'type': 'RACK STORAGE', 'zone': rack_storage_zone, 'cap_coef': 0.045144},
        {'type': 'PICKING RACK', 'zone': 'XS', 'cap_coef': 0.012000},
        {'type': 'PICKING PALLET', 'zone': 'B', 'cap_coef': 1.344000},
        {'type': 'PALLET STORAGE', 'zone': 'BT', 'cap_coef': 2.064000},
        {'type': 'PICKING PALLET', 'zone': 'C', 'cap_coef': 1.344000},
    ]

    items = []
    tot_empty = 0
    tot_used = 0
    tot_loc = 0
    tot_max_cap = 0
    tot_vol_used = 0
    tot_onhand = 0

    for cfg in core_zone_configs:
        zid = cfg['zone']
        ztype = cfg['type']
        coef = cfg['cap_coef']

        z_cells = df_cell[df_cell['zone_id'] == zid]
        loc_qty = int(len(z_cells[z_cells['cell_occupied'].isin(['Y', 'N'])]))
        used_qty = int(len(z_cells[z_cells['cell_occupied'] == 'Y']))
        empty_qty = int(len(z_cells[z_cells['cell_occupied'] == 'N']))

        # Max_capacity làm tròn số nguyên (int)
        max_cap = int(round(loc_qty * coef, 0))

        onh_info = onh_by_zone.get(zid, {'cbm_num': 0, 'vol_item_cbm': 0, 'qty_num': 0})
        # Volume Used(m3) = tổng CBM thực tế (nhân số lượng tồn kho) làm tròn số nguyên
        vol_used_calc = onh_info.get('vol_item_cbm', 0)
        if vol_used_calc == 0 and onh_info.get('cbm_num', 0) > 0:
            vol_used_calc = onh_info.get('cbm_num', 0)
        vol_used = int(round(float(vol_used_calc), 0))
        qty_onhand = int(onh_info.get('qty_num', 0))

        # Empty_Capacity = Max_capacity - Volume Used làm tròn số nguyên
        empty_cap = int(round(max_cap - vol_used, 0))
        avg_items = round(qty_onhand / used_qty, 1) if used_qty > 0 else 0
        pct_used = round((used_qty / loc_qty * 100), 1) if loc_qty > 0 else 0
        pct_cbm = round((vol_used / max_cap * 100), 1) if max_cap > 0 else 0

        tot_empty += empty_qty
        tot_used += used_qty
        tot_loc += loc_qty
        tot_max_cap += max_cap
        tot_vol_used += vol_used
        tot_onhand += qty_onhand

        # Đặt thứ tự cột: Max_capacity(m³) -> Volume Used(m³) -> Empty_Capacity(m³) -> Qty Onhands
        items.append({
            'Type': ztype,
            'Zone': zid,
            'Empty Qty': empty_qty,
            'Used Qty': used_qty,
            'Location Qty': loc_qty,
            'Max_capacity(m³)': max_cap,
            'Volume Used(m³)': vol_used,
            'Empty_Capacity(m³)': empty_cap,
            'Qty Onhands': qty_onhand,
            'AVG Items available': avg_items,
            'Occupancy_Rate': pct_used,
            'CBM_Rate': pct_cbm
        })

    tot_vol_used = int(round(tot_vol_used, 0))
    tot_empty_cap = int(round(tot_max_cap - tot_vol_used, 0))
    tot_avg_items = round(tot_onhand / tot_used, 1) if tot_used > 0 else 0
    tot_occupancy = round(tot_used / tot_loc * 100, 1) if tot_loc > 0 else 0
    tot_cbm_rate = round(tot_vol_used / tot_max_cap * 100, 1) if tot_max_cap > 0 else 0

    total_row = {
        'Type': 'TOTAL',
        'Zone': 'ALL',
        'Empty Qty': tot_empty,
        'Used Qty': tot_used,
        'Location Qty': tot_loc,
        'Max_capacity(m³)': tot_max_cap,
        'Volume Used(m³)': tot_vol_used,
        'Empty_Capacity(m³)': tot_empty_cap,
        'Qty Onhands': tot_onhand,
        'AVG Items available': tot_avg_items,
        'Occupancy_Rate': tot_occupancy,
        'CBM_Rate': tot_cbm_rate
    }

    # Bảng dữ liệu SQLite
    df_save = pd.DataFrame(items + [total_row])
    with conn:
        conn.execute("DROP TABLE IF EXISTS [Over view VNDL]")
    df_save.to_sql('Over view VNDL', conn, if_exists='replace', index=False)
    conn.commit()
    conn.close()

    dur = round(time.time() - t0, 2)
    print(f"[Overview VNDL] Đã lưu bảng [Over view VNDL] ({len(items)} zones + 1 TOTAL) trong {dur}s")

    return {
        'items': items,
        'total': total_row,
        'chart_data': [
            {
                'zone': r['Zone'],
                'empty': r['Empty Qty'],
                'used': r['Used Qty'],
                'total': r['Location Qty'],
                'empty_pct': round((r['Empty Qty'] / (r['Location Qty'] or 1)) * 100, 1),
                'used_pct': round((r['Used Qty'] / (r['Location Qty'] or 1)) * 100, 1)
            }
            for r in items
        ],
        'updated_at': time.strftime("%H:%M:%S %d/%m/%Y")
    }


if __name__ == '__main__':
    data = compute_and_save_overview_vndl()
    print("Items:", len(data['items']))
    print("Total:", data['total'])
