# -*- coding: utf-8 -*-
"""
Sync Data Plan MTI from Google Sheet to SQLite table [Data Plan MTI]
"""
import gspread
import pandas as pd
import sqlite3
import time
import os
import sys

def parse_plan_qty(val_raw):
    s = str(val_raw or '').strip().replace(' ', '')
    if not s or s in ('-', '  -  ', 'nan', 'None'):
        return 0
    try:
        if ',' in s:
            parts = s.split(',')
            if len(parts) == 2 and len(parts[1]) != 3:
                # Số thập phân theo nghìn (VD: 21,6 -> 21.6 * 1000 = 21,600)
                val = float(parts[0] + '.' + parts[1]) * 1000
                return int(round(val))
            else:
                # Dấu phẩy ngăn cách hàng nghìn (VD: 8,543 -> 8543, 54,489 -> 54489)
                return int(round(float(s.replace(',', ''))))
        elif '.' in s:
            parts = s.split('.')
            if len(parts) == 2 and len(parts[1]) != 3:
                val = float(s) * 1000
                return int(round(val))
            else:
                return int(round(float(s.replace('.', ''))))
        else:
            val = float(s)
            # Nếu người dùng nhập theo nghìn (k) (VD: 20 -> 20.000, 18 -> 18.000, 1 -> 1.000, 39 -> 39.000)
            if 0 < val < 500:
                return int(round(val * 1000))
            return int(round(val))
    except Exception:
        return 0

def sync_data_plan_mti(client, db_path, emit=print):
    emit("\n--- [PHẦN 5] ĐỒNG BỘ 'Data Plan MTI' ➔ table 'Data Plan MTI' ---")
    t0 = time.time()
    try:
        SPREADSHEET_ID_PLAN = '1LlMGkx2FM9lmvREUqsm35QNNjEXpj4MifbZwPZ8rMjw'
        sh = client.open_by_key(SPREADSHEET_ID_PLAN)
        ws = sh.worksheet('Data Plan MTI')
        vals = ws.get_all_values()
        if not vals or len(vals) < 5:
            emit("    [LƯU Ý]: Sheet 'Data Plan MTI' không đủ số dòng.")
            return

        date_row = vals[3]  # Row 4 (0-indexed: 3)
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
                yyyy = '2026'
                date_col_map[c_idx] = f"{day_str}/{mm}/{yyyy}"

        parsed_records = []
        in_target_section = False

        for r_idx, r in enumerate(vals):
            col_a = r[0].strip() if len(r) > 0 else ''
            col_d = r[3].strip() if len(r) > 3 else ''

            if 'item_mti_fbs mp_vnd*' in col_a.lower():
                in_target_section = True

            if in_target_section:
                # Dòng bắt đầu bằng VN... đến hết VN.. qua chữ khác là dừng
                clean_wh = col_d.replace('*', '').strip().upper()
                if clean_wh.startswith('VN'):
                    for c_idx, dt_str in date_col_map.items():
                        if c_idx < len(r):
                            val_raw = r[c_idx].strip()
                            qty = parse_plan_qty(val_raw)
                            if qty > 0:
                                parsed_records.append({
                                    'date': dt_str,
                                    'warehouse': clean_wh,
                                    'raw_label': col_a,
                                    'qty': qty
                                })
                elif in_target_section and r_idx > 45:
                    # Gặp dòng không bắt đầu bằng VN -> Dừng lại
                    break

        df_plan = pd.DataFrame(parsed_records)
        conn = sqlite3.connect(db_path)
        df_plan.to_sql('Data Plan MTI', conn, if_exists='replace', index=False)
        conn.commit()
        conn.close()

        dur = round(time.time() - t0, 2)
        emit(f"    [THÀNH CÔNG] Đã lưu {len(df_plan):,} bản ghi vào bảng [Data Plan MTI] ({dur}s)")
        return True
    except Exception as e:
        emit(f"    [LỖI LINK 5 - Data Plan MTI]: {e}")
        return False

if __name__ == '__main__':
    creds_path = r'c:\Users\srvn0514\Desktop\mới\credentials.json'
    db_path = r'C:\Users\srvn0514\Desktop\SQL\Data MTI.db'
    client = gspread.service_account(filename=creds_path)
    sync_data_plan_mti(client, db_path)
