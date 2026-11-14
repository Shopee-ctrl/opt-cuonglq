# -*- coding: utf-8 -*-
import os, shutil

files_to_sync = [
    'compute_summary_onh.py',
    'compute_overview_vndl.py',
    'compute_summary_rts.py',
    'compute_summary_ib.py',
    'dashboard.py'
]

target_dirs = [
    r'C:\Users\srvn0514\Desktop\Inbound_Backlog_Dashboard',
    r'C:\Users\srvn0514\Desktop\SQL',
    r'C:\Users\srvn0514\Desktop\Tool_Dong_Bo_SQL'
]

src_dir = r'c:\Users\srvn0514\Desktop\mới'

for fn in files_to_sync:
    src_file = os.path.join(src_dir, fn)
    if os.path.exists(src_file):
        for td in target_dirs:
            if os.path.exists(td):
                dst_file = os.path.join(td, fn)
                shutil.copy2(src_file, dst_file)
                print(f"Copied {fn} -> {td}")
print("Done syncing files to all directories!")
