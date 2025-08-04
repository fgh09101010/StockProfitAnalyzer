from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import pandas as pd
from datetime import datetime
import os

# 1. 認證 需要有 client_secrets.json 檔案
#    這個檔案可以從 Google Cloud Console 下載，並放在與此腳本同一目錄下。
#    如果沒有這個檔案，請參考 PyDrive 的官方文檔來建立。
gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)

# 本地 Excel 路徑與雲端檔案名稱（固定檔名）
local_file = '未實現損益試算.xlsx'
remote_title = '未實現損益試算.xlsx'

# 取得本地檔案建立日期
mtime = os.path.getmtime(local_file)
created_date_str = datetime.fromtimestamp(mtime).strftime('%Y-%m-%d %H:%M:%S')
print(f"本地上次修改日期: {created_date_str}")

# 讀本地檔案，加上「資料日期」欄位
df = pd.read_excel(local_file)
df['資料日期'] = created_date_str

# 查雲端有沒有同名檔案
file_list = drive.ListFile({'q': f"title='{remote_title}' and trashed=false"}).GetList()

if file_list:
    # 覆蓋更新
    file_drive = file_list[0]
    df.to_excel('temp_with_date.xlsx', index=False)
    file_drive.SetContentFile('temp_with_date.xlsx')
    file_drive.Upload()
    print(f"已更新雲端檔案: {remote_title}")
else:
    # 第一次上傳
    df.to_excel('temp_with_date.xlsx', index=False)
    file_drive = drive.CreateFile({'title': remote_title})
    file_drive.SetContentFile('temp_with_date.xlsx')
    file_drive.Upload()
    print(f"首次上傳檔案到雲端: {remote_title}")
