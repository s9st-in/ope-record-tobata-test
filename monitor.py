import os
import time
import threading
import sqlite3
import pandas as pd
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

BASE_DIR = os.getcwd()
WATCH_DIR = os.path.join(BASE_DIR, 'watch_folder')
DB_PATH   = os.path.join(BASE_DIR, 'data.db')

COLUMNS = [
    '手術日','手術開始','手術終了','病棟','オーダ区分',
    '患者氏名','患者性別','患者年齢','疾患名','術式',
    '麻酔医','主治医','執刀医','助手','器械出し','外回り'
]

class ExcelHandler(FileSystemEventHandler):
    def on_created(self, event):
        self.process(event.src_path)
    def on_modified(self, event):
        self.process(event.src_path)
    def process(self, path):
        if not (path.lower().endswith('.xls') or path.lower().endswith('.xlsx')):
            return
        print(f'[監視] 取り込み: {path}')
        df = pd.read_excel(path, header=2)  # 3行目がヘッダー
        df = df.reindex(columns=COLUMNS)
        df = df.dropna(how='all')
        # 日付列を「YYYY-MM-DD」だけ、時刻列を「HH:MM」だけに整形
        if '手術日' in df.columns:
            df['手術日'] = pd.to_datetime(df['手術日'], errors='coerce').dt.strftime('%Y-%m-%d')
        if '手術開始' in df.columns:
            df['手術開始'] = pd.to_datetime(df['手術開始'], errors='coerce').dt.strftime('%H:%M')
        if '手術終了' in df.columns:
            df['手術終了'] = pd.to_datetime(df['手術終了'], errors='coerce').dt.strftime('%H:%M')
        df = df.fillna('')
        conn = sqlite3.connect(DB_PATH)
        df.to_sql('records', conn, if_exists='replace', index=False)
        conn.close()
        print('[監視] SQLite 更新完了')

def start_watcher():
    os.makedirs(WATCH_DIR, exist_ok=True)
    observer = Observer()
    handler = ExcelHandler()
    observer.schedule(handler, WATCH_DIR, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1800)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
