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
        if not event.is_directory:  # ディレクトリでない場合のみ処理
            self.process(event.src_path)
    def on_modified(self, event):
        if not event.is_directory:  # ディレクトリでない場合のみ処理
            self.process(event.src_path)
    def process(self, path):
        if not (path.lower().endswith('.xls') or path.lower().endswith('.xlsx')):
            return
        print(f'[監視] 取り込み: {path}')

        # edited by Takashi Osako by 2025/05/23
        # ファイルロック待機（他プロセスが書き込み中の可能性）
        max_retries = 3
        for attempt in range(max_retries):
            try:
                df = pd.read_excel(path, header=2)  # 3行目がヘッダー
                break
            except (PermissionError, pd.errors.ExcelFileError) as e:
                if attempt < max_retries - 1:
                    print(f'[監視] 読み込み再試行 {attempt + 1}/{max_retries}: {e}')
                    time.sleep(2)
                else:
                    raise
        df = df.reindex(columns=COLUMNS)
        df = df.dropna(how='all')
        # 日付列を「YYYY-MM-DD」だけ、時刻列を「HH:MM」だけに整形
        #  日付列がない場合は空文字を入れる（NaT値を処理しないとエラーが発生する）
        if '手術日' in df.columns:
            df['手術日'] = pd.to_datetime(df['手術日'], errors='coerce').dt.strftime('%Y-%m-%d').fillna('')
        if '手術開始' in df.columns:
            df['手術開始'] = pd.to_datetime(df['手術開始'], errors='coerce').dt.strftime('%H:%M').fillna('')
        if '手術終了' in df.columns:
            df['手術終了'] = pd.to_datetime(df['手術終了'], errors='coerce').dt.strftime('%H:%M').fillna('')
        df = df.fillna('')
        conn = None
        try:
            conn = sqlite3.connect(DB_PATH)
            df.to_sql('records', conn, if_exists='replace', index=False)
            print('[監視] SQLite 更新完了')
        finally:
            if conn:
                conn.close()

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
