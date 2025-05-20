import os
import time
import threading
import sqlite3
import pandas as pd
from datetime import datetime, timedelta
from flask import Flask, jsonify, render_template, request
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

def clean_row(row):
    return [str(v).replace('\u3000', ' ').replace('\n', '').strip() if pd.notnull(v) else '' for v in row]

def read_excel(file):
    try:
        df = pd.read_excel(file, header=2, dtype=str)
        df = df[COLUMNS]
        df = df.fillna("")
        df = df.apply(clean_row, axis=1, result_type="expand")
        df.columns = COLUMNS
        return df
    except Exception as e:
        print(f'[監視] 取り込み失敗: {file} {e}')
        return pd.DataFrame(columns=COLUMNS)

def load_db():
    if not os.path.exists(DB_PATH):
        return pd.DataFrame(columns=COLUMNS)
    conn = sqlite3.connect(DB_PATH)
    df = pd.read_sql('SELECT * FROM records', conn)
    conn.close()
    return df

def save_db(df):
    conn = sqlite3.connect(DB_PATH)
    df.to_sql('records', conn, if_exists='replace', index=False)
    conn.close()

def update_all_from_folder():
    files = [os.path.join(WATCH_DIR, f) for f in os.listdir(WATCH_DIR) if f.lower().endswith(('.xls','.xlsx'))]
    dfs = []
    for file in files:
        df = read_excel(file)
        if not df.empty:
            dfs.append(df)
    if dfs:
        all_df = pd.concat(dfs, ignore_index=True)
        all_df = all_df.drop_duplicates(subset=COLUMNS, keep="last")
        save_db(all_df)
    else:
        save_db(pd.DataFrame(columns=COLUMNS))

class ExcelHandler(FileSystemEventHandler):
    def on_created(self, event): self.process()
    def on_modified(self, event): self.process()
    def process(self):
        print('[監視] 全Excel再集計')
        update_all_from_folder()
        print('[監視] SQLite 更新完了')

def start_watcher():
    os.makedirs(WATCH_DIR, exist_ok=True)
    observer = Observer()
    handler = ExcelHandler()
    observer.schedule(handler, WATCH_DIR, recursive=False)
    observer.start()
    try:
        while True:
            time.sleep(1800)  # 30分毎
    except KeyboardInterrupt:
        observer.stop()
    observer.join()

app = Flask(__name__)

def format_date(date_str):
    if not date_str: return ""
    try:
        dt = None
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
            try:
                dt = datetime.strptime(date_str, fmt)
                break
            except: continue
        if not dt:
            dt = pd.to_datetime(date_str)
        wdays = '月火水木金土日'
        return f"{dt.month}月{dt.day}日({wdays[dt.weekday()]})"
    except:
        return date_str

def format_time(tstr):
    if not tstr or tstr in ['nan', 'None', None, ""]:
        return ""
    try:
        if '.' in tstr:
            tstr = tstr.split('.')[0]
        parts = tstr.split(":")
        if len(parts) >= 2:
            return f"{int(parts[0]):02d}:{int(parts[1]):02d}"
        return tstr
    except:
        return tstr

def to_dt(date_str):
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try: return datetime.strptime(str(date_str), fmt)
        except: continue
    try: return pd.to_datetime(date_str)
    except: return None

def to_time(tstr):
    if not tstr or tstr in ['nan', 'None', None, ""]:
        return datetime.strptime("23:59", "%H:%M").time()  # 時刻無しは最後尾へ
    try:
        if '.' in tstr:
            tstr = tstr.split('.')[0]
        if tstr.count(':') == 2:
            h, m, s = tstr.split(':')
            return datetime.strptime(f"{int(h):02d}:{int(m):02d}", "%H:%M").time()
        elif tstr.count(':') == 1:
            h, m = tstr.split(':')
            return datetime.strptime(f"{int(h):02d}:{int(m):02d}", "%H:%M").time()
        return datetime.strptime(tstr, "%H:%M").time()
    except:
        return datetime.strptime("23:59", "%H:%M").time()

def query_records():
    df = load_db()
    if df.empty:
        return []
    today = datetime.now()
    week_ago = today - timedelta(days=10)
    df['手術日_dt'] = df['手術日'].apply(to_dt)
    df = df[df['手術日_dt'].notnull()]
    df = df[df['手術日_dt'] >= week_ago]
    # ソートロジック
    df['手術開始_time'] = df['手術開始'].apply(to_time)
    df = df.sort_values(by=['手術日_dt','手術開始_time'], ascending=[False,True])
    result = []
    for _, row in df.iterrows():
        rec = {}
        for col in COLUMNS:
            val = row[col]
            if col == '手術日':
                val = format_date(val)
            elif col in ['手術開始','手術終了']:
                val = format_time(val)
            if val in [None, "nan", "NaT"]: val = ""
            rec[col] = val
        result.append(rec)
    return result

@app.route('/api/records')
def api_records():
    return jsonify(query_records())

@app.route('/api/refresh', methods=['POST'])
def api_refresh():
    update_all_from_folder()
    return jsonify({"ok": True})

@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    t = threading.Thread(target=start_watcher, daemon=True)
    t.start()
    app.run(debug=True)

