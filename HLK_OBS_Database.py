import streamlit as st
import sqlite3
import pandas as pd
import os
import re
import json
import logging
import time
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import subprocess
import sys
import importlib.util

# --- 必要套件自動安裝 ---
def ensure_extract_msg():
    if importlib.util.find_spec("extract_msg") is None:
        subprocess.check_call([sys.executable, "-m", "-q", "pip", "install", "extract-msg"])

ensure_extract_msg()
import extract_msg

# --- 目錄與資料庫設定 ---
UPLOAD_DIR = "hlk_attachments"
if not os.path.exists(UPLOAD_DIR):
    os.makedirs(UPLOAD_DIR)

DB_NAME = "hlk_test_tracker.db"
OUTLOOK_KEYWORD = "Observation Details"
OUTLOOK_SYNC_INTERVAL_SECONDS = 300

def connect_db():
    return sqlite3.connect(DB_NAME, timeout=30)

def init_db():
    conn = connect_db()
    conn.execute('''
        CREATE TABLE IF NOT EXISTS hlk_bugs (
            id TEXT PRIMARY KEY, short_desc TEXT, long_desc TEXT, 
            steps_to_reproduce TEXT, file_path TEXT, created_at TEXT
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS app_settings (
            key TEXT PRIMARY KEY, value TEXT NOT NULL
        )
    ''')
    conn.execute('''
        CREATE TABLE IF NOT EXISTS outlook_processed (
            store_id TEXT NOT NULL, entry_id TEXT NOT NULL,
            processed_at TEXT NOT NULL, outcome TEXT NOT NULL,
            obs_id TEXT,
            PRIMARY KEY (store_id, entry_id)
        )
    ''')
    columns = {row[1] for row in conn.execute("PRAGMA table_info(outlook_processed)")}
    if "obs_id" not in columns:
        conn.execute("ALTER TABLE outlook_processed ADD COLUMN obs_id TEXT")
    conn.commit(); conn.close()

def check_id_exists(sid):
    if not sid: return False
    conn = connect_db()
    res = conn.execute("SELECT id FROM hlk_bugs WHERE id=?", (sid,)).fetchone()
    conn.close()
    return res is not None

def get_setting(key):
    conn = connect_db()
    row = conn.execute("SELECT value FROM app_settings WHERE key=?", (key,)).fetchone()
    conn.close()
    return row[0] if row else None

def set_setting(key, value):
    conn = connect_db()
    try:
        with conn:
            conn.execute("INSERT OR REPLACE INTO app_settings (key, value) VALUES (?, ?)", (key, value))
    finally:
        conn.close()

def execute_db(query, parameters=()):
    conn = connect_db()
    try:
        conn.execute(query, parameters)
        conn.commit()
    finally:
        conn.close()

def delete_bug_record(sid):
    conn = connect_db()
    try:
        with conn:
            conn.execute("DELETE FROM hlk_bugs WHERE id=?", (sid,))
            conn.execute("DELETE FROM outlook_processed WHERE obs_id=?", (sid,))
    finally:
        conn.close()

def get_outlook_folders_setting():
    saved_value = get_setting("outlook_folders")
    if not saved_value:
        legacy_folder = get_setting("outlook_folder")
        return [legacy_folder] if legacy_folder else []
    try:
        folders = json.loads(saved_value)
    except json.JSONDecodeError:
        return [saved_value]
    return folders if isinstance(folders, list) else [saved_value]

def save_multiple_files(obs_id, uploaded_files):
    paths = []
    for u_file in uploaded_files:
        f_path = os.path.join(UPLOAD_DIR, f"{obs_id}_{u_file.name}")
        with open(f_path, "wb") as f:
            f.write(u_file.getbuffer())
        paths.append(f_path)
    return paths

# --- 解析與輔助邏輯 ---
def parse_mail_content(subject, body):
    body = body or ""
    subject = subject or ""
    full_text = f"{subject}\n{body}"
    id_m = re.search(r'OBS#?(SIO\d+)', full_text, re.IGNORECASE)
    sio_id = id_m.group(1) if id_m else ""
    sd_m = re.search(r'Short Description\s*:?\s*(.*)', body, re.IGNORECASE)
    short_desc = sd_m.group(1).strip() if sd_m else subject
    steps_content = ""
    steps_m = re.search(r'Steps to Reproduce(.*?)(?=Long Description|$)',
                        body, re.DOTALL | re.IGNORECASE)
    if steps_m:
        clean_lines = [line.strip() for line in steps_m.group(1).splitlines()
                       if "safelinks.protection" not in line]
        steps_content = "\n".join(line for line in clean_lines if line)
    ld_pattern = re.compile(r'(Long Description|Test Platform)(.*?)(?=EA Complete Date|$)',
                            re.DOTALL | re.IGNORECASE)
    ld_m = ld_pattern.search(body)
    long_content = ld_m.group(1) + ld_m.group(2).strip() if ld_m else ""
    return {"id": sio_id, "short": short_desc, "long": long_content, "steps": steps_content}

def parse_msg_file(uploaded_file):
    msg = extract_msg.Message(uploaded_file)
    return parse_mail_content(msg.subject, msg.body)

@st.cache_data(show_spinner=False)
def get_outlook_mail_folders():
    import pythoncom
    import pywintypes

    pythoncom.CoInitialize()
    try:
        for attempt in range(3):
            try:
                return _enumerate_outlook_mail_folders()
            except pywintypes.com_error as exc:
                hresult = exc.args[0] if exc.args and isinstance(exc.args[0], int) else 0
                if hresult & 0xFFFF != 0x0115 or attempt == 2:
                    raise
                time.sleep(attempt + 1)
    finally:
        pythoncom.CoUninitialize()

def _enumerate_outlook_mail_folders():
    import win32com.client

    namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    folders = []
    inbox_path = str(namespace.GetDefaultFolder(6).FolderPath).strip("\\")

    def visit(folder, path):
        folder_path = f"{path}\\{folder.Name}" if path else folder.Name
        folders.append(folder_path)
        for child in folder.Folders:
            visit(child, folder_path)

    for root in namespace.Folders:
        visit(root, "")
    return folders, inbox_path

def outlook_folder_error_message(exc):
    hresult = exc.args[0] if exc.args and isinstance(exc.args[0], int) else 0
    if hresult & 0xFFFF == 0x0115:
        return (
            "Outlook 無法連線到郵件伺服器。請確認 Outlook 已登入且連線正常，"
            f"等待郵件同步後重新整理頁面。原始錯誤：{exc}"
        )
    return f"無法讀取 Outlook 資料夾：{exc}"

def build_outlook_folder_tree(folder_paths):
    tree = {}
    for folder_path in folder_paths:
        nodes = tree
        path_parts = []
        for folder_name in folder_path.split("\\"):
            path_parts.append(folder_name)
            node = nodes.setdefault(folder_name, {
                "path": "\\".join(path_parts),
                "children": {}
            })
            nodes = node["children"]
    return tree

def render_outlook_folder_tree(tree, saved_folders, inbox_path, parent_parts=None):
    parent_parts = parent_parts or []
    inbox_parts = inbox_path.split("\\") if inbox_path else []
    priority_folder = (
        inbox_parts[len(parent_parts)]
        if inbox_parts[:len(parent_parts)] == parent_parts and len(inbox_parts) > len(parent_parts)
        else None
    )
    selected_folders = []
    folder_names = sorted(
        tree,
        key=lambda name: (name.casefold() != priority_folder.casefold()
                          if priority_folder else True, name.casefold())
    )
    for folder_name in folder_names:
        node = tree[folder_name]
        folder_path = node["path"]
        folder_parts = parent_parts + [folder_name]
        has_selected_descendant = any(
            saved_path.startswith(folder_path + "\\") for saved_path in saved_folders
        )
        is_inbox_branch = inbox_parts[:len(folder_parts)] == folder_parts
        if node["children"]:
            with st.expander(
                f"📁 {folder_name}",
                expanded=folder_path in saved_folders or has_selected_descendant or is_inbox_branch
            ):
                if st.checkbox(
                    "同步此資料夾（不含子資料夾）",
                    value=folder_path in saved_folders,
                    key=f"outlook_folder_{folder_path}"
                ):
                    selected_folders.append(folder_path)
                selected_folders.extend(
                    render_outlook_folder_tree(node["children"], saved_folders, inbox_path, folder_parts)
                )
        elif st.checkbox(
            f"📁 {folder_name}",
            value=folder_path in saved_folders,
            key=f"outlook_folder_{folder_path}"
        ):
            selected_folders.append(folder_path)
    return selected_folders

def get_outlook_folder(folder_path):
    import win32com.client

    namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    parts = folder_path.split("\\")
    folder = namespace.Folders.Item(parts[0])
    for part in parts[1:]:
        folder = folder.Folders.Item(part)
    store_id = folder.StoreID
    entry_id = folder.EntryID
    return namespace.GetFolderFromID(entry_id, store_id), store_id

def read_outlook_folder_messages(folder_path):
    import pythoncom

    def read_messages():
        folder, store_id = get_outlook_folder(folder_path)
        messages = []
        for item in folder.Items:
            if item.Class == 43:
                messages.append((str(item.EntryID), item.Subject or "", item.Body or ""))
        return store_id, messages

    pythoncom.CoInitialize()
    try:
        return read_messages()
    finally:
        pythoncom.CoUninitialize()

def sync_outlook_folder(folder_path):
    store_id, messages = read_outlook_folder_messages(folder_path)
    conn = connect_db()
    try:
        processed = {
            row[0]: (row[1], row[2])
            for row in conn.execute(
                "SELECT entry_id, outcome, obs_id FROM outlook_processed WHERE store_id=?",
                (store_id,)
            )
        }
        existing_ids = {row[0] for row in conn.execute("SELECT id FROM hlk_bugs")}
    finally:
        conn.close()

    pending = []
    for entry_id, subject, body in messages:
        previous = processed.get(entry_id)
        if previous:
            previous_outcome, previous_obs_id = previous
            if previous_outcome not in {"imported", "duplicate_obs_id"}:
                continue
            if OUTLOOK_KEYWORD.casefold() not in f"{subject}\n{body}".casefold():
                continue

            data = parse_mail_content(subject, body)
            obs_id = data["id"] or previous_obs_id
            if not obs_id:
                continue
            if data["id"] and data["short"] and data["id"] not in existing_ids:
                pending.append((entry_id, "imported", data, data["id"], True))
                existing_ids.add(data["id"])
            elif previous_obs_id is None:
                pending.append((entry_id, previous_outcome, None, obs_id, True))
            continue

        if OUTLOOK_KEYWORD.casefold() not in f"{subject}\n{body}".casefold():
            pending.append((entry_id, "keyword_not_found", None, None, False))
            continue

        data = parse_mail_content(subject, body)
        if not data["id"] or not data["short"]:
            pending.append((entry_id, "missing_required_fields", None, data["id"] or None, False))
        elif data["id"] in existing_ids:
            pending.append((entry_id, "duplicate_obs_id", None, data["id"], False))
        else:
            pending.append((entry_id, "imported", data, data["id"], False))
            existing_ids.add(data["id"])

    imported = skipped = unmatched = 0
    conn = connect_db()
    try:
        with conn:
            for entry_id, outcome, data, obs_id, previously_processed in pending:
                if previously_processed:
                    if data is not None:
                        cursor = conn.execute(
                            "INSERT OR IGNORE INTO hlk_bugs VALUES (?, ?, ?, ?, ?, ?)",
                            (data["id"], data["short"], data["long"], data["steps"], "No File",
                             datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                        )
                        if cursor.rowcount:
                            imported += 1
                        else:
                            outcome = "duplicate_obs_id"
                            skipped += 1
                    conn.execute(
                        "UPDATE outlook_processed SET processed_at=?, outcome=?, obs_id=? "
                        "WHERE store_id=? AND entry_id=?",
                        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"), outcome, obs_id,
                         store_id, entry_id)
                    )
                    continue

                cursor = conn.execute(
                    "INSERT OR IGNORE INTO outlook_processed "
                    "(store_id, entry_id, processed_at, outcome, obs_id) VALUES (?, ?, ?, ?, ?)",
                    (store_id, entry_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                     outcome, obs_id)
                )
                if cursor.rowcount == 0:
                    continue

                if data is not None:
                    cursor = conn.execute(
                        "INSERT OR IGNORE INTO hlk_bugs VALUES (?, ?, ?, ?, ?, ?)",
                        (data["id"], data["short"], data["long"], data["steps"], "No File",
                         datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                    )
                    if cursor.rowcount:
                        imported += 1
                    else:
                        conn.execute(
                            "UPDATE outlook_processed SET outcome=? WHERE store_id=? AND entry_id=?",
                            ("duplicate_obs_id", store_id, entry_id)
                        )
                        skipped += 1
                elif outcome == "keyword_not_found":
                    unmatched += 1
                else:
                    skipped += 1
    finally:
        conn.close()
    return imported, skipped, unmatched

def sync_outlook_folders(folder_paths):
    imported = skipped = unmatched = 0
    errors = []
    for folder_path in folder_paths:
        try:
            folder_imported, folder_skipped, folder_unmatched = sync_outlook_folder(folder_path)
            imported += folder_imported
            skipped += folder_skipped
            unmatched += folder_unmatched
        except Exception as exc:
            errors.append(f"{folder_path}：{exc}")
    return {
        "imported": imported,
        "skipped": skipped,
        "unmatched": unmatched,
        "errors": errors
    }

@st.cache_resource
def get_outlook_sync_service():
    return OutlookSyncService()

class OutlookSyncService:
    def __init__(self):
        self._executor = ThreadPoolExecutor(max_workers=1)
        self._lock = threading.Lock()
        self._future = None
        self._stop_event = threading.Event()
        self._scheduler_thread = threading.Thread(target=self._schedule_syncs, daemon=True)
        self._scheduler_thread.start()

    def submit(self, folder_paths):
        with self._lock:
            if self._future is not None and not self._future.done():
                return self._future
            self._future = self._executor.submit(sync_outlook_folders, list(folder_paths))
            self._future.add_done_callback(self._log_sync_result)
            return self._future

    def current_future(self):
        with self._lock:
            return self._future

    def is_running(self):
        future = self.current_future()
        return future is not None and not future.done()

    def _schedule_syncs(self):
        while not self._stop_event.wait(OUTLOOK_SYNC_INTERVAL_SECONDS):
            try:
                folder_paths = get_outlook_folders_setting()
                if folder_paths:
                    self.submit(folder_paths)
            except Exception:
                logging.exception("Scheduled Outlook synchronization failed")

    @staticmethod
    def _log_sync_result(future):
        try:
            result = future.result()
            for error in result["errors"]:
                logging.error("Outlook synchronization failed: %s", error)
        except Exception:
            logging.exception("Outlook synchronization failed")

def start_outlook_sync(folder_paths):
    st.session_state.outlook_sync_result = None
    st.session_state.outlook_sync_future = get_outlook_sync_service().submit(folder_paths)
    st.session_state.outlook_sync_attempted = True

def render_outlook_sync_status():
    future = st.session_state.get("outlook_sync_future")
    if future is None:
        latest_future = get_outlook_sync_service().current_future()
        if latest_future is not st.session_state.get("outlook_sync_displayed_future"):
            future = latest_future
            st.session_state.outlook_sync_future = future

    if future is not None:
        if not future.done():
            st.info("Outlook 郵件正在背景同步，您仍可使用資料庫其他功能。")
            return
        try:
            st.session_state.outlook_sync_result = future.result()
        except Exception as exc:
            st.session_state.outlook_sync_result = {"error": str(exc)}
        st.session_state.outlook_sync_displayed_future = future
        st.session_state.outlook_sync_future = None

    if get_outlook_sync_service().is_running():
        st.info("Outlook 郵件正在背景同步，您仍可使用資料庫其他功能。")
        return

    result = st.session_state.get("outlook_sync_result")
    if result is None:
        return
    if "error" in result:
        st.error(f"Outlook 同步失敗：{result['error']}")
        return
    st.info(
        f"Outlook 同步完成：新增 {result['imported']} 筆，略過 {result['skipped']} 筆，"
        f"不符合關鍵字 {result['unmatched']} 封。"
    )
    for error in result["errors"]:
        st.error(f"Outlook 同步失敗：{error}")

def main():
    st.set_page_config(page_title="HLK OBS Database", layout="wide")
    init_db()
    sync_service = get_outlook_sync_service()
    
    if "p_data" not in st.session_state:
        st.session_state.p_data = {"id": "", "short": "", "long": "", "steps": ""}
    if "form_id" not in st.session_state:
        st.session_state.form_id = 0

    st.title("🧪 HLK Test Fail Database (OBS)")
    col_l, col_r = st.columns([1, 1], gap="large")

    # --- 左側：管理區 ---
    with col_l:
        st.subheader("📝 紀錄管理")

        with st.expander("📨 Outlook 郵件自動匯入", expanded=True):
            st.caption(
                f"關鍵字：{OUTLOOK_KEYWORD}；OBS ID 重複的紀錄會跳過。"
                "頁面啟動後立即同步，之後每 5 分鐘背景同步一次；關閉瀏覽器不會停止，"
                "關閉程式命令視窗才會停止。"
            )
            try:
                outlook_folders, inbox_folder_path = get_outlook_mail_folders()
            except Exception as exc:
                outlook_folders = []
                inbox_folder_path = ""
                st.error(outlook_folder_error_message(exc))

            saved_folders = get_outlook_folders_setting()
            if st.button("重新讀取 Outlook 資料夾", disabled=sync_service.is_running()):
                get_outlook_mail_folders.clear()
                st.rerun()

            if outlook_folders:
                st.caption(
                    "展開資料夾即可勾選；勾選父資料夾只同步該資料夾，不會自動包含子資料夾。"
                    "新增或移除 Outlook 資料夾後，請按「重新讀取 Outlook 資料夾」。"
                )
                selected_folders = render_outlook_folder_tree(
                    build_outlook_folder_tree(outlook_folders), saved_folders, inbox_folder_path
                )
                sync_in_progress = sync_service.is_running()
                if st.button("儲存資料夾", disabled=sync_in_progress):
                    set_setting("outlook_folders", json.dumps(selected_folders, ensure_ascii=False))
                    start_outlook_sync(selected_folders)
                    st.rerun()

                if st.button(
                    "立即同步所選資料夾",
                    disabled=sync_in_progress or not selected_folders
                ):
                    set_setting("outlook_folders", json.dumps(selected_folders, ensure_ascii=False))
                    start_outlook_sync(selected_folders)
                    st.rerun()

                missing_folders = [folder for folder in saved_folders if folder not in outlook_folders]
                if missing_folders:
                    st.warning(f"找不到先前設定的資料夾：{', '.join(missing_folders)}")
                if (saved_folders and set(selected_folders) == set(saved_folders)
                        and not st.session_state.get("outlook_sync_attempted", False)):
                    start_outlook_sync(saved_folders)
            else:
                st.info("找不到可用的 Outlook 郵件資料夾。")
            render_outlook_sync_status()
        
        # 1. 單筆新增
        with st.expander("➕ 單筆新增紀錄 (含重複檢查)", expanded=False):
            msg_file = st.file_uploader("📥 拖入郵件檔", type=["msg"], key=f"u_{st.session_state.form_id}")
            if msg_file:
                try:
                    st.session_state.p_data = parse_msg_file(msg_file)
                    if check_id_exists(st.session_state.p_data["id"]):
                        st.warning(f"⚠️ ID {st.session_state.p_data['id']} 已存在，儲存將覆蓋！")
                except Exception as exc:
                    st.error(f"郵件解析失敗：{exc}")
            with st.form(key=f"add_f_{st.session_state.form_id}", clear_on_submit=True):
                c_id = st.text_input("OBS ID", value=st.session_state.p_data["id"])
                c_short = st.text_input("Short Description", value=st.session_state.p_data["short"])
                c_long = st.text_area("Long Description", value=st.session_state.p_data["long"], height=150)
                c_steps = st.text_area("Steps to Reproduce", value=st.session_state.p_data["steps"], height=150)
                c_files = st.file_uploader("上傳附件", accept_multiple_files=True)
                if st.form_submit_button("💾 儲存紀錄"):
                    if c_id and c_short:
                        fps = ";".join(save_multiple_files(c_id, c_files)) if c_files else "No File"
                        execute_db(
                            "INSERT OR REPLACE INTO hlk_bugs VALUES (?,?,?,?,?,?)",
                            (c_id, c_short, c_long, c_steps, fps,
                             datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                        )
                        st.session_state.p_data = {"id": "", "short": "", "long": "", "steps": ""}
                        st.session_state.form_id += 1
                        st.rerun()

        # 2. 批次匯入 (補回功能)
        with st.expander("📂 批次匯入 Excel (自動去重)", expanded=False):
            xl_file = st.file_uploader("選擇 Excel", type=["xlsx"], key="batch_xl")
            if st.button("🚀 開始解析並匯入"):
                if xl_file:
                    all_sheets = pd.read_excel(xl_file, sheet_name=None, header=None)
                    conn = connect_db()
                    added, skipped = 0, 0
                    for _, df in all_sheets.items():
                        obs_ids = re.findall(r'SIO\d+', df.to_string())
                        for sid in set(obs_ids):
                            if check_id_exists(sid):
                                skipped += 1; continue
                            conn.execute("INSERT INTO hlk_bugs (id, short_desc, created_at) VALUES (?,?,?)", (sid, "Batch Import", datetime.now().strftime("%Y-%m-%d %H:%M:%S")))
                            added += 1
                    conn.commit(); conn.close()
                    st.success(f"匯入完成！新增: {added} 筆, 跳過重複: {skipped} 筆")
                    st.rerun()

    # --- 右側：詳細資訊、編輯與刪除 ---
    with col_r:
        st.subheader("ℹ️ 詳細資料與修改")
        conn = connect_db()
        df_all = pd.read_sql_query("SELECT * FROM hlk_bugs ORDER BY created_at DESC", conn)
        conn.close()

        if "selected_row_id" in st.session_state and not df_all.empty:
            res_df = df_all[df_all['id'] == st.session_state.selected_row_id]
            if not res_df.empty:
                res = res_df.iloc[0]; sid = res['id']
                if st.session_state.get(f"m_{sid}") != "edit":
                    b1, b2 = st.columns(2)
                    with b1:
                        if st.button("✏️ 進入編輯", key=f"eb_{sid}"):
                            st.session_state[f"m_{sid}"] = "edit"
                            st.session_state[f"temp_fps_{sid}"] = [p for p in res['file_path'].split(";") if p != "No File" and p != ""]
                            st.rerun()
                    with b2:
                        if st.button("🗑️ 刪除", type="primary", key=f"db_{sid}"):
                            delete_bug_record(sid)
                            del st.session_state.selected_row_id; st.rerun()
                    
                    for label, color, field in [("OBS ID", "#1e3a5f", 'id'), ("Short Description", "#1b4d3e", 'short_desc'), ("Long Description", "#4d4d1b", 'long_desc'), ("Steps to Reproduce", "#5f1e1e", 'steps_to_reproduce')]:
                        st.markdown(f"<div style='background-color: {color}; padding: 8px; border-radius: 5px; color: white; font-weight: bold;'>{label}:</div>", unsafe_allow_html=True)
                        if label == "OBS ID": st.write(f"&nbsp;&nbsp;{res[field]}")
                        else: st.code(res[field], language='text')
                        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
                    
                    if res['file_path'] and res['file_path'] != "No File":
                        st.write("**📎 附件:**")
                        for p in res['file_path'].split(";"):
                            if os.path.exists(p):
                                fn = os.path.basename(p)
                                c1, c2 = st.columns([4, 1])
                                with c1:
                                    if st.button(f"📂 開啟 {fn}", key=f"v_{p}"): os.startfile(os.path.abspath(p))
                                with c2:
                                    with open(p, "rb") as f: st.download_button("📥", f, file_name=fn, key=f"dl_{p}")
                            else: st.error(f"找不到檔案: {os.path.basename(p)}")
                else:
                    # 編輯模式 (含附件刪除與取消)
                    e_s = st.text_input("Short", res['short_desc'])
                    e_l = st.text_area("Long", res['long_desc'], height=150)
                    e_t = st.text_area("Steps", res['steps_to_reproduce'], height=150)
                    st.write("**📁 附件管理 (❌ 刪除):**")
                    current_files = st.session_state.get(f"temp_fps_{sid}", [])
                    files_to_keep = []
                    for p in current_files:
                        ec1, ec2 = st.columns([5, 1])
                        ec1.text(f"📄 {os.path.basename(p)}")
                        if not ec2.button("❌", key=f"df_{p}"): files_to_keep.append(p)
                    st.session_state[f"temp_fps_{sid}"] = files_to_keep
                    e_f = st.file_uploader("補充附件", accept_multiple_files=True, key=f"efu_{sid}")
                    eb1, eb2 = st.columns(2)
                    with eb1:
                        if st.button("💾 儲存修改"):
                            final_paths = st.session_state[f"temp_fps_{sid}"]
                            if e_f: final_paths.extend(save_multiple_files(sid, e_f))
                            fps_str = ";".join(final_paths) if final_paths else "No File"
                            execute_db(
                                "UPDATE hlk_bugs SET short_desc=?, long_desc=?, steps_to_reproduce=?, file_path=? WHERE id=?",
                                (e_s, e_l, e_t, fps_str, sid)
                            )
                            st.session_state[f"m_{sid}"] = "view"; st.rerun()
                    with eb2:
                        if st.button("❌ 取消"): st.session_state[f"m_{sid}"] = "view"; st.rerun()
        else: st.info("請從下表選取紀錄")

    # --- 下方：檢索區 (筆數統計與 5 筆高度) ---
    st.markdown("---")
    st.subheader("🔍 歷史問題檢索")
    sk = st.text_input("搜尋...", placeholder="Search...")
    df_disp = df_all.copy()
    if sk: df_disp = df_disp[df_disp.apply(lambda r: r.astype(str).str.contains(sk, case=False).any(), axis=1)]
    
    t1, t2, _ = st.columns([1, 1, 4])
    t1.metric("搜尋筆數", len(df_disp)); t2.metric("總筆數", len(df_all))

    event = st.dataframe(df_disp, use_container_width=True, on_select="rerun", selection_mode="single-row", height=225, hide_index=True, column_config={"long_desc": None, "steps_to_reproduce": None, "file_path": None, "created_at": "時間"})
    if len(event.selection["rows"]) > 0:
        new_id = df_disp.iloc[event.selection["rows"][0]]['id']
        if st.session_state.get("selected_row_id") != new_id:
            st.session_state.selected_row_id = new_id; st.rerun()

if __name__ == '__main__':
    main()