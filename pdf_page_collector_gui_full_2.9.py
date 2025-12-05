import os
import time
import threading
import tkinter as tk
from tkinter import messagebox, scrolledtext, filedialog, ttk
import json
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import PyPDF2
import unicodedata
import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
from win10toast import ToastNotifier
from pdf_list_find_write import PDFPassengerSearchApp
from excel_write_preview_gui import NSExcelPreviewer
import queue
import sys
import winsound
import socket

ock_socket = None  # ← これがロック保持に必要

if getattr(sys, 'frozen', False):
    # PyInstaller で exe 化した場合
    base_dir = os.path.dirname(sys.executable)
else:
    # スクリプトとして実行している場合
    base_dir = os.path.dirname(os.path.abspath(__file__))


MAIN_FILE = os.path.join(base_dir, "出力便名リスト.txt")
CONFIG_FILE = os.path.join(base_dir, "config.json")
WATCH_FOLDER = base_dir
OUTPUT_FOLDER = base_dir
LOCK_FILE = os.path.join(base_dir, "app.lock")

ICON_FILE = os.path.join(base_dir, "tray_icon.png")


# =====================
# 重複起動防止準備
# =====================
def acquire_single_instance_lock(port=56789):
    global lock_socket
    try:
        lock_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        lock_socket.bind(("127.0.0.1", port))  # ポート確保
        lock_socket.listen(1)  # listen してロック維持
        return True
    except OSError:
        return False

# =====================
# 設定ロード/保存
# =====================
def load_config(ben_list):
    cfg = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE,"r",encoding="utf-8") as f:
                cfg = json.load(f)
        except:
            cfg = {}
    config={}
    for ben in ben_list:
        config[ben]={
            "座席表": cfg.get(ben,{}).get("座席表",False),
            "バス号車別明細表_乗務員用": cfg.get(ben,{}).get("バス号車別明細表_乗務員用",False),
            "バス号車別明細表_保管用": cfg.get(ben,{}).get("バス号車別明細表_保管用",False)
        }
    return config

def save_config(config):
    with open(CONFIG_FILE,"w",encoding="utf-8") as f:
        json.dump(config,f,ensure_ascii=False,indent=2)

# =====================
# PDF抽出
# （省略せず既存のまま）
# =====================
def extract_pdf_by_criteria(pdf_folder, ben_list, config, output_folder, log_queue, status_queue, folder_display):
    import re
    import hashlib
    import unicodedata
    import PyPDF2
    import os

    # --- テキスト正規化 ---
    def normalize_text(s):
        if s is None:
            return ""
        s = unicodedata.normalize("NFKC", s)
        s = s.replace("\u3000", " ").replace("\u200b", "")
        return re.sub(r"\s+", " ", s).strip().lower()

    # --- 厳密マッチ ---
    def keyword_strict_match(norm_text, norm_kw):
        boundary_chars = r"0-9A-Za-z\u4e00-\u9fff\u3040-\u309f\u30a0-\u30ff"
        strict_pat = rf"(?<![{boundary_chars}]){re.escape(norm_kw)}(?![{boundary_chars}])"
        if re.search(strict_pat, norm_text):
            return True
        spaced = "".join([re.escape(ch) + r"\s*" for ch in norm_kw])
        spaced_pat = rf"(?<![{boundary_chars}]){spaced}(?![{boundary_chars}])"
        return re.search(spaced_pat, norm_text) is not None

    # --- ファイル内容ハッシュ ---
    def file_hash(path):
        h = hashlib.md5()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(4096), b""):
                h.update(chunk)
        return h.hexdigest()

    log_queue.put(f"[INFO] PDF抽出開始: {folder_display} ({pdf_folder})")

    # --- PDFファイル取得（重複判定あり） ---
    pdf_files = []
    seen_names = set()
    seen_hashes = set()
    for f in os.listdir(pdf_folder):
        if not f.lower().endswith(".pdf"):
            continue
        full_path = os.path.join(pdf_folder, f)
        norm_name = f.lower().strip()

        # ファイル名重複
        if norm_name in seen_names:
            log_queue.put(f"[SKIP] 重複ファイル名: {f}")
            continue
        seen_names.add(norm_name)

        # 内容重複
        try:
            h = file_hash(full_path)
            if h in seen_hashes:
                log_queue.put(f"[SKIP] 重複内容: {f}")
                continue
            seen_hashes.add(h)
        except Exception as e:
            log_queue.put(f"[WARN] ハッシュ計算失敗: {f} ({e})")
            continue

        pdf_files.append(full_path)

    if not pdf_files:
        log_queue.put(f"[INFO] PDFなし: {folder_display}")
        return

    # --- ページ抽出 ---
    intermediate_files = []
    extract_counts = {
        ben: {
            "座席表": 0,
            "バス号車別明細表(乗務員用)": 0,
            "バス号車別明細表(保管用)": 0
        }
        for ben in ben_list
    }

    for pdf_path in pdf_files:
        fname = os.path.basename(pdf_path)
        try:
            reader = PyPDF2.PdfReader(pdf_path)
        except Exception as e:
            log_queue.put(f"[ERROR] {fname} 読み込み失敗 ({e})")
            continue

        for i, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            norm_text = normalize_text(text)

            for ben in ben_list:
                if not keyword_strict_match(norm_text, normalize_text(ben)):
                    continue

                # 座席表
                if "座席表" in text:
                    if config[ben]["座席表"]:
                        intermediate_files.append(("乗務員用", ben, "座席表", page))
                        extract_counts[ben]["座席表"] += 1  # ✅ 抽出数カウント
                        log_queue.put(f"[PAGE] {fname} → {ben} 座席表")
                        status_queue.put((ben, "座席表", 1))
                    else:
                        status_queue.put((ben, "座席表", 0))  # 赤判定（印刷OFF）

                # バス号車別明細
                if "バス号車別明細表" in text:
                    # 乗務員用
                    if config[ben]["バス号車別明細表_乗務員用"]:
                        intermediate_files.append(("乗務員用", ben, "バス号車別明細表", page))
                        extract_counts[ben]["バス号車別明細表(乗務員用)"] += 1  # ✅ 抽出数カウント
                        log_queue.put(f"[PAGE] {fname} → {ben} バス号車別明細表(乗務員用)")
                        status_queue.put((ben, "バス号車別明細表(乗務員用)", 1))
                    else:
                        status_queue.put((ben, "バス号車別明細表(乗務員用)", 0))

                    # 保管用
                    if config[ben]["バス号車別明細表_保管用"]:
                        intermediate_files.append(("保管用", ben, "バス号車別明細表", page))
                        extract_counts[ben]["バス号車別明細表(保管用)"] += 1  # ✅ 抽出数カウント
                        log_queue.put(f"[PAGE] {fname} → {ben} バス号車別明細表(保管用)")
                        status_queue.put((ben, "バス号車別明細表(保管用)", 1))
                    else:
                        status_queue.put((ben, "バス号車別明細表(保管用)", 0))

    # --- 黄色判定（印刷ONなのに抽出なし） ---
    for ben in ben_list:
        if config[ben]["座席表"] and extract_counts[ben]["座席表"] == 0:
            status_queue.put((ben, "座席表", 0))
        if config[ben]["バス号車別明細表_乗務員用"] and extract_counts[ben]["バス号車別明細表(乗務員用)"] == 0:
            status_queue.put((ben, "バス号車別明細表(乗務員用)", 0))
        if config[ben]["バス号車別明細表_保管用"] and extract_counts[ben]["バス号車別明細表(保管用)"] == 0:
            status_queue.put((ben, "バス号車別明細表(保管用)", 0))

    if not intermediate_files:
        log_queue.put(f"[INFO] 抽出結果なし: {folder_display}（PDFは出力しません）")
        return

    # --- PDF出力 ---
    for mode in ["乗務員用", "保管用"]:
        writer = PyPDF2.PdfWriter()
        page_count = 0

        if mode == "乗務員用":
            for ben in ben_list:
                for typ in ["座席表", "バス号車別明細表"]:
                    for entry in intermediate_files:
                        if entry[0] == mode and entry[1] == ben and entry[2] == typ:
                            writer.add_page(entry[3])
                            page_count += 1
            out_path = os.path.join(output_folder, f"{folder_display}_乗務員用.pdf")

        else:  # 保管用
            for ben in reversed(ben_list):
                for entry in intermediate_files:
                    if entry[0] == mode and entry[1] == ben and entry[2] == "バス号車別明細表":
                        writer.add_page(entry[3])
                        page_count += 1
            out_path = os.path.join(output_folder, f"{folder_display}_保管用.pdf")

        if page_count > 0:
            with open(out_path, "wb") as f:
                writer.write(f)
            log_queue.put(f"[DONE] {mode}PDF出力: {out_path}")
        else:
            log_queue.put(f"[SKIP] {mode}PDFは出力対象ページなし（スキップ）")
    
    set_current_folder(f"{folder_display}　抽出完了")
    
     # --- 抽出完了通知音 ---
    sound_path = os.path.join(base_dir, "finish_sound.wav")  # または .wav
    if os.path.exists(sound_path):
        try:
            threading.Thread(target=lambda: winsound.PlaySound(sound_path, winsound.SND_FILENAME | winsound.SND_ASYNC), daemon=True).start()
        except Exception as e:
            log_queue.put(f"[WARN] 音声再生失敗: {e}")
    else:
        log_queue.put("[INFO] 通知音ファイルが見つからなかったため、音声再生をスキップしました。")


# =====================
# フォルダ監視
# =====================
import os
import time
import threading
from watchdog.events import FileSystemEventHandler

class FolderHandler(FileSystemEventHandler):
    def __init__(self, log_queue, notify_func, ben_list, config,
                 status_queue, reset_status_callback=None,
                 bring_front_callback=None,
                 set_current_folder_callback=None):  # ★ 追加
        self.reset_status_callback = reset_status_callback
        self.bring_front_callback = bring_front_callback
        self.set_current_folder_callback = set_current_folder_callback  # ★ 追加
        self.log_queue = log_queue
        self.notify_func = notify_func
        self.processed = set()
        self.ben_list = ben_list
        self.config = config
        self.status_queue = status_queue

    def on_created(self, event): self._check_folder(event)
    def on_moved(self, event): self._check_folder(event)
    def on_modified(self, event): self._check_folder(event)

    def _check_folder(self, event):
        # フォルダ以外（PDFなどのファイル変更）は無視する
        if not getattr(event, "is_directory", False):
            return

        folder_path = getattr(event, 'dest_path', event.src_path)
        folder_name = os.path.basename(folder_path)
        today = time.strftime("%m.%d")

        if "出発名簿" in folder_name and today in folder_name and "●" in folder_name:
            if folder_name not in self.processed:
                self.processed.add(folder_name)

                # ★ フォルダ検知時点でUI更新
                if callable(self.set_current_folder_callback):
                    self.set_current_folder_callback(folder_name)

                if self.reset_status_callback:
                    self.reset_status_callback()

                if self.bring_front_callback:
                    self.bring_front_callback()

                self.log_queue.put(f"[INFO] 検知対象フォルダ: {folder_name}")
                try:
                    self.notify_func("フォルダ検出", f"{folder_name} の抽出を開始します")
                except:
                    pass

                threading.Thread(
                    target=self.process_folder,
                    args=(folder_path, folder_name),
                    daemon=True
                ).start()

    def process_folder(self, folder_path, folder_name):
        # 念のため開始時にもう一度ラベルを更新
        if callable(self.set_current_folder_callback):
            try:
                self.set_current_folder_callback(folder_name)
            except Exception as e:
                self.log_queue.put(f"[WARN] set_current_folder_callback失敗: {e}")

        try:
            extract_pdf_by_criteria(
                folder_path,
                self.ben_list,
                self.config,
                OUTPUT_FOLDER,
                self.log_queue,
                self.status_queue,
                folder_name
            )
        except Exception as e:
            self.log_queue.put(f"[ERROR] {e}")

# =====================
# 起動時/手動フォルダスキャン
# =====================
def scan_existing_folders(ben_list, config, log_queue, status_queue, notify_func, ignore_dot=False):
    today = time.strftime("%m.%d")
    for fname in os.listdir(WATCH_FOLDER):
        folder_path = os.path.join(WATCH_FOLDER, fname)
        if os.path.isdir(folder_path) and "出発名簿" in fname and today in fname:
            if ignore_dot or "●" in fname:
                set_current_folder(fname)
                threading.Thread(
                    target=extract_pdf_by_criteria,
                    args=(folder_path, ben_list, config, OUTPUT_FOLDER, log_queue, status_queue, fname),
                    daemon=True
                ).start()
                log_queue.put(f"[INFO] フォルダを検知・処理開始: {fname}")
                try:
                    notify_func("フォルダ検知", f"{fname} の抽出を開始します")
                except:
                    pass

# =====================
# GUI + トレイ + ステータスウィンドウ統合
# =====================
def run_gui():
    global WATCH_FOLDER, OUTPUT_FOLDER
    global set_current_folder
    root = tk.Tk()
    root.withdraw()  # メインウィンドウ非表示
    root.title("📄 出発名簿自動PDF抽出ツール")
    root.geometry("700x500")
    root.configure(bg="#f4f6f8")

    # --- メインログUI ---
    tk.Label(root, text="📑 出発名簿監視システム", bg="#f4f6f8", font=("Segoe UI", 15, "bold")).pack(pady=10)
    log_box = scrolledtext.ScrolledText(root, wrap="word", font=("Consolas", 10), height=20, state="disabled")
    log_box.pack(fill="both", expand=True, padx=15, pady=10)
    watch_label = tk.Label(root, text="", bg="#f4f6f8")
    watch_label.pack()

    log_queue = queue.Queue()
    status_queue = queue.Queue()
    exit_queue = queue.Queue()

    # --- ログ更新 ---
    def poll_log_queue():
        while True:
            try:
                msg = log_queue.get_nowait()
            except queue.Empty:
                break
            else:
                try:
                    log_box.configure(state="normal")
                    log_box.insert(tk.END, msg + "\n")
                    log_box.see(tk.END)
                    log_box.configure(state="disabled")
                except tk.TclError:
                    break

        root.after(200, poll_log_queue)
    root.after(200, poll_log_queue)

    # --- 終了監視 ---
    def poll_exit_queue():
        while True:
            try:
                exit_queue.get_nowait()
            except queue.Empty:
                break
            else:
                if root.winfo_exists():
                    root.quit()
                    root.destroy()
        root.after(200, poll_exit_queue)
    root.after(200, poll_exit_queue)

    # --- トースト通知 ---
    toast = ToastNotifier()
    def tray_notify(title, message):
        def _notify():
            toast.show_toast(title, message, duration=5, threaded=True)
        root.after(0, _notify)

    # --- 出発便リスト ---
    ben_list = []
    if os.path.exists(MAIN_FILE):
        with open(MAIN_FILE,"r",encoding="utf-8") as f:
            ben_list = [line.strip() for line in f if line.strip()]

    # --- config 読み込み ---
    cfg = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
        except:
            cfg = {}
    if "folders" not in cfg:
        cfg["folders"] = {"watch_folder": base_dir, "output_folder": base_dir}
    WATCH_FOLDER = cfg["folders"].get("watch_folder", base_dir)
    OUTPUT_FOLDER = cfg["folders"].get("output_folder", base_dir)
    # --- config 読み込み ---
    if "ben_settings" not in cfg:
        cfg["ben_settings"] = load_config(ben_list)

    config = cfg["ben_settings"]

    # ★ ここを追加：便名リストに合わせて不足設定を自動追加
    for ben in ben_list:
        if ben not in config:
            config[ben] = {
                "座席表": False,
                "バス号車別明細表_乗務員用": False,
                "バス号車別明細表_保管用": False
            }

    # ついでに削除された便名は消しておく（※任意）
    for ben in list(config.keys()):
        if ben not in ben_list:
            del config[ben]

    save_config(cfg)


    if not os.path.isdir(WATCH_FOLDER):
        log_queue.put(f"[WARNING] 監視フォルダが存在しません: {WATCH_FOLDER} → デフォルトに変更")
        WATCH_FOLDER = base_dir
    if not os.path.isdir(OUTPUT_FOLDER):
        log_queue.put(f"[WARNING] PDF出力フォルダが存在しません: {OUTPUT_FOLDER} → デフォルトに変更")
        OUTPUT_FOLDER = base_dir
    watch_label.config(text=f"監視フォルダ: {WATCH_FOLDER}")

    # --- ステータスウィンドウ ---
    status_window = tk.Toplevel(root)
    status_window.title("乗客名簿出力ステータス")

    current_folder_label = tk.Label(status_window, text="抽出フォルダー：-", bg="#eef", fg="black", anchor="w")
    current_folder_label.grid(row=0, column=0, columnspan=4, sticky="nsew", padx=1, pady=(3,6))

    columns = ["座席表","バス号車(乗務員用)","バス号車(保存用)"]
    status_labels = {}

    for r, ben in enumerate(ben_list):
        status_labels[ben] = {}
        lbl_ben = tk.Label(status_window, text=ben, width=20, relief="ridge", bg="white")
        lbl_ben.grid(row=r+2, column=0, sticky="nsew", padx=1, pady=1)
        status_labels[ben]["便名"] = lbl_ben
        for c, col in enumerate(columns, start=1):
            lbl = tk.Label(status_window, text="0", width=15, relief="ridge", bg="white")
            lbl.grid(row=r+2, column=c, sticky="nsew", padx=1, pady=1)
            status_labels[ben][col] = lbl

    tk.Label(status_window, text="便名", relief="ridge", bg="#cccccc").grid(row=1, column=0, sticky="nsew")
    for c, col in enumerate(columns, start=1):
        tk.Label(status_window, text=col, relief="ridge", bg="#cccccc").grid(row=1, column=c, sticky="nsew")

    status_window.update_idletasks()
    status_window.geometry(f"{status_window.winfo_reqwidth()}x{status_window.winfo_reqheight()}")
    status_window.resizable(False, False)
    status_window.attributes('-toolwindow', True)
    status_visible = [True]

     # ❌ 閉じるボタンを完全に無効化（何も起きない）
    def disable_close_button():
        pass  # 何もしない
    status_window.protocol("WM_DELETE_WINDOW", disable_close_button)

    # --- ステータス画面リセット関数 ---
    def reset_status_display():
        """ステータス画面の内容をリセット"""
        for ben in status_labels:
            for col in status_labels[ben]:
                lbl = status_labels[ben][col]
                if col == "便名":
                    lbl.config(bg="white")
                else:
                    lbl.config(text="0", bg="white")
        log_queue.put("[INFO] ステータス画面をリセットしました。")

        #handler = FolderHandler(log_queue, tray_notify, ben_list, config, status_queue, reset_status_callback=reset_status_display)
        
    
    # --- ステータス画面を最前面に出す関数 ---
    def bring_status_to_front():
        try:
            status_window.attributes('-topmost', True)
            status_window.lift()
            status_window.focus_force()
            root.after(1000, lambda: status_window.attributes('-topmost', False))  # 1秒後解除
        except Exception as e:
            log_queue.put(f"[ERROR] ステータス前面化エラー: {e}")

    
    # --- ステータス更新ループ ---
    # --- ステータス更新ループ --- の変更点を移植
    def update_status_loop():
        while True:
            try:
                ben, item_name, count = status_queue.get_nowait()
            except queue.Empty:
                break
            else:
                col_name_map = {
                    "座席表": "座席表",
                    "バス号車別明細表(乗務員用)": "バス号車(乗務員用)",
                    "バス号車別明細表(保管用)": "バス号車(保存用)"
                }
                if item_name not in col_name_map:
                    continue
                col_name = col_name_map[item_name]

                lbl = status_labels[ben][col_name]
                current = int(lbl.cget("text"))
                new_count = current + max(0, count)  # ←マイナス防止
                lbl.config(text=str(new_count))

                # 子セル色判定
                config_key_map = {
                    "座席表": "座席表",
                    "バス号車別明細表(乗務員用)": "バス号車別明細表_乗務員用",
                    "バス号車別明細表(保管用)": "バス号車別明細表_保管用"
                }
                checked_key = config_key_map[item_name]
                checked = config[ben].get(checked_key, False)

                if checked:  # 印刷ON
                    if count > 0:
                        color = "lightgreen"  # 抽出あり
                    else:
                        color = "yellow"      # 抽出なし → 黄色
                else:        # 印刷OFF
                    if count > 0:
                        color = "red"         # 抽出あり
                    else:
                        color = "white"       # 抽出なし
                lbl.config(bg=color)

        # 便名セル色判定
        for ben in status_labels:
            child_colors = [status_labels[ben][col].cget("bg") for col in columns]

            # 赤が1つでも
            if "red" in child_colors:
                ben_color = "red"
            # 緑＋白だけ
            elif all(c in ("lightgreen","white") for c in child_colors) and "lightgreen" in child_colors:
                ben_color = "lightgreen"
            # 白＋黄色だけ or 緑＋黄色
            elif any(c=="yellow" for c in child_colors) and all(c in ("white","yellow","lightgreen") for c in child_colors):
                ben_color = "#d9d9d9"
            # 全て白
            else:
                ben_color = "white"

            status_labels[ben]["便名"].config(bg=ben_color)

        status_window.after(200, update_status_loop)

    status_window.after(200, update_status_loop)


    def set_current_folder(name: str):
        def _upd():
            current_folder_label.config(text=f"抽出フォルダー：{name}")

        root.after(0, _upd)  # UIスレッドへ確実に投げる



        # --- 子ウィンドウ管理 ---
    child_windows = {}

    def toggle_status_window(icon=None, item=None):
        if status_visible[0]:
            status_window.withdraw()
            status_visible[0] = False
        else:
            status_window.deiconify()
            status_window.lift()
            status_visible[0] = True

    def show_status_window(icon=None, item=None):
        status_window.deiconify()
        status_window.lift()
        status_visible[0] = True

    # --- フォルダー設定 ---
    def open_folder_settings(icon_obj=None, item=None):
        if "folder" in child_windows and child_windows["folder"].winfo_exists():
            child_windows["folder"].lift()
            return
        folder_win = tk.Toplevel()
        folder_win.title("フォルダー設定")
        folder_win.configure(bg="#f4f6f8")
        folder_win.grab_set()
        frm = tk.Frame(folder_win, bg="#f4f6f8")
        frm.pack(padx=15, pady=10, fill="both", expand=True)
        tk.Label(frm, text="監視フォルダ:", bg="#f4f6f8").pack(anchor="w")
        watch_entry = tk.Entry(frm, width=60)
        watch_entry.insert(0, WATCH_FOLDER)
        watch_entry.pack(pady=5)
        tk.Button(frm, text="参照", command=lambda: watch_entry.delete(0, tk.END) or watch_entry.insert(0, filedialog.askdirectory())).pack(pady=5)
        tk.Label(frm, text="PDF出力フォルダ:", bg="#f4f6f8").pack(anchor="w")
        output_entry = tk.Entry(frm, width=60)
        output_entry.insert(0, OUTPUT_FOLDER)
        output_entry.pack(pady=5)
        tk.Button(frm, text="参照", command=lambda: output_entry.delete(0, tk.END) or output_entry.insert(0, filedialog.askdirectory())).pack(pady=5)
        def apply():
            global WATCH_FOLDER, OUTPUT_FOLDER
            w,o = watch_entry.get(), output_entry.get()
            if os.path.isdir(w) and os.path.isdir(o):
                WATCH_FOLDER, OUTPUT_FOLDER = w,o
                watch_label.config(text=f"監視フォルダ: {WATCH_FOLDER}")
                cfg["folders"]["watch_folder"], cfg["folders"]["output_folder"] = w,o
                save_config(cfg)
                folder_win.destroy()
            else:
                messagebox.showerror("エラー","有効なフォルダを選択してください")
        tk.Button(frm, text="保存して閉じる", command=apply).pack(pady=15)
        child_windows["folder"] = folder_win

    # --- 印刷設定画面を正常動作版に置換 ---
    def open_print_settings(icon_obj=None,item=None):
        if "print" in child_windows and child_windows["print"].winfo_exists():
            child_windows["print"].lift()
            return

        ps_win = tk.Toplevel()
        ps_win.title("印刷設定")
        ps_win.grab_set()
        ps_win.focus_set()

        main_frame = tk.Frame(ps_win)
        main_frame.pack(padx=10,pady=10)

        for ben in ben_list:
            frame = tk.Frame(main_frame)
            frame.pack(fill="x", pady=2)
            tk.Label(frame, text=ben, width=20, anchor="w").pack(side="left")
            var_seat = tk.BooleanVar(value=config[ben]["座席表"])
            tk.Checkbutton(frame,text="座席表",variable=var_seat).pack(side="left")
            var_bus_crew = tk.BooleanVar(value=config[ben]["バス号車別明細表_乗務員用"])
            tk.Checkbutton(frame,text="バス号車別明細表(乗務員用)",variable=var_bus_crew).pack(side="left")
            var_bus_store = tk.BooleanVar(value=config[ben]["バス号車別明細表_保管用"])
            tk.Checkbutton(frame,text="バス号車別明細表(保管用)",variable=var_bus_store).pack(side="left")

            def make_update(ben,var_seat,var_bus_crew,var_bus_store):
                def update():
                    config[ben]["座席表"] = var_seat.get()
                    config[ben]["バス号車別明細表_乗務員用"] = var_bus_crew.get()
                    config[ben]["バス号車別明細表_保管用"] = var_bus_store.get()
                    cfg["ben_settings"] = config
                    save_config(cfg)
                return update

            for cb in frame.winfo_children()[1:]:
                cb.configure(command=make_update(ben,var_seat,var_bus_crew,var_bus_store))

        ps_win.update_idletasks()
        ps_win.geometry(f"{main_frame.winfo_reqwidth()+20}x{main_frame.winfo_reqheight()+20}")
        child_windows["print"] = ps_win

    
    def open_passenger_search(icon=None, item=None):
        """乗客名簿検索ツールを別ウィンドウとして開く"""
        if "passenger" in child_windows and child_windows["passenger"].winfo_exists():
            child_windows["passenger"].lift()
            return

        top = tk.Toplevel()
        top.title("乗客名簿検索ツール")
        app = PDFPassengerSearchApp(top)
        top.geometry("1200x800")
        child_windows["passenger"] = top


    def open_excel_write_preview(icon=None, item=None):
            """乗客名簿検索ツールを別ウィンドウとして開く"""
            if "excel" in child_windows and child_windows["excel"].winfo_exists():
                child_windows["excel"].lift()
                return

            top = tk.Toplevel()
            top.title("NS報告作成ツール")
            app = NSExcelPreviewer(top)
            top.geometry("1200x800")
            child_windows["excel"] = top

    # --- 手動抽出 ---
    def manual_extract(*args):
        log_queue.put("[INFO] 手動抽出開始")
        bring_status_to_front()  # ★ 追加
        reset_status_display()  # ★ ステータスリセットを追加
        threading.Thread(target=lambda: scan_existing_folders(ben_list, config, log_queue, status_queue, tray_notify, ignore_dot=True), daemon=True).start()

    # --- トレイアイコン ---
    def load_tray_icon():
        try:
            return Image.open(ICON_FILE)  # PNG 読み込み
        except:
            # 読み込めなかった場合は fallback で簡易アイコン生成
            img = Image.new("RGB", (64,64), (200,200,200))
            d = ImageDraw.Draw(img)
            d.text((10,20), "PDF", fill=(0,0,0))
            return img

    tray_icon = pystray.Icon("pdf_watcher", load_tray_icon(), "出発名簿監視")
    tray_thread_started = [False]
    def start_tray_icon_once():
        if not tray_thread_started[0]:
            tray_icon.run_detached()
            tray_thread_started[0] = True

    def show_window(icon=None,item=None):
        if root.winfo_exists():
            root.deiconify()
            root.lift()
            root.after(500, lambda: root.attributes("-topmost", False))

    def quit_app(icon=None,item=None):
        try:
            if observer:  # observer が定義済みか確認
                observer.stop()
                observer.join(timeout=1)
        except NameError:
            pass
        try:
            tray_icon.stop()
        except:
            pass
        exit_queue.put(True)

    tray_icon.menu = pystray.Menu(
        item("表示", show_window),
        item(lambda text: "ステータス表示OFF" if status_visible[0] else "ステータス表示ON", toggle_status_window),
        item("フォルダー設定", open_folder_settings),
        item("印刷設定", open_print_settings),
        item("手動抽出", manual_extract),
        item("乗客名簿検索ツール（テスト用）", open_passenger_search),
        item("NS報告作成ツール（テスト用）", open_excel_write_preview),
        item("終了", quit_app)
    )

    # --- フォルダ監視 ---
    handler = FolderHandler(
        log_queue, tray_notify, ben_list, config, status_queue,
        reset_status_callback=reset_status_display,
        bring_front_callback=bring_status_to_front,
        set_current_folder_callback=set_current_folder  # ★ ここで渡す
    )
    #handler.set_current_folder = set_current_folder  # ← これが有効に働く
    observer = Observer()
    observer.schedule(handler, WATCH_FOLDER, recursive=False)
    observer.start()
    log_queue.put(f"[INFO] 監視開始: {WATCH_FOLDER}")
    scan_existing_folders(ben_list, config, log_queue, status_queue, tray_notify)

    # --- 起動時に常駐トレイ表示 ---
    start_tray_icon_once()
    log_queue.put("[INFO] 常駐トレイ起動")

    # --- 閉じるときは最小化してトレイ常駐 ---
    def on_close():
        root.withdraw()
        start_tray_icon_once()
        log_queue.put("[INFO] 最小化して常駐しました（トレイから操作可）")
    root.protocol("WM_DELETE_WINDOW", on_close)

    root.mainloop()



if __name__=="__main__":
    if not acquire_single_instance_lock():
        try:
            from win10toast import ToastNotifier
            ToastNotifier().show_toast("起動中", "アプリはすでに実行されています。", duration=5, threaded=True)
        except:
            pass
        sys.exit(0)

        
    run_gui()
