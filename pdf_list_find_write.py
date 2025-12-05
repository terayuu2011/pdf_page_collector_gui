import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import tkinter.font as tkfont
import re, os, json
from datetime import datetime
from cryptography.fernet import Fernet

CONFIG_PATH = "config.json"
FLIGHT_LIST_PATH = "出力便名リスト.txt"
MAX_PASSENGER_COUNT = 20

def get_encryption_key(key_path="status_key.key"):
    """
    暗号化キーを取得。存在しなければ一度だけ生成して保存。
    すでに存在する場合は再利用。
    """
    if os.path.exists(key_path):
        with open(key_path, "rb") as f:
            key = f.read()
    else:
        key = Fernet.generate_key()
        with open(key_path, "wb") as f:
            f.write(key)
    return key


class PDFPassengerSearchApp:
    LINE_WIDTH = 0.8
    LINE_MARGIN = 1.5
    TV_COL_MIN, TV_COL_MAX, TV_COL_PAD = 60, 180, 18

    # 🔧 ステータス描画位置オフセット設定（単位: pt）
    STATUS_OFFSET_X = 170   # ← 予約番号の左側にずらす距離（マイナスで左、プラスで右）
    STATUS_OFFSET_Y = 15    # ↑ 縦方向の調整（マイナスで上、プラスで下）

    def __init__(self, root):
        self.root = root
        self.root.title("乗客名簿検索システム（便名検索＋NS/CXL付与＋JSON復元）")

        # ---------------- 設定ファイル読込 ----------------
        self.pdf_folder = ""
        self.load_config()

        # ---------------- 上部ツールバー ----------------
        toolbar = tk.Frame(root)
        toolbar.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        toolbar.grid_columnconfigure(1, weight=1)

        tk.Label(toolbar, text="便名:").grid(row=0, column=0, sticky="w")

        # 便名リスト読み込み
        #self.flight_list = self.load_flight_list()
        #self.flight_cb = ttk.Combobox(toolbar, values=self.flight_list, width=20, state="readonly")
        #self.flight_cb.grid(row=0, column=1, sticky="w", padx=5)
        #self.last_flight_name = self.flight_cb.get()

        # 便名リスト読み込み
        self.flight_list = self.load_flight_list()

        # ▼ 追加：StringVar を使って常に変更検知
        self.flight_var = tk.StringVar(value=(self.flight_list[0] if self.flight_list else ""))
        self._suspend_flight_trace = False         # 変更巻き戻し時の無限ループ防止
        self.last_flight_name = self.flight_var.get()  # 直前の便名を保持

        # Combobox に textvariable をセット（編集不可にしたい場合は readonly）
        self.flight_cb = ttk.Combobox(toolbar, values=self.flight_list, width=20,
                                    textvariable=self.flight_var, state="readonly")
        self.flight_cb.grid(row=0, column=1, sticky="w", padx=5)

        # ▼ 追加：変更イベント（書き込み前警告）
        self.flight_var.trace_add("write", self._on_flight_var_change)

        #tk.Button(toolbar, text="検索", command=self.search_by_flight_name).grid(row=0, column=2, padx=6)
        tk.Button(toolbar, text="送信（PDFに書き込み）", command=self.write_all_status_to_pdf).grid(row=0, column=3, padx=6)

        # ---------------- Treeview ----------------
        columns = (
            "Status", "№", "予約番号", "氏名", "男", "女", "子供", "合計",
            "電話番号", "乗車地", "下車地", "便名", "旅行期間", "予約サイト", "クラス"
        )
        self.tree = ttk.Treeview(root, columns=columns, show="headings", height=20)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=90, anchor="center")
        self.tree.grid(row=1, column=0, padx=6, pady=4, sticky="nsew")
        self.tree.tag_configure('status_red', foreground='red', font=('Arial', 10, 'bold'))
        self.tree.tag_configure('status_blue', foreground='blue', font=('Arial', 10, 'bold'))
        self.tree.tag_configure('status_cxl_cs', background="#F6CBDA", foreground='red', font=('Arial', 10, 'bold'))

        # ▼▼▼ 追加：固定フッター（Treeview列幅と同期） ▼▼▼
        self.footer_canvas = tk.Canvas(root, height=24, bg="#f4f4f4", highlightthickness=0)
        self.footer_canvas.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 4))

        # 描画テキストIDを保持
        self.footer_texts = {}

        # __init__ の末尾あたりに追加
        self.baseline_snapshot = {}   # 検索直後 or 保存直後の基準
        self.unsaved_changes = False


        def update_footer_position():
            """Treeview列幅を取得してフッターにラベルを再描画"""
            self.footer_canvas.delete("all")

            # Treeviewの列幅を取得
            cols = self.tree["columns"]
            x_offset = 0
            col_positions = []
            for col in cols:
                w = self.tree.column(col, "width")
                col_positions.append((col, x_offset, w))
                x_offset += w

            # 背景塗り（視覚的なバー）
            self.footer_canvas.create_rectangle(0, 0, x_offset, 24, fill="#f4f4f4", outline="#cccccc")

            # 「合計人数」ラベル（男列の左隣に配置）
            male_col_x = next((pos for (c, pos, _) in col_positions if c == "男"), 0)
            self.footer_canvas.create_text(
                male_col_x - 60, 12, text="合計人数", anchor="w", font=("Arial", 10, "bold")
            )

            # 値取得
            total_m = total_f = total_k = total_sum = 0
            for iid in self.tree.get_children(""):
                vals = self.tree.item(iid, "values")
                if len(vals) < 8:
                    continue
                def safe_int(v):
                    if isinstance(v, str) and "→" in v:
                        try:
                            return int(v.split("→")[-1])
                        except:
                            return 0
                    return int(v) if str(v).isdigit() else 0
                total_m += safe_int(vals[4])
                total_f += safe_int(vals[5])
                total_k += safe_int(vals[6])
                total_sum += safe_int(vals[7])

            # 各列の中央に数値を配置
            for col, x, w in col_positions:
                cx = x + w / 2
                if col == "男":
                    self.footer_canvas.create_text(cx, 12, text=str(total_m), font=("Arial", 10, "bold"))
                elif col == "女":
                    self.footer_canvas.create_text(cx, 12, text=str(total_f), font=("Arial", 10, "bold"))
                elif col == "子供":
                    self.footer_canvas.create_text(cx, 12, text=str(total_k), font=("Arial", 10, "bold"))
                elif col == "合計":
                    self.footer_canvas.create_text(cx, 12, text=str(total_sum), font=("Arial", 10, "bold"), fill="#000000")

        # 再描画イベント（列幅変更・ウィンドウサイズ変更時）
        self.tree.bind("<Configure>", lambda e: update_footer_position())
        root.bind("<Configure>", lambda e: update_footer_position())

        # 手動更新用にメソッドとして登録
        self.update_footer_totals = update_footer_position
        # ▲▲▲ 追加ここまで ▲▲▲

        # コンテキストメニュー（右クリック）
        self.menu = tk.Menu(root, tearoff=0)
        self.tree.bind("<Button-3>", self.show_context_menu)

        # ---------------- ログ ----------------
        tk.Label(root, text="ログ:").grid(row=2, column=0, sticky="w", padx=6)
        self.log_text = tk.Text(root, height=8, width=120)
        self.log_text.grid(row=3, column=0, padx=6, pady=4, sticky="ew")

        root.grid_rowconfigure(1, weight=1)
        root.grid_columnconfigure(0, weight=1)

        self.cxl_deduction_map = {}
        self.current_pdf_path = None

        self.log_text.insert(tk.END, f"[設定] PDFフォルダ: {self.pdf_folder}\n")


    def _on_flight_var_change(self, *_):
        """便名が変わった瞬間に呼ばれる（trace）。未書き込みなら確認し、却下なら元に戻す。"""
        if self._suspend_flight_trace:
            return

        new_val = self.flight_var.get().strip()
        old_val = getattr(self, "last_flight_name", "")

        # 変更がないなら何もしない
        if new_val == old_val:
            return

        # 未書き込みフラグが立っているか？（is_dirty は既存実装を使用）
        self._update_dirty_flag()  # ← 先に最新の差分を再計算
        if self.unsaved_changes:
            ans = messagebox.askyesno(
                "確認",
                f"変更内容がまだ書き込まれていません。\n\n"
                f"便名を『{old_val}』→『{new_val}』に切り替えますか？",
                parent=self.root
            )
            if not ans:
                # いいえ → 値を元に戻す（trace 再発火防止のためガード）
                try:
                    self._suspend_flight_trace = True
                    self.flight_var.set(old_val)
                finally:
                    self._suspend_flight_trace = False
                self.log_text.insert(tk.END, "[INFO] 便名変更をキャンセルしました。\n")
                return
            else:
                # はい → 今の未保存変更を破棄して続行
                self.unsaved_changes = False
                self.log_text.insert(tk.END, "[INFO] 便名変更を続行します（未保存内容を破棄）。\n")

        # ここに来たら変更を確定：基準値（last_flight_name）を更新
        self.last_flight_name = new_val
        self.search_by_flight_name()



    def _safe_int_view(self, v):
        # "2→1" 形式は after 側を採用
        if isinstance(v, str) and "→" in v:
            try:
                return int(v.split("→")[-1])
            except:
                return 0
        return int(v) if str(v).isdigit() else 0

    def _make_snapshot_from_tree(self):
        """
        現在のTreeview内容を {resv: {status,male,female,child,total}} で返す。
        ※ resv（予約番号）をキーにするので行順やiidが変わってもOK
        """
        snap = {}
        for iid in self.tree.get_children(""):
            vals = self.tree.item(iid, "values")
            if len(vals) < 8:
                continue
            resv   = vals[2]
            status = vals[0] or ""
            male   = self._safe_int_view(vals[4])
            female = self._safe_int_view(vals[5])
            child  = self._safe_int_view(vals[6])
            total  = self._safe_int_view(vals[7])
            snap[resv] = {"status": status, "male": male, "female": female, "child": child, "total": total}
        return snap

    def _update_dirty_flag(self):
        """
        現在の表示と baseline_snapshot を比較して unsaved_changes を更新。
        '元の状態に戻した' 場合は False になる。
        """
        current = self._make_snapshot_from_tree()
        self.unsaved_changes = (current != getattr(self, "baseline_snapshot", {}))


    # どこでも良いですが class 内のユーティリティ群の近くに
    def _apply_ns_for_item(self, item_id):
        """選択行を NS として '元→0' 表示に更新（内部数値は safe_int で解釈）"""
        values = list(self.tree.item(item_id, "values"))
        if not values or len(values) < 8:
            return
        # いま表示されている数値（素の数値 / '2→1' の後ろ側など）を基準に NS 化
        def aft(x):  # 現表示から after 値を取得（safe_int は既存実装）
            return self.safe_int(x)

        orig_m = aft(values[4])
        orig_f = aft(values[5])
        orig_k = aft(values[6])
        orig_t = aft(values[7]) if str(values[7]).strip() else (orig_m + orig_f + orig_k)

        values[0] = "NS"
        # 表示は「元→0」
        values[4] = f"{orig_m}→0" if orig_m > 0 else "0"
        values[5] = f"{orig_f}→0" if orig_f > 0 else "0"
        values[6] = f"{orig_k}→0" if orig_k > 0 else "0"
        values[7] = f"{orig_t}→0" if orig_t > 0 else "0"

        # NS は色タグだけ（減算情報は保持不要）
        self.tree.item(item_id, values=values, tags=('status_blue',))


    def _safe_int_for_total(self, v: str) -> int:
            """セル表示が '2→1' のような書式でも後値を拾って整数化。非数は 0。"""
            if isinstance(v, str) and "→" in v:
                try:
                    return int(v.split("→")[-1])
                except Exception:
                    return 0
            return int(v) if str(v).isdigit() else 0

    def update_footer_totals(self):
        """Treeview を走査して固定フッターの合計を更新。"""
        total_m = total_f = total_k = total_sum = 0
        for iid in self.tree.get_children(""):
            vals = self.tree.item(iid, "values")
            if len(vals) < 8:
                continue
            total_m   += self._safe_int_for_total(vals[4])
            total_f   += self._safe_int_for_total(vals[5])
            total_k   += self._safe_int_for_total(vals[6])
            total_sum += self._safe_int_for_total(vals[7])

        self.sum_male.config(text=str(total_m))
        self.sum_female.config(text=str(total_f))
        self.sum_child.config(text=str(total_k))
        self.sum_total.config(text=str(total_sum))

    # ---------------- 設定ファイル読込 ----------------
    def load_config(self):
        """pdf_page_collector_gui_full.pyと共通のconfig.jsonを読み込み、output_folderを参照"""
        if os.path.exists(CONFIG_PATH):
            try:
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)

                # pdf_page_collector_gui_fullの構造に対応
                folders = data.get("folders", {})
                self.pdf_folder = folders.get("output_folder", "")

                if not self.pdf_folder:
                    # 旧形式（pdf_folder直下キー）に対応
                    self.pdf_folder = data.get("pdf_folder", "")

                # ✅ log_textが存在する場合のみ出力（初期化前でも安全）
                if hasattr(self, "log_text"):
                    if self.pdf_folder:
                        self.log_text.insert(tk.END, f"[設定読込] PDF出力フォルダ: {self.pdf_folder}\n")
                    else:
                        self.log_text.insert(tk.END, "[WARN] config.jsonにoutput_folderが見つかりません。\n")

            except Exception as e:
                # log_textがまだ存在しない場合はprintで代用
                msg = f"config.jsonの読込に失敗しました: {e}"
                if hasattr(self, "log_text"):
                    self.log_text.insert(tk.END, f"[ERROR] {msg}\n")
                else:
                    print("[設定エラー]", msg)
                messagebox.showerror("設定エラー", msg)
        else:
            warning = "config.jsonが見つかりません。PDFフォルダを設定してください。"
            if hasattr(self, "log_text"):
                self.log_text.insert(tk.END, f"[WARN] {warning}\n")
            else:
                print("[警告]", warning)
            messagebox.showwarning("警告", warning)


    # ---------------- 便名リスト読込 ----------------
    def load_flight_list(self):
        if not os.path.exists(FLIGHT_LIST_PATH):
            messagebox.showwarning("警告", f"便名リストファイル {FLIGHT_LIST_PATH} が見つかりません。", parent=self.root)
            return []
        with open(FLIGHT_LIST_PATH, "r", encoding="utf-8") as f:
            return [line.strip() for line in f if line.strip()]

    # ---------------- Treeview列自動調整 ----------------
    def autosize_tree_columns(self):
        tv = self.tree
        try:
            style = ttk.Style()
            tv_font_name = style.lookup("Treeview", "font") or "TkDefaultFont"
            f = tkfont.nametofont(tv_font_name)
        except Exception:
            f = tkfont.nametofont("TkDefaultFont")

        for col in tv["columns"]:
            max_px = f.measure(col)
            for iid in tv.get_children(""):
                val = tv.set(iid, col)
                max_px = max(max_px, f.measure(str(val)))
            width = max(self.TV_COL_MIN, min(max_px + self.TV_COL_PAD, self.TV_COL_MAX))
            tv.column(col, width=int(width))

    # ---------------- コンテキストメニュー ----------------
    def show_context_menu(self, event):
        row_id = self.tree.identify_row(event.y)
        if not row_id:
            return
        self.tree.selection_set(row_id)
        values = list(self.tree.item(row_id, "values"))
        current_status = values[0] if values else ""

        self.menu.delete(0, tk.END)
        if current_status == "NS":
            self.menu.add_command(label="NSを解除", command=lambda: self.unset_status_for_selected("NS"))
        elif current_status == "CXL":
            self.menu.add_command(label="CXLを解除", command=lambda: self.unset_status_for_selected("CXL"))
        elif current_status == "CXL-CS":
            self.menu.add_command(label="CXL-CSを解除", command=lambda: self.unset_status_for_selected("CXL-CS"))
        else:
            self.menu.add_command(label="NSに設定", command=lambda: self.set_status_for_selected("NS"))
            self.menu.add_command(label="CXLに設定", command=lambda: self.open_cxl_dialog("CXL"))
            self.menu.add_command(label="CXL(CS報告分)に設定", command=lambda: self.open_cxl_dialog("CXL_CS"))

        self.menu.post(event.x_root, event.y_root)

    def set_status_for_selected(self, status):
        selected = self.tree.selection()
        if not selected:
            return
        if status == "NS":
            for item in selected:
                self._apply_ns_for_item(item)
            self.log_text.insert(tk.END, f"[STATUS更新] NS を {len(selected)}件に設定（元→0 表示）\n")
            self._update_dirty_flag()
            try:
                self.update_footer_totals()
            except Exception:
                pass

        # 既存：CXL を選ぶとダイアログ、などの処理
        for item in selected:
            values = list(self.tree.item(item, "values"))
            if not values:
                continue
            values[0] = status
            # ✅ ステータス別にタグを正しく振り分け
            if status == "NS":
                tag = ('status_blue',)
            elif status in ("CXL", "CXL_CS"):
                tag = ('status_red',) if status == "CXL" else ('status_cxl_cs',)
            else:
                tag = ()
                
            self.tree.item(item, values=values, tags=tag)
        self.log_text.insert(tk.END, f"[STATUS更新] {status} を {len(selected)}件に設定\n")


    def clear_status(self, item_id):
        """ステータス解除処理"""
        values = list(self.tree.item(item_id, "values"))
        if not values:
            return
        prev = values[0]
        values[0] = ""
        self.tree.item(item_id, values=values, tags=())
        if item_id in self.cxl_deduction_map:
            del self.cxl_deduction_map[item_id]
        self.log_text.insert(tk.END, f"[解除] {values[3]} の {prev} を解除しました\n")
    
    def unset_status_for_selected(self, status_to_unset: str):
        """
        指定ステータス（NS/CXL）を解除し、
        現在表示中PDFの “該当ページのみ” を予約番号キーで再解析して TreeView を上書き。
        """
        selected = self.tree.selection()
        if not selected:
            return

        pdf_path = getattr(self, "current_pdf_path", None)
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showwarning("警告", "現在表示中のPDFが見つかりません。", parent=self.root)
            return

        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            messagebox.showerror("エラー", f"PDFを開けませんでした:\n{e}", parent=self.root)
            return

        restored = 0

        for item_id in selected:
            values = list(self.tree.item(item_id, "values"))
            if len(values) < 4:
                continue

            # 今その行に付いているステータスが対象以外ならスキップ（安全策）
            curr_status = values[0]
            if curr_status != status_to_unset:
                continue

            resv_no = values[2]               # 予約番号（例: 9J-123456）
            norm_resv = self.normalize_text(resv_no)

            # ✅ TreeView が保持している “ページ番号” を取得（末尾を想定）
            try:
                page_index = int(values[-1])
            except Exception:
                self.log_text.insert(tk.END, f"[WARN] 予約番号 {resv_no}: ページ番号が不正のため解除スキップ\n")
                continue

            if not (0 <= page_index < len(doc)):
                self.log_text.insert(tk.END, f"[WARN] 予約番号 {resv_no}: ページ番号 {page_index+1} が範囲外です\n")
                continue

            page = doc[page_index]
            words = page.get_text("words")

            # --- 該当ページだけで行再構築 ---
            lines_by_y = {}
            for w in words:
                x0, y0, x1, y1, text = w[:5]
                y = round(y0, 1)
                found_y = next((yy for yy in lines_by_y if abs(yy - y) <= 1.5), None)
                if found_y is not None:
                    lines_by_y[found_y].append((x0, text))
                else:
                    lines_by_y[y] = [(x0, text)]

            found_line = None
            for y in sorted(lines_by_y.keys()):
                line_items = sorted(lines_by_y[y], key=lambda x: x[0])
                raw_line = "".join([t for _, t in line_items])
                norm_line = self.normalize_text(raw_line)

                # ▶ 予約番号でヒット判定（normalize済みで比較）
                if norm_resv in norm_line:
                    found_line = norm_line
                    break

            if not found_line:
                self.log_text.insert(tk.END, f"[WARN] 予約番号 {resv_no}: 指定ページ p.{page_index+1} で元行が見つかりません\n")
                continue

            # --- 解析して項目分割 ---
            parsed = self.parse_passenger_line(found_line)
            if not parsed:
                self.log_text.insert(tk.END, f"[WARN] 予約番号 {resv_no}: 行の再解析に失敗しました\n")
                continue

            # --- TreeView を “ステータス空” かつ “同じページ番号” で置き換え
            new_values = [""] + parsed + [page_index]
            self.tree.item(item_id, values=new_values, tags=())

            # CXL 減算データも除去
            if item_id in self.cxl_deduction_map:
                del self.cxl_deduction_map[item_id]

            self.log_text.insert(tk.END, f"[{status_to_unset}解除] 予約番号 {resv_no}: p.{page_index+1} から元行を復元しました\n")
            restored += 1

        doc.close()

        if restored:
            self.autosize_tree_columns()
            self._update_dirty_flag()
            try:
                self.root.after(120, self.update_footer_totals)
            except Exception:
                pass
            


    # ---------------- 安全な数値変換 ----------------
    def safe_int(self, val):
        """'2→1' のような表記があっても後値を整数で返す"""
        if isinstance(val, str):
            if "→" in val:
                val = val.split("→")[-1]
            val = val.strip()
            if val.isdigit():
                return int(val)
        elif isinstance(val, (int, float)):
            return int(val)
        return 0

    # ---------------- CXL人数ダイアログ（2→1対応版） ----------------
    def open_cxl_dialog(self, status):
        selected = self.tree.selection()
        if not selected:
            return
        item_id = selected[0]
        values = list(self.tree.item(item_id, "values"))
        if not values:
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("CXL人数指定")
        dialog.transient(self.root)  # メイン画面の上に固定
        dialog.grab_set()

        tk.Label(dialog, text=f"{values[3]} さん").grid(row=0, column=0, columnspan=6, pady=5)

        tk.Label(dialog, text="男:").grid(row=1, column=0)
        male_cb = ttk.Combobox(dialog, width=4, values=list(range(MAX_PASSENGER_COUNT + 1)))
        male_cb.grid(row=1, column=1, padx=3)

        tk.Label(dialog, text="女:").grid(row=1, column=2)
        female_cb = ttk.Combobox(dialog, width=4, values=list(range(MAX_PASSENGER_COUNT + 1)))
        female_cb.grid(row=1, column=3, padx=3)

        tk.Label(dialog, text="子供:").grid(row=1, column=4)
        child_cb = ttk.Combobox(dialog, width=4, values=list(range(MAX_PASSENGER_COUNT + 1)))
        child_cb.grid(row=1, column=5, padx=3)

        # 現在（＝元）の値（safe_intで確実に取得）
        orig_m = self.safe_int(values[4])
        orig_f = self.safe_int(values[5])
        orig_k = self.safe_int(values[6])
        orig_total = orig_m + orig_f + orig_k

        # 初期値セット
        male_cb.set(str(orig_m))
        female_cb.set(str(orig_f))
        child_cb.set(str(orig_k))

        dialog.update_idletasks()

        # 親ウィンドウ位置・サイズ取得
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_w = self.root.winfo_width()
        root_h = self.root.winfo_height()

        # ダイアログサイズ
        w = dialog.winfo_width()
        h = dialog.winfo_height()

        # 中央座標計算（親ウィンドウ中央基準）
        x = root_x + (root_w // 2) - (w // 2)
        y = root_y + (root_h // 2) - (h // 2)
        dialog.geometry(f"+{x}+{y}")


        def fmt(before, after):
            return f"{before}→{after}" if before != after else str(after)

        def apply_cxl():
            male = int(male_cb.get() or 0)
            female = int(female_cb.get() or 0)
            child = int(child_cb.get() or 0)
            total = male + female + child

            # --- 元の値（orig）は「→」の左側、なければ数値部分 ---
            def parse_orig(token):
                s = str(token)
                if "→" in s:
                    head = s.split("→")[0]
                    return int(head) if head.isdigit() else 0
                return int(s) if s.isdigit() else 0

            orig_m = parse_orig(values[4])
            orig_f = parse_orig(values[5])
            orig_k = parse_orig(values[6])
            orig_t = parse_orig(values[7])

            # --- CXL減算情報を登録（元値→変更後）---
            self.cxl_deduction_map[item_id] = {
                "orig": {"男": orig_m, "女": orig_f, "子供": orig_k, "合計": orig_t},
                "after": {"男": male, "女": female, "子供": child, "合計": total}
            }

            #values[0] = "CXL"
            values[0] = "CXL-CS" if status == "CXL_CS" else "CXL"

            # ✅ 減算なし（＝全て同値）または after が全0 の場合 → 「→0」表記
            if (orig_m == male and orig_f == female and orig_k == child) or (male + female + child == 0):
                def fmt_ns(before):
                    return f"{before}→0" if int(before) > 0 else str(before)
                values[4] = fmt_ns(orig_m)
                values[5] = fmt_ns(orig_f)
                values[6] = fmt_ns(orig_k)
                values[7] = fmt_ns(orig_t)
            else:
                # ✅ 減算あり：通常「元→後」
                def fmt(before, after):
                    return f"{before}→{after}" if before != after else str(after)
                values[4] = fmt(orig_m, male)
                values[5] = fmt(orig_f, female)
                values[6] = fmt(orig_k, child)
                values[7] = fmt(orig_t, total)

            # TreeView更新
            #self.tree.item(item_id, values=values, tags=('status_red',))
            if status == "CXL_CS":
                self.tree.item(item_id, values=values, tags=('status_cxl_cs',))
            else:
                self.tree.item(item_id, values=values, tags=('status_red',))
                
            self.log_text.insert(
                tk.END,
                f"[CXL設定] {values[3]} 男:{values[4]} 女:{values[5]} 子:{values[6]} 合:{values[7]}\n"
            )

            # 更新フラグ・フッター更新
            self._update_dirty_flag()
            try:
                self.update_footer_totals()
            except Exception:
                pass

            dialog.destroy()


        tk.Button(dialog, text="確定", command=apply_cxl).grid(row=3, column=0, columnspan=6, pady=10)


    def normalize_text(self, txt: str):
        """全角数字・カタカナ・記号などを半角・正規形に整える"""
        import re
        import unicodedata

        if not txt:
            return ""

        # Unicode正規化（濁点付き文字や異体字を統一）
        txt = unicodedata.normalize("NFKC", txt)

        # よくある記号ゆらぎの統一
        txt = txt.replace("―", "-").replace("ー", "-").replace("−", "-")
        txt = txt.replace("⇒", "→").replace("＞", ">").replace("＜", "<")

        # 全角数字・記号を半角へ
        z2h = str.maketrans({
            "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
            "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
            "　": " ", "：": ":", "．": ".", "／": "/", "～": "-",
        })
        txt = txt.translate(z2h)

        # 複数文字の変換（maketransでは不可）
        txt = txt.replace("號車", "号車")

        # スペース・改行・タブ除去
        txt = re.sub(r"\s+", "", txt)

        return txt

     # ---------------- データ検索 ----------------
    def parse_passenger_line(self, line_text: str):
        """
        号車別明細表の1行から各項目を抽出。
        電話番号が「0」以外で始まる(例: 336-5266-7188, 15089178424)ケースにも対応。
        """
        import re

        def normalize_phone(raw: str) -> str:
            """電話番号を統一フォーマットに整形"""
            d = re.sub(r'\D', '', raw)
            if len(d) == 11:
                return f"{d[:3]}-{d[3:7]}-{d[7:]}"
            elif len(d) == 10:
                return f"{d[:2]}-{d[2:6]}-{d[6:]}"
            else:
                return raw

        s = re.sub(r"\s+", "", line_text)
        s = s.replace("⇒", "→").replace("―", "-").replace("ー", "-").replace("−", "-")

        # 1️⃣ No(1〜2桁 任意) + 予約番号（数字1 + 英字1〜2 + '-' + 数字4+）
        m_head = re.match(r'^(?:(?P<no>\d{1,2}))?(?P<resv>\d[A-Z]{1,2}-\d{4,})', s)
        if not m_head:
            return []
        no = m_head.group("no") or ""
        resv = m_head.group("resv")
        idx = m_head.end()

        # 2️⃣ 氏名 → 人数4桁（男女子計）
        m_cnt = re.search(r'(\d)(\d)(\d)(\d)', s[idx:])
        if not m_cnt:
            return []
        name = s[idx: idx + m_cnt.start()]
        male, female, child, total = m_cnt.groups()
        idx += m_cnt.end()

        # 3️⃣ 電話番号抽出（拡張版）
        tel = ""
        # ハイフン付き、もしくは11桁数字、または3〜4桁始まり
        phone_patterns = [
            r'(?:0\d{1,4}|[1-9]\d{1,3})-\d{2,4}-\d{3,4}',  # ハイフン付き (例: 336-5266-7188, 03-1234-5678)
            r'(?:0\d{9,10}|[1-9]\d{8,10})'                # ハイフンなし (例: 15089178424)
        ]
        phone_match = None
        for p in phone_patterns:
            m = re.search(p, s[idx:])
            if m:
                phone_match = m
                break

        if phone_match:
            tel = normalize_phone(phone_match.group())
            # 検出した電話部分を削除
            start, end = idx + phone_match.start(), idx + phone_match.end()
            s = s[:start] + s[end:]

        # 4️⃣ 乗車地 → 下車地 + 便名
        pickup, dropoff, flight = "", "", ""
        m_route = re.search(r'([^→]+)→([^→]+?)(\d{1,3}便)', s[idx:])
        if m_route:
            pickup, dropoff, flight = m_route.group(1), m_route.group(2), m_route.group(3)
            idx = idx + m_route.end()
        else:
            # 便名だけある場合
            m_flight = re.search(r'(\d{1,3}便)', s[idx:])
            if m_flight:
                flight = m_flight.group(1)
                before = s[idx: idx + m_flight.start()]
                m_route2 = re.search(r'([^→]+)→([^→]+)', before)
                if m_route2:
                    pickup, dropoff = m_route2.group(1), m_route2.group(2)
                idx = idx + m_flight.end()

        # 5️⃣ 旅行期間
        period = ""
        m_period = re.search(r'\d{2}/\d{2}/\d{2}-\d{2}/\d{2}', s[idx:])
        if m_period:
            period = m_period.group(0)
            idx += m_period.end()

        # 6️⃣ サイト + クラス
        site, bus_class = "", ""
        rest = s[idx:]
        known_sites = [
            "ｼﾞｬﾑｼﾞｬﾑﾋｶｸ", "ｼﾞｬﾑｼﾞｬﾑﾗｲﾅｰ",
            "ジャムジャムヒカク", "ジャムジャムライナー",
            "WILLER", "ﾗｸﾃﾝ", "ラクテン"
        ]
        for st in known_sites:
            if st in rest:
                site = st
                after = rest.split(st, 1)[1]
                m_cls = re.search(r'([0-9I][0-9])$', after)
                if m_cls:
                    bus_class = m_cls.group(1)
                break
        if not site:
            m_cls = re.search(r'([0-9I][0-9])$', rest)
            bus_class = m_cls.group(1) if m_cls else ""

        return [
            no, resv, name, male, female, child, total, tel,
            pickup, dropoff, flight, period, site, bus_class
        ]



    def search_by_flight_name(self):
        # search_by_flight_name() の先頭に追加（便名チェックの前でOK）
        #self._update_dirty_flag()  # ← 先に最新の差分を再計算
        #if self.unsaved_changes:
            #proceed = messagebox.askyesno("確認", "未保存の変更があります。\n保存せずに便名を再検索しますか？", parent=self.root)
            #if not proceed:
            #    self.log_text.insert(tk.END, "[CANCEL] 便名検索を中止（未保存の変更あり）\n")
            #    return


        """コンボボックスの便名（例：262号車）を基に、PDF内で該当便の行を抽出して分割表示"""
        flight_keyword = self.flight_cb.get().strip()
        if not flight_keyword:
            messagebox.showwarning("警告", "便名（号車）を選択してください。", parent=self.root)
            return

        # 「号車」→「便」に変換
        normalized_flight = self.normalize_text(flight_keyword)
        normalized_flight = re.sub(r"号車$", "便", normalized_flight)

        self.tree.delete(*self.tree.get_children())
        self.log_text.insert(tk.END, f"\n--- [便名検索] {normalized_flight} ---\n")

        # ✅ 「保管用」を含むPDFのみ対象、かつ _marked.pdf は除外
        candidate_pdfs = [
            os.path.join(self.pdf_folder, f)
            for f in os.listdir(self.pdf_folder)
            if f.lower().endswith(".pdf")
            and "保管用" in f
            and "_marked" not in f.lower()  # ← ★ 追加行：_marked.pdf除外
        ]

        total_hits = 0
        matched_pdf = None  # ✅ 一致したPDFを記録して後で使用

        for pdf_path in candidate_pdfs:
            try:
                doc = fitz.open(pdf_path)
            except Exception as e:
                self.log_text.insert(tk.END, f"[WARN] {pdf_path} を開けません: {e}\n")
                continue

            for page_index, page in enumerate(doc):
                words = page.get_text("words")

                # y座標で行を再構築
                lines_by_y = {}
                for w in words:
                    x0, y0, x1, y1, text = w[:5]
                    y = round(y0, 1)
                    found_y = next((yy for yy in lines_by_y if abs(yy - y) <= 1.5), None)
                    if found_y is not None:
                        lines_by_y[found_y].append((x0, text))
                    else:
                        lines_by_y[y] = [(x0, text)]

                # 各行を解析
                for y in sorted(lines_by_y.keys()):
                    line_items = sorted(lines_by_y[y], key=lambda x: x[0])
                    line_text = "".join([t for _, t in line_items])
                    norm_line = self.normalize_text(line_text)

                    # 「262便」などが含まれる行を抽出
                    if normalized_flight not in norm_line:
                        continue

                    # 予約番号（9J-xxxxxxなど）を含む行のみ採用
                    if not re.search(r"[A-Z0-9]{1,5}-[0-9]{3,}", norm_line):
                        continue

                    # 🔹既存の解析関数を呼び出し
                    parsed = self.parse_passenger_line(norm_line)

                    if parsed and len(parsed) >= 3:
                        # 🔹 ステータス自動判定（NS/CXL）
                        status = ""
                        if re.search(r"NS(?![A-Za-z0-9])", norm_line):
                            status = "NS"
                        elif re.search(r"CXL(?![A-Za-z0-9])", norm_line):
                            status = "CXL"

                        # 🔹 Treeview に追加
                        self.tree.insert("", "end", values=[status, *parsed, page_index])
                        total_hits += 1
                        matched_pdf = pdf_path
                        self.log_text.insert(tk.END, f"[抽出] p.{page_index+1}: {norm_line[:80]}...\n")

            doc.close()

        # ✅ 抽出結果を記録（ここで current_pdf_path にセット）
        if matched_pdf:
            self.current_pdf_path = matched_pdf
            self.log_text.insert(tk.END, f"[INFO] 対象PDFを設定: {os.path.basename(matched_pdf)}\n")
        else:
            self.current_pdf_path = None

        # ✅ 結果出力
        if total_hits == 0:
            messagebox.showinfo("結果", f"{normalized_flight} の便に該当する行は見つかりませんでした。", parent=self.root)
            self.log_text.insert(tk.END, "[INFO] 条件に合う行なし。\n")
        else:
            self.log_text.insert(tk.END, f"[完了] {total_hits} 行を抽出しました。\n")

        # ===== JSONステータスファイルの読み込み =====
        base_status_folder = os.path.join(self.pdf_folder, "status_data")

        # --- PDFファイル名から日付フォルダー名を生成 ---
        pdf_name = os.path.basename(self.current_pdf_path) if getattr(self, "current_pdf_path", None) else ""
        m = re.search(r"(\d{1,2})[.\-](\d{1,2})", pdf_name)
        if m:
            month, day = m.groups()
            folder_name = f"{datetime.now().year}-{month.zfill(2)}-{day.zfill(2)}"
            status_folder = os.path.join(base_status_folder, folder_name)
        else:
            status_folder = base_status_folder

        # --- 安全な存在チェック ---
        if not os.path.exists(status_folder):
            self.log_text.insert(tk.END, f"[INFO] ステータスフォルダーが存在しません: {status_folder}\n")
            self.unsaved_changes = False  # ← 強制的に変更なし
            self.baseline_snapshot = self._make_snapshot_from_tree()  # 現在を基準に
            self.log_text.insert(tk.END, "[INIT] 初期基準確定（フォルダー未生成）\n")
        else:
            # --- JSON検索処理 ---
            normalized_json_prefix = re.sub(r"号車$", "便", normalized_flight)
            json_candidates = [
                f for f in os.listdir(status_folder)
                if f.startswith(normalized_json_prefix) and f.endswith("_status.json")
            ]

            if not json_candidates:
                self.log_text.insert(tk.END, "[INFO] 該当するステータスJSONが見つかりません。\n")
                self.unsaved_changes = False
                self.baseline_snapshot = self._make_snapshot_from_tree()
                self.log_text.insert(tk.END, "[INIT] 初期基準確定（JSONなし）\n")
            else:
                json_candidates.sort(
                    key=lambda f: os.path.getmtime(os.path.join(status_folder, f)),
                    reverse=True
                )
                latest_json = json_candidates[0]
                json_path = os.path.join(status_folder, latest_json)

            try:
                #with open(json_path, "r", encoding="utf-8") as f:
                #    data = json.load(f)
                key = get_encryption_key()  # ← 自動生成＋永続再利用
                fernet = Fernet(key)

                with open(json_path, "rb") as f:
                    enc = f.read()

                # 復号してからJSONとして読込
                try:
                    dec = fernet.decrypt(enc)
                    data = json.loads(dec.decode("utf-8"))
                except Exception as e:
                    self.log_text.insert(tk.END, f"[WARN] ステータスJSONの復号に失敗: {e}\n")
                    return

                restored_count = 0
                for record in data.get("records", []):
                    name = record.get("name", "")
                    status = record.get("status", "")
                    male = record.get("male", "")
                    female = record.get("female", "")
                    child = record.get("child", "")
                    total = record.get("total", "")
                    cxl_deduction = record.get("cxl_deduction", {})

                    for item_id in self.tree.get_children():
                        values = list(self.tree.item(item_id, "values"))
                        if len(values) > 3 and values[3] == name:
                            values[0] = status

                            # ✅ CXL処理：減算あり or なしを判定
                            if status in ("CXL", "CXL-CS") and isinstance(cxl_deduction, dict):
                                orig = cxl_deduction.get("orig", {})
                                after = cxl_deduction.get("after", {})

                                # 🔹 各列ごとに個別比較して、変化があるときだけ before→after 表示
                                def fmt_each(before, after):
                                    """変化がある場合のみ before→after、同じ値なら after のみ"""
                                    try:
                                        b = int(before)
                                        a = int(after)
                                        if b != a:
                                            return f"{b}→{a}"
                                        else:
                                            return str(a)
                                    except Exception:
                                        if before != after:
                                            return f"{before}→{after}"
                                        else:
                                            return str(after)

                                values[4] = fmt_each(orig.get("男", ""), after.get("男", ""))
                                values[5] = fmt_each(orig.get("女", ""), after.get("女", ""))
                                values[6] = fmt_each(orig.get("子供", ""), after.get("子供", ""))
                                values[7] = fmt_each(orig.get("合計", ""), after.get("合計", ""))

                                if status == "CXL-CS":
                                    self.tree.item(item_id, values=values, tags=('status_cxl_cs',))
                                else:
                                    self.tree.item(item_id, values=values, tags=('status_red',))

                                self.cxl_deduction_map[item_id] = cxl_deduction

                            # ✅ NS表示：「元→0」
                            elif status == "NS":
                                # 現在の after 値を元として NS 表示へ
                                def aft(x): return self.safe_int(x)
                                om, of_, ok = aft(values[4]), aft(values[5]), aft(values[6])
                                ot = aft(values[7]) if str(values[7]).strip() else (om + of_ + ok)
                                values[4] = f"{om}→0" if om > 0 else "0"
                                values[5] = f"{of_}→0" if of_ > 0 else "0"
                                values[6] = f"{ok}→0" if ok > 0 else "0"
                                values[7] = f"{ot}→0" if ot > 0 else "0"
                                #self.tree.item(item_id, tags=('status_red',))
                                self.tree.item(item_id, tags=('status_blue',))
                            else:
                                self.tree.item(item_id, tags=())

                            self.tree.item(item_id, values=values)
                            restored_count += 1
                            break


                self.log_text.insert(
                    tk.END,
                    f"[JSON読込] {latest_json} から {restored_count} 件の状態を復元しました。\n"
                )


            except Exception as e:
                self.log_text.insert(tk.END, f"[WARN] JSON読込エラー: {e}\n")
        
        # 検索で表示を作り終えた時点を“基準”とする
        self.baseline_snapshot = self._make_snapshot_from_tree()
        self.unsaved_changes = False


        self.autosize_tree_columns()
        self.root.after(120, self.update_footer_totals)


    def add_status_to_pdf_resv(self, page, resv, name, status, log_widget, page_index, fontsize, x_offset=None, y_offset=None):
        """
        予約番号をキーに検索し、その予約番号の左側に NS/CXL を描画。
        文字列の中心が基準位置に来るように調整。
        """
        import fitz

        if not resv:
            return False

        # オフセット設定（外部定数 or デフォルト）
        if x_offset is None:
            x_offset = getattr(self, "STATUS_OFFSET_X", -25)
        if y_offset is None:
            y_offset = getattr(self, "STATUS_OFFSET_Y", -2)

        words = page.get_text("words")
        added = False

        for w in words:
            text = w[4]
            if resv in text:
                # --- 対象予約番号ワード座標取得 ---
                x0, y0, x1, y1 = w[:4]
                y_center = (y0 + y1) / 2

                # --- ステータス文字列の表示幅を算出 ---
                # PyMuPDFのフォントメトリクスを利用
                try:
                    font = fitz.Font("MyArial")
                except Exception:
                    font = fitz.Font("helv")

                text_width = font.text_length(status, fontsize=fontsize)
                text_height = fontsize * 0.4

                # --- 描画位置を調整（文字中心を基準） ---
                x_target = x0 + x_offset - (text_width / 2)
                y_target = y_center + y_offset - (text_height / 2)

                try:
                    page.insert_font(fontfile=r"C:\Windows\Fonts\arial.ttf", fontname="MyArial")
                except Exception:
                    pass

                # --- 描画実行 ---
                page.insert_text(
                    fitz.Point(x_target, y_target),
                    status,
                    fontsize=fontsize,
                    color=(1, 0, 0),
                    fontname="MyArial",
                    overlay=True
                )

                log_widget.insert(
                    tk.END,
                    f"[PDF追記] '{resv}' 左に {status} (中心基準) "
                    f"(x={x_target:.1f}, y={y_target:.1f}, w={text_width:.1f}, offset=({x_offset},{y_offset})) p.{page_index+1}\n"
                )

                added = True
                break

        if not added:
            log_widget.insert(
                tk.END,
                f"[WARN] 予約番号 '{resv}' が p.{page_index+1} に見つからず（印字スキップ）\n"
            )

        return added


    # ---------------- PDF書き込み（2→1対応safe_int統合版） ----------------
    def write_all_status_to_pdf(self):
        """画面表示PDFは常にベースファイル。
        書き込みは既存 _marked.pdf に追記。
        ステータス解除時は該当ページのみ元PDFから再描画。
        JSONは上書き更新。
        """
        import shutil
        from collections import defaultdict

        base_pdf = getattr(self, "current_pdf_path", None)
        if not base_pdf or not os.path.exists(base_pdf):
            messagebox.showwarning("警告", "現在表示中のPDFが見つかりません。", parent=self.root)
            return

        # ✅ 書き込み対象は常に既存 _marked.pdf（なければ元から生成）
        marked_pdf = base_pdf.replace(".pdf", "_marked.pdf")
        if not os.path.exists(marked_pdf):
            shutil.copyfile(base_pdf, marked_pdf)
            self.log_text.insert(tk.END, f"[INFO] 新規 _marked.pdf 作成: {os.path.basename(marked_pdf)}\n")
        else:
            self.log_text.insert(tk.END, f"[INFO] 既存 _marked.pdf に追記します: {os.path.basename(marked_pdf)}\n")

        # ✅ 便名取得
        if self.tree.get_children():
            first_row = self.tree.item(self.tree.get_children()[0], "values")
            flight_name = first_row[11] if len(first_row) > 11 else "Unknown便"
        else:
            flight_name = "Unknown便"

        # ✅ TreeViewからターゲット抽出
        targets = []
        for item_id in self.tree.get_children():
            vals = self.tree.item(item_id, "values")
            if len(vals) < 9:
                continue
            status = vals[0]
            name = vals[3]
            resv = vals[2]
            try:
                page_index = int(vals[-1])
            except Exception:
                continue
            targets.append((item_id, status, resv, name, page_index, vals))

        if not targets:
            messagebox.showinfo("情報", "現在の便にデータがありません。", parent=self.root)
            return

        # === PDF開く ===
        try:
            doc_marked = fitz.open(marked_pdf)
            doc_base = fitz.open(base_pdf)
        except Exception as e:
            messagebox.showerror("エラー", f"PDFを開けませんでした:\n{e}", parent=self.root)
            return

        # === 対象便ページ特定 ===
        target_pages = sorted(set(p for (_, _, _, _, p, _) in targets))
        self.log_text.insert(tk.END, f"[INFO] 対象ページ: {target_pages}\n")

        # === ステータス解除（空欄）のページを元PDFからリセット ===
        reset_pages = [
            int(vals[-1]) for (_, status, _, _, _, vals) in targets if not status
        ]
        if reset_pages:
            for pno in sorted(set(reset_pages)):
                if pno < len(doc_base) and pno < len(doc_marked):
                    doc_marked.delete_page(pno)
                    doc_marked.insert_pdf(doc_base, from_page=pno, to_page=pno, start_at=pno)
                    self.log_text.insert(tk.END, f"[RESET] p.{pno+1} を元PDFから再描画（解除処理）\n")

        # === ステータス付きデータをページ別に分類 ===
        by_page = defaultdict(list)
        for item in targets:
            if item[1] in ("NS", "CXL", "CXL-CS"):
                by_page[item[4]].append(item)

        # === ステータス付き書き込み ===
        self.log_text.insert(tk.END, "\n--- PDF書き込み開始（追記処理） ---\n")

        def ensure_font(page):
            try:
                page.insert_font(fontname="MyArial", fontfile=r"C:\Windows\Fonts\arial.ttf")
                return "MyArial"
            except Exception:
                return "helv"
            
        if not by_page:
            self.log_text.insert(tk.END, "[INFO] NS/CXLなし。合計人数チェックのみ実行。\n")
            last_page_index = targets[-1][4] if targets else 0
            by_page[last_page_index] = []  # ダミー行を入れて処理ループを起動

        for page_index, rows in sorted(by_page.items()):
            if not (0 <= page_index < len(doc_marked)):
                continue
            page = doc_marked[page_index]
            fontname = ensure_font(page)
            page_ns_sum = 0
            page_cxl_ded_sum = 0

            for (item_id, status, resv, name, _, vals) in rows:
                men = int(vals[4]) if str(vals[4]).isdigit() else 0
                women = int(vals[5]) if str(vals[5]).isdigit() else 0
                kids = int(vals[6]) if str(vals[6]).isdigit() else 0
                total = int(vals[7]) if str(vals[7]).isdigit() else 0

                if status == "NS":
                    page_ns_sum += total
                elif status in ("CXL", "CXL-CS"):
                    cxl = self.cxl_deduction_map.get(item_id, {})
                    for k in ("男", "女", "子供"):
                        v = cxl.get(k, 0)
                        if str(v).isdigit():
                            page_cxl_ded_sum += int(v)

                # ✅ 予約番号左にステータス印字（位置補正あり）
                self.add_status_to_pdf_resv(
                    page, resv, name, status, self.log_text,
                    page_index, fontsize=20
                )

                # ✅ 人数欄の取り消し線＆CXL減算処理
                words = page.get_text("words")
                if not words:
                    continue

                resv_words = [w for w in words if resv in w[4]]
                if not resv_words:
                    self.log_text.insert(tk.END, f"[WARN] 予約番号 '{resv}' が見つかりません。\n")
                    continue

                resv_end_x = resv_words[-1][2]
                line_y = (resv_words[-1][1] + resv_words[-1][3]) / 2

                line_numbers = [
                    w for w in words
                    if w[0] > resv_end_x + 2
                    and abs(((w[1] + w[3]) / 2) - line_y) < 6
                    and w[4].strip().isdigit()
                ]
                line_numbers.sort(key=lambda w: w[0])

                seq = ["男", "女", "子供", "合計"]

                # Treeviewには減算後値が入っている
                def to_int(x): return int(x) if str(x).isdigit() else 0

                # 減算後の値
                after_m = to_int(vals[4])
                after_f = to_int(vals[5])
                after_k = to_int(vals[6])
                after_total = to_int(vals[7])

                # 元の人数を cxl_deduction_map に保持している場合はそれを利用、
                # 無ければ減算前データを別途保持（ここでは同じと仮定）
                orig = self.cxl_deduction_map.get(item_id, {})

                # 元の値は after + deduction（Treeviewが減算後なので逆算）
                orig_m = after_m + to_int(orig.get("男", 0))
                orig_f = after_f + to_int(orig.get("女", 0))
                orig_k = after_k + to_int(orig.get("子供", 0))
                orig_total = after_total + to_int(orig.get("合計", 0))

                seq = ["男", "女", "子供", "合計"]
                cxl_info = self.cxl_deduction_map.get(item_id, {}) if status == "CXL" else {}
                orig_map = {}
                after_map = {}

                # PDF上の元数値をキー毎に読む（全角対策）
                for i, key in enumerate(seq):
                    if i >= len(line_numbers):
                        continue
                    wnum = line_numbers[i]
                    tok = re.sub(r"\D", "", wnum[4].strip())
                    orig_map[key] = int(tok) if tok.isdigit() else 0

                # CXLの「減算後値」を決める
                if status in ("CXL", "CXL-CS"):
                    if isinstance(cxl_info, dict) and "after" in cxl_info:
                        # すでに after / orig を保持している形式に対応
                        after_map["男"] = int(cxl_info["after"].get("男", orig_map.get("男", 0)))
                        after_map["女"] = int(cxl_info["after"].get("女", orig_map.get("女", 0)))
                        after_map["子供"] = int(cxl_info["after"].get("子供", orig_map.get("子供", 0)))
                        # 合計は再計算（安全）
                        after_map["合計"] = after_map["男"] + after_map["女"] + after_map["子供"]
                    else:
                        # TreeViewの値（vals[4:7]）は“減算後値”として使う前提
                        def tv_int(idx, default):
                            v = vals[idx]
                            return int(v) if str(v).isdigit() else default
                        after_map["男"]   = tv_int(4, orig_map.get("男", 0))
                        after_map["女"]   = tv_int(5, orig_map.get("女", 0))
                        after_map["子供"] = tv_int(6, orig_map.get("子供", 0))
                        after_map["合計"] = after_map["男"] + after_map["女"] + after_map["子供"]

                # どれか1つでも減算があるか？
                any_reduced = False
                if status in ("CXL", "CXL-CS"):
                    for k in ("男", "女", "子供", "合計"):
                        if k in orig_map and k in after_map and after_map[k] < orig_map[k]:
                            any_reduced = True
                            break

                for i, key in enumerate(seq):
                    if i >= len(line_numbers):
                        continue

                    wnum = line_numbers[i]
                    x0, x1 = wnum[0] - self.LINE_MARGIN, wnum[2] + self.LINE_MARGIN
                    y_mid = (wnum[1] + wnum[3]) / 2

                    orig_val = orig_map.get(key, 0)
                    # 元が0なら全てスキップ
                    if orig_val == 0:
                        continue

                    if status == "NS":
                        # NSは常に線のみ
                        page.draw_line(p1=(x0, y_mid), p2=(x1, y_mid),
                                    color=(1, 0, 0), width=self.LINE_WIDTH)
                        continue

                    if status in ("CXL", "CXL-CS"):
                        after_val = after_map.get(key, orig_val)

                        if any_reduced:
                            # ✅ 減算ありの列のみ：線＋減算後数値（0でも描画）
                            if after_val < orig_val:
                                page.draw_line(
                                    p1=(x0, y_mid),
                                    p2=(x1, y_mid),
                                    color=(1, 0, 0),
                                    width=self.LINE_WIDTH
                                )
                                # 減算後値は 0 でも必ず描画
                                page.insert_text(
                                    (x0 - self.LINE_MARGIN * 2, y_mid - 4),
                                    str(after_val),
                                    fontsize=10,
                                    color=(1, 0, 0),
                                    fontname=fontname,
                                    overlay=True
                                )
                            # 減算なし列は描画しない
                        else:
                            # ✅ CXL全列変更なし → 線のみ
                            page.draw_line(
                                p1=(x0, y_mid),
                                p2=(x1, y_mid),
                                color=(1, 0, 0),
                                width=self.LINE_WIDTH
                            )


            # === ✅ 各便（by_page単位）の最終ページで「合計人数」行を処理 ===
            # === ✅ 合計人数（GUIの最終行）を使ってPDFに反映し、JSONにも保存 ===
            # --- GUIフッターから合計人数を取得（TreeViewではなくfooter_canvasで計算） ---
            try:
                total_m = total_f = total_k = total_sum = 0
                for iid in self.tree.get_children(""):
                    vals = self.tree.item(iid, "values")
                    if len(vals) < 8:
                        continue
                    def safe_int(v):
                        # 「2→1」形式の場合は後ろ側（after値）を使用
                        if isinstance(v, str) and "→" in v:
                            try:
                                return int(v.split("→")[-1])
                            except:
                                return 0
                        return int(v) if str(v).isdigit() else 0
                    total_m += safe_int(vals[4])
                    total_f += safe_int(vals[5])
                    total_k += safe_int(vals[6])
                    total_sum += safe_int(vals[7])

                self.log_text.insert(
                    tk.END,
                    f"[INFO] フッター合計取得: 男={total_m}, 女={total_f}, 子供={total_k}, 合計={total_sum}\n"
                )

                # --- JSON用 orig/after 構造 ---
                total_record = {
                    "resv": "合計人数",
                    "name": "",
                    "status": "合計",
                    "orig": {"男": 0, "女": 0, "子供": 0, "合計": 0},
                    "after": {"男": total_m, "女": total_f, "子供": total_k, "合計": total_sum},
                }

                # === 最終ページ特定 ===
                last_page_index = max(sorted(set(p for (_, _, _, _, p, _) in targets)))
                if last_page_index >= len(doc_marked):
                    raise RuntimeError("最終ページ番号が不正")

                page = doc_marked[last_page_index]
                fontname = ensure_font(page)
                words = page.get_text("words") or []

                # === 「合計人数」行をPDFから検索 ===
                lines_by_y = {}
                for w in words:
                    x0, y0, x1, y1, text = w[:5]
                    yk = round(y0, 1)
                    for yy in lines_by_y.keys():
                        if abs(yy - yk) <= 1.5:
                            yk = yy
                            break
                    lines_by_y.setdefault(yk, []).append((x0, y0, y1, text))

                target_line = None
                for yy, items in sorted(lines_by_y.items()):
                    line_text = "".join(t[3] for t in sorted(items, key=lambda t: t[0]))
                    if "合計人数" in line_text.replace(" ", ""):
                        target_line = sorted(items, key=lambda t: t[0])
                        break

                if not target_line:
                    self.log_text.insert(tk.END, "[INFO] PDF内に『合計人数』行が見つかりません。\n")
                else:
                    # --- 数値トークン抽出 ---
                    seen_label = False
                    num_tokens = []
                    for (x0, y0, y1, text) in target_line:
                        if "合計人数" in text.replace(" ", ""):
                            seen_label = True
                            continue
                        if seen_label and re.fullmatch(r"\d+", text.strip()):
                            num_tokens.append((x0, y0, y1, text))

                    if len(num_tokens) >= 4:
                        seq = ["男", "女", "子供", "合計"]
                        after_vals = [total_m, total_f, total_k, total_sum]

                        for i, (x0, y0, y1, text) in enumerate(num_tokens[:4]):
                            y_mid = (y0 + y1) / 2
                            x_left = x0 - self.LINE_MARGIN
                            x_right = x0 + len(text) * 5

                            # --- PDF上の元値を取得（全角→半角変換） ---
                            try:
                                orig_val = int(re.sub(r"\D", "", text))
                            except Exception:
                                orig_val = None

                            after_val = after_vals[i]

                            # ✅ 元値と同じならスキップ（線も描画しない）
                            if orig_val is not None and orig_val == after_val:
                                continue

                            # --- 取り消し線 ---
                            page.draw_line(
                                p1=(x_left, y_mid),
                                p2=(x_right, y_mid),
                                color=(1, 0, 0),
                                width=self.LINE_WIDTH
                            )

                            # --- 変更後値を描画（赤文字） ---
                            page.insert_text(
                                (x_right + 6, y_mid - 4),
                                str(after_val),
                                fontsize=10,
                                color=(1, 0, 0),
                                fontname=fontname,
                                overlay=True
                            )

                # === ★追加：この位置（forループの外）に配置 ===
                try:
                    self.log_text.insert(
                        tk.END, f"[DEBUG] ○判定: p.{last_page_index+1}\n"
                    )

                    same_flags = []
                    for i in range(min(4, len(num_tokens))):
                        orig_text = num_tokens[i][3]
                        orig_num = re.sub(r"\D", "", orig_text)
                        after_val = after_vals[i]
                        same = (str(after_val) == orig_num)
                        same_flags.append(same)
                        self.log_text.insert(
                            tk.END,
                            f"[DEBUG]  列={seq[i]} orig='{orig_text}'({orig_num}) → after={after_val} same={same}\n"
                        )

                    if all(same_flags):
                        x0, y0, y1, text = num_tokens[3]
                        cx = (x0 + x0 + len(text) * 5) / 2
                        cy = (y0 + y1) / 2
                        radius = max(6, (len(text) * 3))
                        page.draw_circle(
                            center=(cx, cy),
                            radius=radius,
                            color=(1, 0, 0),
                            width=1.2,
                            overlay=True
                        )
                        self.log_text.insert(
                            tk.END,
                            f"[○] p.{last_page_index+1} 合計人数を○で囲み（人数変更なし）\n"
                        )
                    else:
                        self.log_text.insert(
                            tk.END,
                            f"[DEBUG] ○条件未達: same_flags={same_flags}\n"
                        )

                except Exception as e:
                    self.log_text.insert(
                        tk.END,
                        f"[WARN] ○描画処理中エラー: {e}\n"
                    )
            

                # --- JSONに合計人数も保存 ---
                # === PDFファイル名から日付フォルダーを決定 ===
                pdf_name = os.path.basename(base_pdf)
                match = re.search(r"(\d{1,2})[.\-](\d{1,2})", pdf_name)
                if match:
                    month, day = match.groups()
                    try:
                        year = datetime.now().year
                        month = month.zfill(2)
                        day = day.zfill(2)
                        folder_name = f"{year}-{month}-{day}"
                        status_folder = os.path.join(os.path.dirname(base_pdf), "status_data", folder_name)
                        self.log_text.insert(tk.END, f"[INFO] PDF名から日付フォルダー決定: {folder_name}\n")
                    except Exception as e:
                        self.log_text.insert(tk.END, f"[WARN] 日付フォルダー名生成失敗: {e}\n")
                        status_folder = os.path.join(os.path.dirname(base_pdf), "status_data")
                else:
                    status_folder = os.path.join(os.path.dirname(base_pdf), "status_data")
                    self.log_text.insert(tk.END, "[INFO] PDF名に日付が含まれないため既定status_dataを使用。\n")

                os.makedirs(status_folder, exist_ok=True)
                json_path = os.path.join(status_folder, f"{flight_name}_status.json")

                try:
                    if os.path.exists(json_path):
                        #with open(json_path, "r", encoding="utf-8") as f:
                        #    data = json.load(f)
                        key = get_encryption_key()  # ← 自動生成＋永続再利用
                        fernet = Fernet(key)

                        with open(json_path, "rb") as f:
                            enc = f.read()

                        # 復号してからJSONとして読込
                        try:
                            dec = fernet.decrypt(enc)
                            data = json.loads(dec.decode("utf-8"))
                        except Exception as e:
                            self.log_text.insert(tk.END, f"[WARN] ステータスJSONの復号に失敗: {e}\n")
                            return
                    else:
                        data = {"records": []}
                except Exception:
                    data = {"records": []}

                # 合計人数レコードを更新／追加
                data["records"] = [r for r in data.get("records", []) if r.get("resv") != "合計人数"]
                data["records"].append(total_record)

                #with open(json_path, "w", encoding="utf-8") as f:
                #    json.dump(data, f, ensure_ascii=False, indent=2)
                # 暗号化キー読込
                key = get_encryption_key()  # ← 自動生成＋永続再利用
                fernet = Fernet(key)

                # JSON文字列化
                json_str = json.dumps(data, ensure_ascii=False, indent=2)

                # 暗号化してバイナリ書き込み
                enc = fernet.encrypt(json_str.encode("utf-8"))
                with open(json_path, "wb") as f:
                    f.write(enc)


                self.log_text.insert(tk.END, f"[JSON更新] 合計人数を保存しました ({json_path})\n")

            except Exception as e:
                self.log_text.insert(tk.END, f"[ERROR] フッター合計人数処理失敗: {e}\n")

        import time

        # --- PDF保存 ---
        temp_path = marked_pdf + ".tmp"
        doc_marked.save(temp_path)
        doc_marked.close()

        time.sleep(0.3)
        os.replace(temp_path, marked_pdf)
        self.log_text.insert(tk.END, f"[PDF保存] {os.path.basename(marked_pdf)} に追記完了。\n")

        # --- JSON保存 ---
        # --- JSON保存 ---
        # === PDFファイル名から日付フォルダーを決定 ===
        pdf_name = os.path.basename(base_pdf)
        match = re.search(r"(\d{1,2})[.\-](\d{1,2})", pdf_name)
        if match:
            month, day = match.groups()
            try:
                year = datetime.now().year
                month = month.zfill(2)
                day = day.zfill(2)
                folder_name = f"{year}-{month}-{day}"
                status_folder = os.path.join(os.path.dirname(base_pdf), "status_data", folder_name)
                self.log_text.insert(tk.END, f"[INFO] PDF名から日付フォルダー決定: {folder_name}\n")
            except Exception as e:
                self.log_text.insert(tk.END, f"[WARN] 日付フォルダー名生成失敗: {e}\n")
                status_folder = os.path.join(os.path.dirname(base_pdf), "status_data")
        else:
            status_folder = os.path.join(os.path.dirname(base_pdf), "status_data")
            self.log_text.insert(tk.END, "[INFO] PDF名に日付が含まれないため既定status_dataを使用。\n")

        os.makedirs(status_folder, exist_ok=True)
        json_path = os.path.join(status_folder, f"{flight_name}_status.json")


        # ▼▼▼ ここから置換：orig/after の堅牢な算出ロジック ▼▼▼
        import re as regex  # ← 変数名衝突を完全回避

        def _num_tail(token: object) -> int:
            s = str(token)
            if "→" in s:
                tail = s.split("→")[-1]
                tail = regex.sub(r"\D", "", tail)
                return int(tail) if tail else 0
            s = regex.sub(r"\D", "", s)
            return int(s) if s else 0

        def _num_head(token: object) -> int:
            s = str(token)
            if "→" in s:
                head = s.split("→")[0]
                head = regex.sub(r"\D", "", head)
                return int(head) if head else 0
            s = regex.sub(r"\D", "", s)
            return int(s) if s else 0

        def _get_after_values(status, item_id, vals):
            """JSONに書き込む減算後（after）を返す"""
            def _to_after_int(token):
                s = str(token)
                if "→" in s:
                    tail = s.split("→")[-1]
                    return int(tail) if tail.isdigit() else 0
                return int(s) if s.isdigit() else 0

            after_m = _to_after_int(vals[4]) if len(vals) > 4 else 0
            after_f = _to_after_int(vals[5]) if len(vals) > 5 else 0
            after_k = _to_after_int(vals[6]) if len(vals) > 6 else 0
            after_t = _to_after_int(vals[7]) if len(vals) > 7 else (after_m + after_f + after_k)

            ded = self.cxl_deduction_map.get(item_id)

            if status in ("CXL", "NS") and isinstance(ded, dict) and "after" in ded:
                a = ded["after"]
                after_m = int(a.get("男", after_m) or 0)
                after_f = int(a.get("女", after_f) or 0)
                after_k = int(a.get("子供", after_k) or 0)
                after_t = int(a.get("合計", after_m + after_f + after_k) or (after_m + after_f + after_k))

                # ★ CXL の場合で after == orig の場合 → 全て 0 に変更
                if status == "CXL" and isinstance(ded.get("orig"), dict):
                    o = ded["orig"]
                    if (
                        int(o.get("男", 0)) == after_m and
                        int(o.get("女", 0)) == after_f and
                        int(o.get("子供", 0)) == after_k
                    ):
                        self.log_text.insert(tk.END, f"[CXL変換] 減算なし検出 → after を全0に変換\n")
                        after_m = after_f = after_k = after_t = 0

            return after_m, after_f, after_k, after_t


        def _get_orig_values(status, item_id, vals):
            """
            JSONに書く orig（元値）。
            1) cxl_deduction_map に orig があれば最優先
            2) なければ Treeview の '2→1' の左側（単数はその数＝afterと同じになることも）
            """
            ded = self.cxl_deduction_map.get(item_id)
            if status in ("CXL", "NS") and isinstance(ded, dict) and "orig" in ded:
                o = ded["orig"]
                om = int(o.get("男", 0) or 0)
                of = int(o.get("女", 0) or 0)
                ok = int(o.get("子供", 0) or 0)
                ot = int(o.get("合計", om + of + ok) or (om + of + ok))
                return om, of, ok, ot

            om = _num_head(vals[4]) if len(vals) > 4 else 0
            of = _num_head(vals[5]) if len(vals) > 5 else 0
            ok = _num_head(vals[6]) if len(vals) > 6 else 0
            ot = _num_head(vals[7]) if len(vals) > 7 else (om + of + ok)
            return om, of, ok, ot
        # ▲▲▲ ここまで置換 ▲▲▲

        data = {
            "便名": flight_name,
            "pdf_path": base_pdf,
            "timestamp": datetime.now().isoformat(),
            "records": []
        }

        for item_id, status, resv, name, _, vals in targets:
            # after/orig を必ず両方確定（NSもCXLも同じ枠に格納する）
            after_m, after_f, after_k, after_t = _get_after_values(status, item_id, vals)
            orig_m,  orig_f,  orig_k,  orig_t  = _get_orig_values(status,  item_id, vals)

            record = {
                "resv": resv,
                "name": name,
                "status": status,
                # トップレベルは after（Excel 側で「乗車人数」に利用）
                "male": after_m,
                "female": after_f,
                "child": after_k,
                "total": after_t,
            }

            # NS / CXL は Excel 側で予定=orig合計 を使うため、常に orig/after を同梱
            if status in ("CXL", "CXL-CS", "NS"):
                record["cxl_deduction"] = {
                    "orig":  {"男": orig_m,  "女": orig_f,  "子供": orig_k,  "合計": orig_t},
                    "after": {"男": after_m, "女": after_f, "子供": after_k, "合計": after_t},
                }

            data["records"].append(record)

        #with open(json_path, "w", encoding="utf-8") as f:
        #    json.dump(data, f, ensure_ascii=False, indent=2)
        # 暗号化キー読込
        key = get_encryption_key()  # ← 自動生成＋永続再利用
        fernet = Fernet(key)

        # JSON文字列化
        json_str = json.dumps(data, ensure_ascii=False, indent=2)

        # 暗号化してバイナリ書き込み
        enc = fernet.encrypt(json_str.encode("utf-8"))
        with open(json_path, "wb") as f:
            f.write(enc)

        self.log_text.insert(tk.END, f"[JSON上書き] {json_path}\n")


        # すべて正常保存できたら、現在表示を新たな基準にする
        self.baseline_snapshot = self._make_snapshot_from_tree()
        self.unsaved_changes = False
        self.log_text.insert(tk.END, "[INFO] 保存完了 → 未保存フラグOFF\n")

        messagebox.showinfo("完了", "PDFへの書き込みが完了しました。", parent=self.root)



# -------------------------------------------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFPassengerSearchApp(root)
    root.mainloop()
else:
    # モジュールとして利用可能
    PDFPassengerSearchApp = PDFPassengerSearchApp
