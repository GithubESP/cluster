import json
import threading
import time
import re
import os
import sys
import random
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog  # 新增 filedialog
import pyperclip
import pydirectinput
import pyautogui
import keyboard


# ---------- 設定 ----------
# 判斷是否為打包後的環境 (Frozen/EXE)
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

MOD_FILE = os.path.join(BASE_DIR, "stats.ndjson")
CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
LOG_FILE = os.path.join(BASE_DIR, "roll_log.txt")

# 預設延遲 (秒)
DELAY_AFTER_CLICK = 0.25
DELAY_AFTER_COPY = 0.15
# --------------------------

stop_event = threading.Event()
roll_count = 0

# ---------- 檔案 / 模式處理 ----------
def is_target_mod(mod, filter_config):
    """ 使用設定檔中的關鍵字來判斷是否為目標詞綴 """
    ref = mod.get("ref", "")
    matchers = mod.get("matchers", [])
    
    ref_startswith_list = filter_config.get("ref_startswith", [])
    for keyword in ref_startswith_list:
        if ref.startswith(keyword):
            return True
            
    string_contains_list = filter_config.get("string_contains", [])
    for m in matchers:
        s = m.get("string", "")
        for keyword in string_contains_list:
            if keyword in s:
                return True

    return False

def load_mod_list(file_path, filter_config):
    mods = []
    # 如果路徑不存在，直接返回空陣列，讓 GUI 層處理
    if not os.path.exists(file_path):
        return mods

    try:
        with open(file_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    mod = json.loads(line)
                    if is_target_mod(mod, filter_config):
                        mods.append(mod)
                except Exception:
                    continue
    except Exception as e:
        print(f"讀取錯誤: {e}")
        return []
        
    return mods


def pattern_to_regex(pattern):
    escaped = re.escape(pattern)
    escaped = escaped.replace(r"\#", "#")
    num_pat = r"([-+]?\d*\.?\d+)"  # 捕獲數字 (包含小數和正負號)
    regex = escaped.replace("#", num_pat)
    regex = regex.replace(r"\ ", r"\s*")
    return re.compile(regex)


def mod_match_line_with_value(line, compiled_pattern):
    match = compiled_pattern.search(line)
    if not match:
        return None
    
    # 沒有捕獲組，代表是沒有 # 的詞綴，直接回傳 True 表示匹配成功
    if not match.groups():
        return True

    try:
        # 從第一個捕獲組中提取數值
        value = float(match.group(1))
        return value
    except (IndexError, ValueError):
        # 如果沒有捕獲組或轉換失敗，也當作是純文字匹配成功
        return True


def extract_mod_section_from_clipboard(text):
    if not text:
        return []
    sections = text.split("--------")
    if len(sections) < 2:
        return []
    mod_section = sections[-2] if len(sections) >= 2 else text
    lines = [ln.strip() for ln in mod_section.splitlines() if ln.strip()]
    return lines


def check_hit(mod_lines, target_mods, require_k=None):
    if not target_mods:
        return False, []

    hit_details = []
    
    # 預先編譯所有目標詞綴的正規表示式
    compiled_targets = []
    for t in target_mods:
        patterns = []
        mod = t['mod']
        for m in mod.get("matchers", []):
            pat = m.get("string", "")
            patterns.append(pattern_to_regex(pat))
        compiled_targets.append({
            "mod": mod,
            "min": t.get("min"),
            "max": t.get("max"),
            "patterns": patterns
        })

    hit_count = 0
    for target in compiled_targets:
        target_matched = False
        for p in target['patterns']:
            for line in mod_lines:
                # 使用新的匹配函式
                match_result = mod_match_line_with_value(line, p)

                if match_result is None:
                    continue

                # 提取數值和範圍
                val = match_result if isinstance(match_result, float) else None
                min_val = target.get("min")
                max_val = target.get("max")
                
                # 檢查數值範圍
                value_in_range = True
                if val is not None:
                    if min_val is not None and val < min_val:
                        value_in_range = False
                    if max_val is not None and val > max_val:
                        value_in_range = False

                if value_in_range:
                    target_matched = True
                    desc = target['mod'].get("matchers", [{}])[0].get("string", "??")
                    hit_details.append(f"{desc} ({val})" if val is not None else desc)
                    break # 找到符合的 line 就不用再找了
            
            if target_matched:
                break # 找到符合的 pattern 就不用再找了
        
        if target_matched:
            hit_count += 1
            
    # 判斷總命中結果
    final_hit = False
    if require_k is None:
        final_hit = (hit_count == len(target_mods))
    else:
        final_hit = (hit_count >= require_k)

    return final_hit, hit_details

# ---------- IO: config / log ----------

def save_config(cfg):
    try:
        # Preserve keywords if they exist in the current config on disk
        current_config = load_config()
        if current_config and 'mod_filter_keywords' in current_config:
            cfg['mod_filter_keywords'] = current_config['mod_filter_keywords']

        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
        print(f"成功儲存設定到 {CONFIG_FILE}")
    except Exception as e:
        print(f"儲存設定失敗: {e}")


def load_config():
    default_filters = {
        "mod_filter_keywords": {
            "ref_startswith": [
                "Added Small Passive Skills also grant",
                "1 Added Passive Skill is",
                "Added Small Passive Skills have"
            ],
            "string_contains": [
                "附加的小天賦給予",
                "附加的小型天賦給予",
                "附加的小天賦增加",
                "1 個附加天賦為"
            ]
        }
    }
    
    if not os.path.exists(CONFIG_FILE):
        # 如果設定檔不存在，建立一個包含預設篩選的
        save_config(default_filters)
        return default_filters

    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
            # 如果讀取的設定檔沒有篩選關鍵字，則補上
            if 'mod_filter_keywords' not in cfg:
                cfg.update(default_filters)
                # 馬上存回去
                save_config(cfg)
            return cfg
    except Exception as e:
        print(f"載入設定檔錯誤: {e}")
        return default_filters # 發生錯誤時返回預設值



def append_log_line(line):
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except:
        pass

# ---------- 偏移與點擊 ----------

def click_with_offset(x, y, offset=3, button="left"):
    ox = x + random.randint(-offset, offset)
    oy = y + random.randint(-offset, offset)
    pydirectinput.moveTo(ox, oy)
    time.sleep(0.01)
    if button == "left":
        pydirectinput.leftClick()
    else:
        pydirectinput.rightClick()


def do_click_sequence(alt_pos, cluster_pos, offset, click_delay, copy_delay, workflow="single", item2_pos=None):
    if workflow == 'double' and item2_pos is not None:
        # 兩種通貨流程
        # 1. 右鍵改造石 (通貨A)
        click_with_offset(alt_pos[0], alt_pos[1], offset=offset, button="right")
        time.sleep(0.06)
        # 2. 左鍵星團
        click_with_offset(cluster_pos[0], cluster_pos[1], offset=offset, button="left")
        time.sleep(0.06) 
        # 3. 右鍵通貨B
        click_with_offset(item2_pos[0], item2_pos[1], offset=offset, button="right")
        time.sleep(0.06)
        # 4. 左鍵星團
        click_with_offset(cluster_pos[0], cluster_pos[1], offset=offset, button="left")
        time.sleep(click_delay)
    else:
        # 原本的單一通貨流程
        # 右鍵改造石
        click_with_offset(alt_pos[0], alt_pos[1], offset=offset, button="right")
        time.sleep(0.06)
        # 左鍵星團
        click_with_offset(cluster_pos[0], cluster_pos[1], offset=offset, button="left")
        time.sleep(click_delay)

    # 共同的複製步驟
    pydirectinput.keyDown("ctrl")
    pydirectinput.press("c")
    pydirectinput.keyUp("ctrl")
    time.sleep(copy_delay)


# ---------- 背景 worker ----------

def worker_loop(gui_vars):
    global roll_count
    stop_event.clear()
    roll_count = 0
    start_time = datetime.now()
    gui_vars['append_log']("開始自動洗石: " + start_time.strftime("%Y-%m-%d %H:%M:%S"))
    while not stop_event.is_set():
        do_click_sequence(
            gui_vars['alt_pos'], 
            gui_vars['cluster_pos'], 
            gui_vars['offset'], 
            gui_vars['click_delay'], 
            gui_vars['copy_delay'],
            workflow=gui_vars['workflow'],
            item2_pos=gui_vars.get('item2_pos')
        )
        roll_count += 1
        gui_vars['set_count'](roll_count)
        clip = pyperclip.paste()
        mod_lines = extract_mod_section_from_clipboard(clip)
        
        hit, hit_details = check_hit(mod_lines, gui_vars['targets'], require_k=gui_vars['require_k'])
        
        log_msg = f"[{roll_count}] 讀取 {len(mod_lines)} 行 -> {'HIT' if hit else 'MISS'}"
        if hit:
            log_msg += " | " + ", ".join(hit_details)
            
        gui_vars['append_log'](log_msg)
        append_log_line(f"{datetime.now().isoformat()} | #{roll_count} | HIT={hit} | details={','.join(hit_details)} | lines={len(mod_lines)}")

        if hit:
            gui_vars['append_log']("命中條件，停止腳本。")
            try:
                pyautogui.alert('命中條件，停止腳本！') 
            except Exception as e:
                print(f"無法顯示 alert: {e}")
            stop_event.set()
        
        time.sleep(gui_vars.get('loop_delay', 0.2))

    end_time = datetime.now()
    duration = end_time - start_time
    gui_vars['append_log']("腳本結束: " + end_time.strftime("%Y-%m-%d %H:%M:%S"))
    gui_vars['append_log'](f"總共洗了 {roll_count} 次，耗時: {duration}")
    # 在背景執行緒結束後，通知主執行緒更新 UI
    gui_vars['on_stop']()


# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Cluster Washer (Press ESC to stop)")

        # 1. 決定要讀取的檔案路徑 (預設為全域變數 MOD_FILE)
        self.current_mod_file = MOD_FILE

        # 2. 檢查檔案是否存在，若不存在則詢問使用者
        if not os.path.exists(self.current_mod_file):
            user_choice = messagebox.askyesno(
                "找不到檔案", 
                f"在預設路徑找不到 stats.ndjson：\n{self.current_mod_file}\n\n是否要手動選擇檔案？"
            )
            
            if user_choice:
                selected_file = filedialog.askopenfilename(
                    title="請選擇 stats.ndjson",
                    initialdir=BASE_DIR,
                    filetypes=[("NDJSON files", "*.ndjson"), ("JSON files", "*.json"), ("All files", "*.*")]
                )
                if selected_file:
                    self.current_mod_file = selected_file

        # 3. 載入詞綴資料 (使用剛剛決定的路徑)
        self.loaded_config = load_config()
        self.mods = load_mod_list(self.current_mod_file, self.loaded_config.get("mod_filter_keywords", {}))
        print(f"載入 {len(self.mods)} 個詞墜 (來源: {self.current_mod_file})")

        # ---- 座標設定（預設值）----
        self.alt_pos = (100, 200)
        self.cluster_pos = (300, 400)
        
        if self.loaded_config:
            self.alt_pos = tuple(self.loaded_config.get("alt_pos", self.alt_pos))
            self.cluster_pos = tuple(self.loaded_config.get("cluster_pos", self.cluster_pos))
        
        # StringVar 用於 Entry 顯示
        self.alteration_pos_var = tk.StringVar(value=f"{self.alt_pos[0]},{self.alt_pos[1]}")
        self.cluster_pos_var = tk.StringVar(value=f"{self.cluster_pos[0]},{self.cluster_pos[1]}")

        # ---- 搜尋與篩選 ----
        self.search_var = tk.StringVar()
        self.filtered_affixes = []
        self.cluster_affixes = self.mods[:]
        self.filtered_indices = list(range(len(self.mods)))

        # ---- 滑鼠座標追蹤 ----
        self.follow_mouse = tk.BooleanVar(value=False)

        # ---- 快捷鍵設定 ----
        self.alt_hotkey_var = tk.StringVar(value='f5')
        self.cluster_hotkey_var = tk.StringVar(value='f6')
        self.start_hotkey_var = tk.StringVar(value='f9')
        self.stop_hotkey_var = tk.StringVar(value='f10')
        self.item2_hotkey_var = tk.StringVar(value='f7')
        self._registered_alt_hotkey = None
        self._registered_cluster_hotkey = None
        self._registered_start_hotkey = None
        self._registered_stop_hotkey = None
        self._registered_item2_hotkey = None

        # ---- 流程設定 ----
        self.workflow_var = tk.StringVar(value="single") # 'single' or 'double'
        self.item2_pos = (100, 300)
        self.item2_pos_var = tk.StringVar(value=f"{self.item2_pos[0]},{self.item2_pos[1]}")

        # ---- Prefix / Suffix 暫存 ----
        # 改為儲存 dict，包含 mod 物件和 min/max
        self.selected_prefix = []
        self.selected_suffix = []

        # ---- 命中模式 ----
        self.require_k_mode = tk.StringVar(value="all")
        self.k_value = tk.IntVar(value=1)
        # ---- 延遲設定 ----
        self.loop_delay = tk.DoubleVar(value=0.2)
        self.click_delay = tk.DoubleVar(value=DELAY_AFTER_CLICK)
        self.copy_delay = tk.DoubleVar(value=DELAY_AFTER_COPY)

        # ---- 偏移像素 ----
        self.offset = tk.IntVar(value=3)

        # ---- 建立介面 ----
        self.create_widgets()

        # ---- 載入其他設定到 GUI（前後綴、延遲等）----
        if self.loaded_config:
            self.load_other_settings(self.loaded_config)
        
        # ---- 設定快捷鍵 ----
        self._setup_hotkeys()

        # ---- 開始更新滑鼠座標 ----
        self.update_mouse_pos()

        # 若載入失敗，顯示提示在 Log
        if not self.mods:
            self.append_log("[警告] 未載入任何詞綴，請確認 stats.ndjson 路徑。")
            self.append_log(f"嘗試路徑: {self.current_mod_file}")

    def _setup_hotkeys(self):
        try:
            # 移除舊的快捷鍵
            if self._registered_alt_hotkey:
                keyboard.remove_hotkey(self._registered_alt_hotkey)
            if self._registered_cluster_hotkey:
                keyboard.remove_hotkey(self._registered_cluster_hotkey)
            if self._registered_item2_hotkey:
                keyboard.remove_hotkey(self._registered_item2_hotkey)
            if self._registered_start_hotkey:
                keyboard.remove_hotkey(self._registered_start_hotkey)
            if self._registered_stop_hotkey:
                keyboard.remove_hotkey(self._registered_stop_hotkey)
        except Exception as e:
            # 即使移除失敗也繼續，可能是因為從未設定過
            print(f"移除舊快捷鍵時發生錯誤: {e}")

        try:
            alt_key = self.alt_hotkey_var.get()
            cluster_key = self.cluster_hotkey_var.get()
            item2_key = self.item2_hotkey_var.get()
            start_key = self.start_hotkey_var.get()
            stop_key = self.stop_hotkey_var.get()

            if alt_key:
                self._registered_alt_hotkey = keyboard.add_hotkey(alt_key, self.record_alt_pos)
            if cluster_key:
                self._registered_cluster_hotkey = keyboard.add_hotkey(cluster_key, self.record_cluster_pos)
            if item2_key:
                self._registered_item2_hotkey = keyboard.add_hotkey(item2_key, self.record_item2_pos)
            if start_key:
                self._registered_start_hotkey = keyboard.add_hotkey(start_key, self.start_roll)
            if stop_key:
                self._registered_stop_hotkey = keyboard.add_hotkey(stop_key, self.stop_roll)
            
            self.append_log(f"快捷鍵已更新 (改造石: {alt_key}, 星團: {cluster_key}, 通貨B: {item2_key}, 開始: {start_key}, 停止: {stop_key})")
        except Exception as e:
            self.append_log(f"[錯誤] 無法設定快捷鍵: {e}")
            messagebox.showerror("快捷鍵錯誤", f"無法設定快捷鍵：{e}\n請確認按鍵名稱是否正確。")

    def _on_tree_double_click(self, event, tree):
        region = tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column_id = tree.identify_column(event.x)
        # 取得內部的欄位名 (desc, min, max)
        column_name = tree.column(column_id, "id")
        
        # 只允許編輯 min 和 max
        if column_name not in ["min", "max"]:
            return

        item_id = tree.focus()
        if not item_id:
            return

        # 獲取儲存格位置
        x, y, width, height = tree.bbox(item_id, column_id)

        # 建立一個 Entry 來編輯
        val = tree.set(item_id, column_id)
        entry_var = tk.StringVar(value=val)
        entry = ttk.Entry(tree, textvariable=entry_var)
        entry.place(x=x, y=y, width=width, height=height)
        entry.focus_set()

        def on_focus_out(event):
            update_value()
            entry.destroy()
        
        def on_return(event):
            update_value()
            entry.destroy()

        def update_value():
            new_val_str = entry_var.get().strip()
            tree.set(item_id, column_id, new_val_str)
            
            # 更新後端資料
            new_val = None
            if new_val_str:
                try:
                    new_val = float(new_val_str)
                except ValueError:
                    # 如果輸入不是數字，則清空
                    tree.set(item_id, column_id, "")
                    new_val = None
            
            # 找出是哪個 treeview，並更新對應的 selected list
            is_prefix = (tree == self.prefix_tree)
            data_list = self.selected_prefix if is_prefix else self.selected_suffix
            
            # Treeview 的 iid 是從 I001 開始的十六進制，轉成 index
            try:
                idx = int(item_id.replace("I", ""), 16) - 1
                if 0 <= idx < len(data_list):
                    if column_name == "min":
                        data_list[idx]["min"] = new_val
                    else: # max
                        data_list[idx]["max"] = new_val
            except ValueError:
                # iid 格式可能不是預期的，忽略
                pass

        entry.bind("<FocusOut>", on_focus_out)
        entry.bind("<Return>", on_return)


    def create_widgets(self):
        frm = ttk.Frame(self.root, padding=8)
        frm.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(frm)
        left.grid(row=0, column=0, sticky="nswe", padx=4, pady=4)

        # --- 搜尋框 ---
        affix_search_frame = ttk.Frame(left)
        affix_search_frame.pack(fill="x", padx=4, pady=(4,0))
        ttk.Label(affix_search_frame, text="搜尋詞綴：").pack(side="left")
        search_entry = ttk.Entry(affix_search_frame, textvariable=self.search_var)
        search_entry.pack(side="left", fill="x", expand=True)
        self.search_var.trace_add("write", self.filter_affix_list)

        # --- Listbox ---
        # 顯示當前讀取的檔名 (取 basename 避免太長)
        filename = os.path.basename(self.current_mod_file)
        ttk.Label(left, text=f"詞墜清單 ({filename})").pack(anchor="w")
        
        self.mods_listbox = tk.Listbox(left, width=48, height=18)
        self.mods_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(left, orient=tk.VERTICAL, command=self.mods_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.mods_listbox.config(yscrollcommand=scrollbar.set)

        # 初始載入資料
        for i, m in enumerate(self.mods):
            desc = m.get("matchers", [{}])[0].get("string", m.get("ref", f"mod{i}"))
            self.mods_listbox.insert(tk.END, f"{i+1}. {desc}")

        mid = ttk.Frame(frm)
        mid.grid(row=0, column=1, sticky="n", padx=4)
        ttk.Button(mid, text="加入前綴 <<", command=self.add_prefix).pack(fill=tk.X, pady=2)
        ttk.Button(mid, text="加入後綴 <<", command=self.add_suffix).pack(fill=tk.X, pady=2)
        ttk.Button(mid, text="清除前綴", command=self.clear_prefix).pack(fill=tk.X, pady=8)
        ttk.Button(mid, text="清除後綴", command=self.clear_suffix).pack(fill=tk.X, pady=2)

        right = ttk.Frame(frm)
        right.grid(row=0, column=2, sticky="nswe", padx=4, pady=4)
        
        # --- 可滾動的右側框架 ---
        right_canvas = tk.Canvas(frm)
        right_canvas.grid(row=0, column=2, sticky="nswe", padx=4, pady=4)

        right_scrollbar = ttk.Scrollbar(frm, orient="vertical", command=right_canvas.yview)
        right_scrollbar.grid(row=0, column=3, sticky="ns")
        
        right_canvas.configure(yscrollcommand=right_scrollbar.set)
        right_canvas.bind('<Configure>', lambda e: right_canvas.configure(scrollregion = right_canvas.bbox("all")))

        scrollable_frame = ttk.Frame(right_canvas)
        right_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")

        # 這裡的 'right' 變成了 'scrollable_frame'
        # --- Prefix Treeview ---
        ttk.Label(scrollable_frame, text="前綴 (雙擊 Min/Max 編輯)").pack(anchor="w")
        prefix_tree_frame = ttk.Frame(scrollable_frame)
        prefix_tree_frame.pack(fill=tk.X, expand=True)
        self.prefix_tree = ttk.Treeview(prefix_tree_frame, columns=("desc", "min", "max"), show="headings", height=5)
        self.prefix_tree.heading("desc", text="詞綴")
        self.prefix_tree.heading("min", text="Min")
        self.prefix_tree.heading("max", text="Max")
        self.prefix_tree.column("desc", width=260)
        self.prefix_tree.column("min", width=50, anchor='center')
        self.prefix_tree.column("max", width=50, anchor='center')
        self.prefix_tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.prefix_tree.bind("<Double-1>", lambda e: self._on_tree_double_click(e, self.prefix_tree))

        # --- Suffix Treeview ---
        ttk.Label(scrollable_frame, text="後綴 (雙擊 Min/Max 編輯)").pack(anchor="w", pady=(6,0))
        suffix_tree_frame = ttk.Frame(scrollable_frame)
        suffix_tree_frame.pack(fill=tk.X, expand=True)
        self.suffix_tree = ttk.Treeview(suffix_tree_frame, columns=("desc", "min", "max"), show="headings", height=5)
        self.suffix_tree.heading("desc", text="詞綴")
        self.suffix_tree.heading("min", text="Min")
        self.suffix_tree.heading("max", text="Max")
        self.suffix_tree.column("desc", width=260)
        self.suffix_tree.column("min", width=50, anchor='center')
        self.suffix_tree.column("max", width=50, anchor='center')
        self.suffix_tree.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.suffix_tree.bind("<Double-1>", lambda e: self._on_tree_double_click(e, self.suffix_tree))

        # ---- 滑鼠座標設定區 ----
        pos_frm = ttk.LabelFrame(scrollable_frame, text="滑鼠座標 (設定後會儲存)")
        pos_frm.pack(fill=tk.X, pady=6)
        
        # 改造石座標 Entry（綁定正確的變數）
        ttk.Label(pos_frm, text="改造石座標 (按 F5 設定):").pack(anchor="w", padx=4, pady=(4,0))
        self.alt_entry = ttk.Entry(pos_frm, textvariable=self.alteration_pos_var)
        self.alt_entry.pack(fill=tk.X, padx=4, pady=2)
        
        # 星團座標 Entry（綁定正確的變數）
        ttk.Label(pos_frm, text="星團珠座標 (按 F6 設定):").pack(anchor="w", padx=4, pady=(8,0))
        self.cluster_entry = ttk.Entry(pos_frm, textvariable=self.cluster_pos_var)
        self.cluster_entry.pack(fill=tk.X, padx=4, pady=2)
        
        ttk.Checkbutton(pos_frm, text="顯示即時滑鼠座標", variable=self.follow_mouse).pack(anchor="w", padx=4, pady=(8,0))
        self.mouse_pos_label = ttk.Label(pos_frm, text="鼠標: (0,0)")
        self.mouse_pos_label.pack(anchor="w", padx=4, pady=(2,4))

        # ---- 快捷鍵設定 ----
        hotkey_frm = ttk.LabelFrame(scrollable_frame, text="快捷鍵設定")
        hotkey_frm.pack(fill=tk.X, pady=6)
        
        alt_hotkey_frm = ttk.Frame(hotkey_frm)
        alt_hotkey_frm.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(alt_hotkey_frm, text="設定改造石座標:").pack(side=tk.LEFT)
        ttk.Entry(alt_hotkey_frm, textvariable=self.alt_hotkey_var, width=10).pack(side=tk.LEFT, padx=4)

        cluster_hotkey_frm = ttk.Frame(hotkey_frm)
        cluster_hotkey_frm.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(cluster_hotkey_frm, text="設定星團珠座標:").pack(side=tk.LEFT)
        ttk.Entry(cluster_hotkey_frm, textvariable=self.cluster_hotkey_var, width=10).pack(side=tk.LEFT, padx=4)

        item2_hotkey_frm = ttk.Frame(hotkey_frm)
        item2_hotkey_frm.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(item2_hotkey_frm, text="設定通貨B座標:").pack(side=tk.LEFT)
        ttk.Entry(item2_hotkey_frm, textvariable=self.item2_hotkey_var, width=10).pack(side=tk.LEFT, padx=4)

        start_hotkey_frm = ttk.Frame(hotkey_frm)
        start_hotkey_frm.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(start_hotkey_frm, text="開始快捷鍵:").pack(side=tk.LEFT)
        ttk.Entry(start_hotkey_frm, textvariable=self.start_hotkey_var, width=10).pack(side=tk.LEFT, padx=4)

        stop_hotkey_frm = ttk.Frame(hotkey_frm)
        stop_hotkey_frm.pack(fill=tk.X, padx=4, pady=2)
        ttk.Label(stop_hotkey_frm, text="停止快捷鍵:").pack(side=tk.LEFT)
        ttk.Entry(stop_hotkey_frm, textvariable=self.stop_hotkey_var, width=10).pack(side=tk.LEFT, padx=4)
        
        ttk.Button(hotkey_frm, text="套用快捷鍵", command=self._setup_hotkeys).pack(fill=tk.X, padx=4, pady=4)

        # ---- 流程選擇 ----
        flow_frm = ttk.LabelFrame(scrollable_frame, text="流程選擇")
        flow_frm.pack(fill=tk.X, pady=6)
        ttk.Radiobutton(flow_frm, text="單一通貨 (改造石)", variable=self.workflow_var, value="single").pack(anchor="w", padx=4)
        ttk.Radiobutton(flow_frm, text="兩種通貨 (改造石 -> 通貨B)", variable=self.workflow_var, value="double").pack(anchor="w", padx=4)

        # ---- 通貨B座標設定 ----
        item2_pos_frm = ttk.LabelFrame(scrollable_frame, text="通貨B座標 (可選)")
        item2_pos_frm.pack(fill=tk.X, pady=6)
        ttk.Label(item2_pos_frm, text="通貨B座標 (按 F7 設定):").pack(anchor="w", padx=4, pady=(4,0))
        self.item2_entry = ttk.Entry(item2_pos_frm, textvariable=self.item2_pos_var)
        self.item2_entry.pack(fill=tk.X, padx=4, pady=2)
        
        # ---- 命中規則 ----
        mode_frm = ttk.LabelFrame(scrollable_frame, text="命中規則")
        mode_frm.pack(fill=tk.X, pady=4)
        ttk.Radiobutton(mode_frm, text="全部命中 (AND)", variable=self.require_k_mode, value="all").pack(anchor="w")
        ttk.Radiobutton(mode_frm, text="至少 K of N 命中", variable=self.require_k_mode, value="k_of_n").pack(anchor="w")
        
        kfrm = ttk.Frame(mode_frm)
        kfrm.pack(fill=tk.X)
        ttk.Label(kfrm, text="K=").pack(side=tk.LEFT)
        ttk.Spinbox(kfrm, from_=1, to=6, textvariable=self.k_value, width=5).pack(side=tk.LEFT, padx=4)
        ttk.Label(kfrm, text="迴圈延遲 (秒)：").pack(side=tk.LEFT, padx=(8,2))
        ttk.Spinbox(kfrm, from_=0.05, to=5.0, increment=0.05, textvariable=self.loop_delay, width=6).pack(side=tk.LEFT)

        # ---- 偏移設定 ----
        offset_frm = ttk.LabelFrame(scrollable_frame, text="仿人物差-滑鼠偏移設定")
        offset_frm.pack(fill=tk.X, pady=4)
        ttk.Label(offset_frm, text="偏移範圍 (像素 ±):").pack(anchor="w")

        # 顯示 Label
        self.offset_label = ttk.Label(offset_frm, text=f"{self.offset.get()} Pixels")
        self.offset_label.pack(anchor="e", padx=4)

        def update_offset_label(value):
            self.offset_label.config(text=f"{int(float(value))} Pixels")

        ttk.Scale(offset_frm, from_=0, to=12, orient=tk.HORIZONTAL, variable=self.offset, command=update_offset_label).pack(fill=tk.X, padx=4)

        ttk.Label(offset_frm, text="點擊延遲 (秒)：").pack(anchor="w", pady=(4,0))
        ttk.Entry(offset_frm, textvariable=self.click_delay).pack(fill=tk.X, padx=4)
        ttk.Label(offset_frm, text="Ctrl+C 延遲 (秒)：").pack(anchor="w", pady=(4,0))
        ttk.Entry(offset_frm, textvariable=self.copy_delay).pack(fill=tk.X, padx=4)

        # ---- 控制按鈕 ----
        ctl_frm = ttk.Frame(frm)
        ctl_frm.grid(row=1, column=0, columnspan=3, sticky="we", pady=(6,0))
        ttk.Button(ctl_frm, text="載入詞檔", command=self.reload_mods).pack(side=tk.LEFT, padx=4)
        ttk.Button(ctl_frm, text="儲存設定", command=self.save_config).pack(side=tk.LEFT, padx=4)
        self.start_btn = ttk.Button(ctl_frm, text="Start", command=self.start_roll)
        self.start_btn.pack(side=tk.LEFT, padx=4)
        self.stop_btn = ttk.Button(ctl_frm, text="Stop", command=self.stop_roll, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=4)
        ttk.Button(ctl_frm, text="匯出 log 檔", command=self.open_log_file).pack(side=tk.LEFT, padx=4)

        self.count_label = ttk.Label(ctl_frm, text="已洗次數: 0")
        self.count_label.pack(side=tk.RIGHT, padx=6)

        # ---- Log 區域 ----
        log_frame = ttk.Frame(frm)
        log_frame.grid(row=2, column=0, columnspan=3, sticky="we", pady=(6,0))
        ttk.Label(log_frame, text="Log").pack(anchor="w")
        self.log_text = tk.Text(log_frame, height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def filter_affix_list(self, *args):
        keyword = self.search_var.get().strip().lower()

        self.filtered_affixes = []
        self.filtered_indices = []

        for idx, affix in enumerate(self.cluster_affixes):
            ref = affix.get("ref", "").lower()
            matchers = [m.get("string","").lower() for m in affix.get("matchers",[])]
            if not keyword or keyword in ref or any(keyword in m for m in matchers):
                self.filtered_affixes.append(affix)
                self.filtered_indices.append(idx)

        # 重繪 Listbox
        self.mods_listbox.delete(0, tk.END)
        for affix in self.filtered_affixes:
            desc = affix.get("matchers", [{}])[0].get("string", affix.get("ref",""))
            self.mods_listbox.insert(tk.END, desc)

    def add_prefix(self):
        sel = self.mods_listbox.curselection()
        if not sel or len(sel) == 0:
            return

        ui_idx = sel[0]
        real_idx = self.filtered_indices[ui_idx]
        mod = self.mods[real_idx]

        desc = mod.get("matchers", [{}])[0].get("string", mod.get("ref", f"mod{ui_idx}"))
        
        # 新增到 treeview
        self.prefix_tree.insert("", tk.END, values=(desc, "", ""))
        # 新增到後端資料
        self.selected_prefix.append({"mod": mod, "min": None, "max": None, "original_index": real_idx})


    def add_suffix(self):
        sel = self.mods_listbox.curselection()
        if not sel or len(sel) == 0:
            return

        ui_idx = sel[0]
        real_idx = self.filtered_indices[ui_idx]
        mod = self.mods[real_idx]

        desc = mod.get("matchers", [{}])[0].get("string", mod.get("ref", f"mod{ui_idx}"))
        
        # 新增到 treeview
        self.suffix_tree.insert("", tk.END, values=(desc, "", ""))
        # 新增到後端資料
        self.selected_suffix.append({"mod": mod, "min": None, "max": None, "original_index": real_idx})

    def clear_prefix(self):
        for i in self.prefix_tree.get_children():
            self.prefix_tree.delete(i)
        self.selected_prefix = []

    def clear_suffix(self):
        for i in self.suffix_tree.get_children():
            self.suffix_tree.delete(i)
        self.selected_suffix = []

    def record_alt_pos(self):
        x, y = pyautogui.position()
        self.alt_pos = (x, y)
        self.alteration_pos_var.set(f"{x},{y}")
        self.append_log(f"已設定改造石座標: {x},{y}")

    def record_cluster_pos(self):
        x, y = pyautogui.position()
        self.cluster_pos = (x, y)
        self.cluster_pos_var.set(f"{x},{y}")
        self.append_log(f"已設定星團座標: {x},{y}")

    def record_item2_pos(self):
        x, y = pyautogui.position()
        self.item2_pos = (x, y)
        self.item2_pos_var.set(f"{x},{y}")
        self.append_log(f"已設定通貨B座標: {x},{y}")

    def update_mouse_pos(self):
        x, y = pyautogui.position()
        self.mouse_pos_label.config(text=f"鼠標: ({x},{y})")
        if self.follow_mouse.get():
            self.append_log(f"[DEBUG] 鼠標即時座標: ({x},{y})")
        self.root.after(100, self.update_mouse_pos)

    def append_log(self, text):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{ts}] {text}\n")
        self.log_text.see(tk.END)

    def set_count(self, n):
        self.count_label.config(text=f"已洗次數: {n}")

    def reload_mods(self):
        # 這裡改為使用 self.current_mod_file，這樣就能重新載入「當前選中的檔案」
        self.loaded_config = load_config()
        self.mods = load_mod_list(self.current_mod_file, self.loaded_config.get("mod_filter_keywords", {}))
        self.cluster_affixes = self.mods[:]
        self.filtered_indices = list(range(len(self.mods)))
        
        print(f"重新載入 {len(self.mods)} 個詞墜，來源: {self.current_mod_file}")
        self.mods_listbox.delete(0, tk.END)
        for i, m in enumerate(self.mods):
            desc = m.get("matchers", [{}])[0].get("string", m.get("ref", f"mod{i}"))
            self.mods_listbox.insert(tk.END, f"{i+1}. {desc}")
        self.append_log(f"已重新載入詞檔 ({os.path.basename(self.current_mod_file)})")

    def _update_coords_from_vars(self):
        try:
            alt_coords = self.alteration_pos_var.get().split(',')
            self.alt_pos = (int(alt_coords[0].strip()), int(alt_coords[1].strip()))
        except (ValueError, IndexError):
            self.append_log("[錯誤] 改造石座標格式不正確，請使用 x,y 格式。")
            messagebox.showwarning("格式錯誤", "改造石座標格式不正確，請使用 x,y 格式。")
        
        try:
            cluster_coords = self.cluster_pos_var.get().split(',')
            self.cluster_pos = (int(cluster_coords[0].strip()), int(cluster_coords[1].strip()))
        except (ValueError, IndexError):
            self.append_log("[錯誤] 星團珠寶座標格式不正確，請使用 x,y 格式。")
            messagebox.showwarning("格式錯誤", "星團珠寶座標格式不正確，請使用 x,y 格式。")

        try:
            item2_coords = self.item2_pos_var.get().split(',')
            self.item2_pos = (int(item2_coords[0].strip()), int(item2_coords[1].strip()))
        except (ValueError, IndexError):
            self.append_log("[錯誤] 通貨B座標格式不正確，請使用 x,y 格式。")
            messagebox.showwarning("格式錯誤", "通貨B座標格式不正確，請使用 x,y 格式。")

    def save_config(self):
        self._update_coords_from_vars()
        
        # Helper to convert data for saving
        def prep_list(data_list):
            return [
                {
                    "index": item["original_index"],
                    "min": item["min"],
                    "max": item["max"]
                }
                for item in data_list
            ]

        cfg = {
            "alt_pos": self.alt_pos,
            "cluster_pos": self.cluster_pos,
            "selected_prefixes": prep_list(self.selected_prefix),
            "selected_suffixes": prep_list(self.selected_suffix),
            "require_k_mode": self.require_k_mode.get(),
            "k_value": self.k_value.get(),
            "loop_delay": self.loop_delay.get(),
            "offset": self.offset.get(),
            "click_delay": float(self.click_delay.get()),
            "copy_delay": float(self.copy_delay.get()),
            "alt_hotkey": self.alt_hotkey_var.get(),
            "cluster_hotkey": self.cluster_hotkey_var.get(),
            "start_hotkey": self.start_hotkey_var.get(),
            "stop_hotkey": self.stop_hotkey_var.get(),
            "item2_hotkey": self.item2_hotkey_var.get(),
            "workflow": self.workflow_var.get(),
            "item2_pos": self.item2_pos
        }
        print("[DEBUG] Saving config:", cfg) # 增加存檔前的 Log
        save_config(cfg)
        self.append_log("設定已儲存。")
        print(f"[DEBUG] 儲存座標: alt={self.alt_pos}, cluster={self.cluster_pos}")

    def load_other_settings(self, cfg):
        """載入前後綴和其他設定（座標已在 __init__ 中載入）"""
        
        # Helper to load and populate a tree and data list
        def load_list(key, tree, data_list):
            saved_items = cfg.get(key, [])
            for item in saved_items:
                idx = item.get("index")
                if not (isinstance(idx, int) and 0 <= idx < len(self.mods)):
                    continue
                
                mod = self.mods[idx]
                min_val = item.get("min")
                max_val = item.get("max")
                
                desc = mod.get("matchers", [{}])[0].get("string", mod.get("ref", ""))
                min_str = str(min_val) if min_val is not None else ""
                max_str = str(max_val) if max_val is not None else ""
                
                tree.insert("", tk.END, values=(desc, min_str, max_str))
                data_list.append({
                    "mod": mod,
                    "min": min_val,
                    "max": max_val,
                    "original_index": idx
                })

        try:
            self.clear_prefix()
            self.clear_suffix()
            load_list("selected_prefixes", self.prefix_tree, self.selected_prefix)
            load_list("selected_suffixes", self.suffix_tree, self.selected_suffix)
        except Exception as e:
            print(f"載入前後綴錯誤: {e}")
            self.append_log(f"[ERROR] 載入詞綴選項失敗: {e}")
    
        # 其他設定
        try:
            self.require_k_mode.set(cfg.get("require_k_mode", "all"))
            self.k_value.set(cfg.get("k_value", 1))
            self.loop_delay.set(cfg.get("loop_delay", 0.2))
            self.offset.set(cfg.get("offset", 3))
            self.click_delay.set(cfg.get("click_delay", DELAY_AFTER_CLICK))
            self.copy_delay.set(cfg.get("copy_delay", DELAY_AFTER_COPY))
            self.alt_hotkey_var.set(cfg.get("alt_hotkey", "f5"))
            self.cluster_hotkey_var.set(cfg.get("cluster_hotkey", "f6"))
            self.start_hotkey_var.set(cfg.get("start_hotkey", "f9"))
            self.stop_hotkey_var.set(cfg.get("stop_hotkey", "f10"))
            self.item2_hotkey_var.set(cfg.get("item2_hotkey", "f7"))
            
            self.workflow_var.set(cfg.get("workflow", "single"))
            item2_pos_loaded = cfg.get("item2_pos", None)
            if item2_pos_loaded and isinstance(item2_pos_loaded, list) and len(item2_pos_loaded) == 2:
                self.item2_pos = tuple(item2_pos_loaded)
                self.item2_pos_var.set(f"{self.item2_pos[0]},{self.item2_pos[1]}")
            
            # 載入後要重新套用快捷鍵
            self._setup_hotkeys()
            
        except Exception as e:
            print(f"載入設定錯誤: {e}")
            self.append_log(f"[ERROR] 載入一般設定失敗: {e}")
    
        self.append_log("已載入先前設定。")

    def open_log_file(self):
        if os.path.exists(LOG_FILE):
            os.startfile(LOG_FILE)
        else:
            messagebox.showinfo("Log", "尚未有 log。")

    def on_worker_stop(self):
        # This function is called from the worker thread, so we need to
        # schedule the GUI update on the main thread.
        self.root.after(0, self.update_ui_after_stop)

    def update_ui_after_stop(self):
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    def start_roll(self):
        if not self.selected_prefix and not self.selected_suffix:
            if not messagebox.askyesno("確認", "尚未選擇任何目標詞墜，要繼續嗎？"):
                return

        # Make sure coords are updated from text entry
        self._update_coords_from_vars()

        targets = self.selected_prefix + self.selected_suffix
        require_k = None
        if self.require_k_mode.get() == "k_of_n":
            require_k = int(self.k_value.get())
            if require_k < 1 or require_k > len(targets):
                messagebox.showwarning("設定錯誤", "K 值不合法。")
                return

        gui_vars = {
            "alt_pos": self.alt_pos,
            "cluster_pos": self.cluster_pos,
            "item2_pos": self.item2_pos,
            "workflow": self.workflow_var.get(),
            "targets": targets,
            "require_k": require_k,
            "loop_delay": float(self.loop_delay.get()),
            "append_log": self.append_log,
            "set_count": self.set_count,
            "offset": int(self.offset.get()),
            "click_delay": float(self.click_delay.get()),
            "copy_delay": float(self.copy_delay.get()),
            "on_stop": self.on_worker_stop
        }

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.append_log("執行緒啟動中...")
        t = threading.Thread(target=worker_loop, args=(gui_vars,), daemon=True)
        t.start()

    def stop_roll(self):
        stop_event.set()
        self.stop_btn.config(state=tk.DISABLED) # Disable immediately for feedback
        self.append_log("停止訊號已發出。")


def emergency_stop(event=None):
    stop_event.set()
    print("緊急停止已觸發！")


# 設定 Esc 為緊急停止
keyboard.add_hotkey('esc', emergency_stop)


def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()