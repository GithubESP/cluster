import json
import threading
import time
import re
import os
import random
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip
import pydirectinput
import pyautogui
import keyboard


# ---------- 設定 ----------
# 取得目前腳本所在的絕對路徑目錄
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 使用 join 確保路徑指向腳本同層目錄
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
def is_cluster_mod(mod):
    ref = mod.get("ref", "")
    matchers = mod.get("matchers", [])

    # A. 小被動詞
    if ref.startswith("Added Small Passive Skills also grant"):
        return True

    # B. notable 大詞
    if ref.startswith("1 Added Passive Skill is"):
        return True

    # C. 中文 fallback 過濾（避免資料來源不一致）
    for m in matchers:
        s = m.get("string", "")
        if "附加的小天賦給予" in s:
            return True
        if "附加的小型天賦給予" in s:
            return True
        if "1 個附加天賦為" in s:
            return True

    return False

def load_mod_list(file_path):
    mods = []
    if not os.path.exists(file_path):
        return mods
    with open(file_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue

            mod = json.loads(line)

            try:
                if is_cluster_mod(mod):
                    mods.append(mod)
            except Exception:
                continue
    return mods


def pattern_to_regex(pattern):
    escaped = re.escape(pattern)
    escaped = escaped.replace(r"\#", "#")
    num_pat = r"[-+]?\d+%?"
    regex = escaped.replace("#", num_pat)
    regex = regex.replace(r"\ ", r"\s*")
    return re.compile(regex)


def mod_match_line(line, compiled_pattern):
    return bool(compiled_pattern.search(line))


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
        return False
    compiled_per_mod = []
    for mod in target_mods:
        patterns = []
        for m in mod.get("matchers", []):
            pat = m.get("string", "")
            patterns.append(pattern_to_regex(pat))
        compiled_per_mod.append(patterns)

    hit_count = 0
    for patterns in compiled_per_mod:
        matched = False
        for p in patterns:
            for line in mod_lines:
                if mod_match_line(line, p):
                    matched = True
                    break
            if matched:
                break
        if matched:
            hit_count += 1

    if require_k is None:
        return hit_count == len(target_mods)
    else:
        return hit_count >= require_k

# ---------- IO: config / log ----------

def save_config(cfg):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def load_config():
    if not os.path.exists(CONFIG_FILE):
        return None
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"載入設定檔錯誤: {e}")
        return None


def append_log_line(line):
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line + "\n")

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


def do_click_sequence(alt_pos, cluster_pos, offset, click_delay, copy_delay):
    # 右鍵改造石
    click_with_offset(alt_pos[0], alt_pos[1], offset=offset, button="right")
    time.sleep(0.06)
    # 左鍵星團
    click_with_offset(cluster_pos[0], cluster_pos[1], offset=offset, button="left")
    time.sleep(click_delay)
    # Ctrl+C
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
        do_click_sequence(gui_vars['alt_pos'], gui_vars['cluster_pos'], gui_vars['offset'], gui_vars['click_delay'], gui_vars['copy_delay'])
        roll_count += 1
        gui_vars['set_count'](roll_count)
        clip = pyperclip.paste()
        mod_lines = extract_mod_section_from_clipboard(clip)
        hit = check_hit(mod_lines, gui_vars['targets'], require_k=gui_vars['require_k'])
        gui_vars['append_log'](f"[{roll_count}] 讀取 {len(mod_lines)} 行 -> {'HIT' if hit else 'MISS'}")
        append_log_line(f"{datetime.now().isoformat()} | #{roll_count} | HIT={hit} | lines={len(mod_lines)}")
        if hit:
            gui_vars['append_log']("命中條件，停止腳本。")
            stop_event.set()
            break
        time.sleep(gui_vars.get('loop_delay', 0.2))
    end_time = datetime.now()
    gui_vars['append_log']("腳本結束: " + end_time.strftime("%Y-%m-%d %H:%M:%S"))
    gui_vars['append_log'](f"總共洗了 {roll_count} 次。")


# ---------- GUI ----------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Cluster Washer (Press ESC to stop)")

        # ---- 載入詞綴資料 ----
        self.mods = load_mod_list(MOD_FILE)
        print(f"載入 {len(self.mods)} 個詞墜")

        # ---- 座標設定（預設值）----
        self.alt_pos = (100, 200)
        self.cluster_pos = (300, 400)

        # ---- 先載入設定檔到暫存 ----
        self.loaded_config = load_config()
        
        if self.loaded_config:
            self.alt_pos = tuple(self.loaded_config.get("alt_pos", self.alt_pos))
            self.cluster_pos = tuple(self.loaded_config.get("cluster_pos", self.cluster_pos))
            print(f"[DEBUG] 從設定檔載入座標: alt={self.alt_pos}, cluster={self.cluster_pos}")
        else:
            print(f"[DEBUG] 使用預設座標: alt={self.alt_pos}, cluster={self.cluster_pos}")

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

        # ---- Prefix / Suffix 暫存 ----
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

        # ---- 開始更新滑鼠座標 ----
        self.update_mouse_pos()

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
        ttk.Label(left, text="詞墜清單 (stats.ndjson)").pack(anchor="w")
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
        ttk.Label(right, text="前綴").pack(anchor="w")
        self.prefix_listbox = tk.Listbox(right, width=48, height=6)
        self.prefix_listbox.pack(fill=tk.X)
        ttk.Label(right, text="後綴").pack(anchor="w", pady=(6,0))
        self.suffix_listbox = tk.Listbox(right, width=48, height=6)
        self.suffix_listbox.pack(fill=tk.X)

        # ---- 滑鼠座標設定區 ----
        pos_frm = ttk.LabelFrame(right, text="滑鼠座標 (設定後會儲存)")
        pos_frm.pack(fill=tk.X, pady=6)
        
        # 改造石座標 Entry（綁定正確的變數）
        ttk.Label(pos_frm, text="改造石座標:").pack(anchor="w", padx=4, pady=(4,0))
        self.alt_entry = ttk.Entry(pos_frm, textvariable=self.alteration_pos_var)
        self.alt_entry.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(pos_frm, text="設定改造石座標 (錄當前滑鼠)", command=self.record_alt_pos).pack(fill=tk.X, padx=4, pady=2)
        
        # 星團座標 Entry（綁定正確的變數）
        ttk.Label(pos_frm, text="星團珠座標:").pack(anchor="w", padx=4, pady=(8,0))
        self.cluster_entry = ttk.Entry(pos_frm, textvariable=self.cluster_pos_var)
        self.cluster_entry.pack(fill=tk.X, padx=4, pady=2)
        ttk.Button(pos_frm, text="設定星團座標 (錄當前滑鼠)", command=self.record_cluster_pos).pack(fill=tk.X, padx=4, pady=2)
        
        ttk.Checkbutton(pos_frm, text="顯示即時滑鼠座標", variable=self.follow_mouse).pack(anchor="w", padx=4, pady=(8,0))
        self.mouse_pos_label = ttk.Label(pos_frm, text="鼠標: (0,0)")
        self.mouse_pos_label.pack(anchor="w", padx=4, pady=(2,4))

        # ---- 命中規則 ----
        mode_frm = ttk.LabelFrame(right, text="命中規則")
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
        offset_frm = ttk.LabelFrame(right, text="仿人物差-滑鼠偏移設定")
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
        self.prefix_listbox.insert(tk.END, desc)
        self.selected_prefix.append(mod)

    def add_suffix(self):
        sel = self.mods_listbox.curselection()
        if not sel or len(sel) == 0:
            return

        ui_idx = sel[0]
        real_idx = self.filtered_indices[ui_idx]
        mod = self.mods[real_idx]

        desc = mod.get("matchers", [{}])[0].get("string", mod.get("ref", f"mod{ui_idx}"))
        self.suffix_listbox.insert(tk.END, desc)
        self.selected_suffix.append(mod)

    def clear_prefix(self):
        self.prefix_listbox.delete(0, tk.END)
        self.selected_prefix = []

    def clear_suffix(self):
        self.suffix_listbox.delete(0, tk.END)
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
        self.mods = load_mod_list(MOD_FILE)
        self.cluster_affixes = self.mods[:]
        self.filtered_indices = list(range(len(self.mods)))
        print(f"重新載入 {len(self.mods)} 個詞墜")
        self.mods_listbox.delete(0, tk.END)
        for i, m in enumerate(self.mods):
            desc = m.get("matchers", [{}])[0].get("string", m.get("ref", f"mod{i}"))
            self.mods_listbox.insert(tk.END, f"{i+1}. {desc}")
        self.append_log("已重新載入詞檔。")

    def save_config(self):
        cfg = {
            "alt_pos": self.alt_pos,
            "cluster_pos": self.cluster_pos,
            "prefix_indices": [self.mods.index(m) for m in self.selected_prefix if m in self.mods],
            "suffix_indices": [self.mods.index(m) for m in self.selected_suffix if m in self.mods],
            "require_k_mode": self.require_k_mode.get(),
            "k_value": self.k_value.get(),
            "loop_delay": self.loop_delay.get(),
            "offset": self.offset.get(),
            "click_delay": float(self.click_delay.get()),
            "copy_delay": float(self.copy_delay.get())
        }
        save_config(cfg)
        self.append_log("設定已儲存。")
        print(f"[DEBUG] 儲存座標: alt={self.alt_pos}, cluster={self.cluster_pos}")

    def load_other_settings(self, cfg):
        """載入前後綴和其他設定（座標已在 __init__ 中載入）"""
        # prefix/suffix 填回
        try:
            prefix_idx = cfg.get("prefix_indices", [])
            for i in prefix_idx:
                if 0 <= i < len(self.mods):
                    m = self.mods[i]
                    desc = m.get("matchers", [{}])[0].get("string", m.get("ref", ""))
                    self.prefix_listbox.insert(tk.END, desc)
                    self.selected_prefix.append(m)
    
            suffix_idx = cfg.get("suffix_indices", [])
            for i in suffix_idx:
                if 0 <= i < len(self.mods):
                    m = self.mods[i]
                    desc = m.get("matchers", [{}])[0].get("string", m.get("ref", ""))
                    self.suffix_listbox.insert(tk.END, desc)
                    self.selected_suffix.append(m)
        except Exception as e:
            print(f"載入前後綴錯誤: {e}")
    
        # 其他設定
        try:
            self.require_k_mode.set(cfg.get("require_k_mode", "all"))
            self.k_value.set(cfg.get("k_value", 1))
            self.loop_delay.set(cfg.get("loop_delay", 0.2))
            self.offset.set(cfg.get("offset", 3))
            self.click_delay.set(cfg.get("click_delay", DELAY_AFTER_CLICK))
            self.copy_delay.set(cfg.get("copy_delay", DELAY_AFTER_COPY))
        except Exception as e:
            print(f"載入設定錯誤: {e}")
    
        self.append_log("已載入先前設定。")

    def open_log_file(self):
        if os.path.exists(LOG_FILE):
            os.startfile(LOG_FILE)
        else:
            messagebox.showinfo("Log", "尚未有 log。")

    def start_roll(self):
        if not self.selected_prefix and not self.selected_suffix:
            if not messagebox.askyesno("確認", "尚未選擇任何目標詞墜，要繼續嗎？"):
                return

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
            "targets": targets,
            "require_k": require_k,
            "loop_delay": float(self.loop_delay.get()),
            "append_log": self.append_log,
            "set_count": self.set_count,
            "offset": int(self.offset.get()),
            "click_delay": float(self.click_delay.get()),
            "copy_delay": float(self.copy_delay.get())
        }

        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.append_log("執行緒啟動中...")
        t = threading.Thread(target=worker_loop, args=(gui_vars,), daemon=True)
        t.start()

    def stop_roll(self):
        stop_event.set()
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
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