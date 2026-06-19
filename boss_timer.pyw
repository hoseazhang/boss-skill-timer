#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BOSS Skill Cooldown Timer — Game tool for tracking boss skill cycles.
Each skill counts down independently. Reaches 0 → voice alert → auto-restart.
Zero external dependencies. Windows EXE: pyinstaller --onefile --windowed boss_timer.pyw
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
import json, os, time, threading, platform, sys

APP  = "BOSS技能计时器"
VER  = "4.0"
IS_WIN = platform.system() == "Windows"

FONT   = ("Microsoft YaHei", 11) if IS_WIN else ("PingFang SC", 13)
FONT_B = (FONT[0], FONT[1], "bold")
FONT_S = (FONT[0], FONT[1] - 1)

# Color palette — dark theme, readable contrast
BG    = "#121218"
PANEL = "#1a1a24"
HEAD  = "#1e1e2e"
ROW   = "#161622"
ACC   = "#2a2a3a"
TX    = "#e8e8ec"
MUTED = "#8a8a9a"
DIM   = "#4a4a55"
GREEN = "#22c55e"
YELLOW= "#f59e0b"
RED   = "#ef4444"
BLUE  = "#3b82f6"


def data_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ============================================================
# Skill model
# ============================================================
class Skill:
    def __init__(self, name="技能", cd=30):
        self.name = name
        self.cd   = max(1, min(60, int(cd)))   # cooldown in seconds
        self.left = float(self.cd)              # remaining seconds
        self.run  = False                       # is counting down?

    def reset(self):
        self.left = float(self.cd)
        self.run  = False

    def to_dict(self):
        return {"name": self.name, "cooldown": self.cd}

    @classmethod
    def from_dict(cls, d):
        return cls(d.get("name", "技能"), d.get("cooldown", 30))


# ============================================================
# Voice engine — Windows SAPI / macOS say
# ============================================================
class Voice:
    def __init__(self):
        self.enabled = True
        self.rate   = 2
        self.volume = 100
        self._voice = None
        if IS_WIN:
            try:
                import win32com.client
                self._voice = win32com.client.Dispatch("SAPI.SpVoice")
            except Exception as e:
                print(f"Voice init failed (win32com may not be installed): {e}")

    def speak(self, text):
        if not self.enabled or not text:
            return
        def _do():
            try:
                if IS_WIN and self._voice:
                    self._voice.Rate   = self.rate
                    self._voice.Volume = self.volume
                    self._voice.Speak(text, 1)
                elif platform.system() == "Darwin":
                    os.system(f'say -v Tingting "{text}" &')
            except:
                pass
        threading.Thread(target=_do, daemon=True).start()


# ============================================================
# Timer engine — tracks all skills
# ============================================================
class Timer:
    def __init__(self, tick=0.1):
        self.tick    = tick
        self.skills  = []
        self.state   = "idle"   # idle | running | paused
        self._lock   = threading.Lock()
        self._active = False
        self._thread = None

        # callbacks set by UI
        self.on_frame = lambda: None
        self.on_done  = lambda skill: None
        self.on_state = lambda st: None

    # ---- skill CRUD ----
    def load(self, skills):
        with self._lock:
            self.skills = [Skill(s.name, s.cd) for s in skills]

    def add(self, name="新技能", cd=30):
        with self._lock:
            s = Skill(name, cd)
            self.skills.append(s)
            return len(self.skills) - 1

    def remove(self, i):
        with self._lock:
            if 0 <= i < len(self.skills):
                self.skills.pop(i)

    def update(self, i, name=None, cd=None):
        with self._lock:
            if 0 <= i < len(self.skills):
                s = self.skills[i]
                if name: s.name = name
                if cd:   s.cd = max(1, min(60, int(cd))); s.left = float(s.cd)

    # ---- global controls ----
    def start_all(self):
        with self._lock:
            if self.state == "running":
                return
            self.state = "running"
            for s in self.skills:
                s.run = True
                if s.left <= 0:
                    s.left = float(s.cd)
            self._start_loop()
            self.on_state("running")

    def pause_all(self):
        with self._lock:
            if self.state != "running":
                return
            self.state = "paused"
            self._stop_loop()
            for s in self.skills:
                s.run = False
            self.on_state("paused")

    def resume_all(self):
        with self._lock:
            if self.state != "paused":
                return
            self.state = "running"
            for s in self.skills:
                s.run = True
            self._start_loop()
            self.on_state("running")

    def reset_all(self):
        with self._lock:
            self._stop_loop()
            self.state = "idle"
            for s in self.skills:
                s.reset()
            self.on_state("idle")
            self.on_frame()

    # ---- per-skill controls ----
    def start_one(self, i):
        with self._lock:
            if 0 <= i < len(self.skills):
                s = self.skills[i]
                s.run = True
                if s.left <= 0:             # only reset if expired
                    s.left = float(s.cd)
                self._start_loop()

    def pause_one(self, i):
        with self._lock:
            if 0 <= i < len(self.skills):
                self.skills[i].run = False

    def reset_one(self, i):
        with self._lock:
            if 0 <= i < len(self.skills):
                s = self.skills[i]
                s.left = float(s.cd)
                if self.state == "running":
                    s.run = True
                else:
                    s.run = False

    # ---- internal loop ----
    def _start_loop(self):
        if self._active:
            return
        self._active = True
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _stop_loop(self):
        self._active = False

    def _run(self):
        while self._active:
            t0 = time.time()
            with self._lock:
                fired = []
                for i, s in enumerate(self.skills):
                    if s.run:
                        s.left = max(0.0, s.left - self.tick)
                        if s.left <= 0.0:
                            s.left = 0.0
                            fired.append((i, s))
                # Fire completions → auto-restart
                for i, s in fired:
                    self.on_done(s)          # voice alert
                    s.left = float(s.cd)      # reset cooldown
                    s.run  = True             # keep running (auto-loop)
                self.on_frame()
                # Auto-stop if nothing running
                if self.state != "running" and not any(s.run for s in self.skills):
                    self._active = False
                    break
            elapsed = time.time() - t0
            time.sleep(max(0.0, self.tick - elapsed))
        self._active = False

    def stop(self):
        self._stop_loop()


# ============================================================
# ── Skill row widget ──
# ============================================================
class SkillRow(tk.Frame):
    def __init__(self, parent, index, skill, cb):
        super().__init__(parent, bg=ROW, height=44)
        self.idx = index
        self.sk  = skill
        self.cb  = cb
        self.pack(fill=tk.X, padx=2, pady=1)

        # Grid layout: 5 columns
        # col 0: name (expand slightly)  col 1: cd  col 2: bar (expand most)
        # col 3: time  col 4: buttons
        self.grid_columnconfigure(0, weight=1)   # name
        self.grid_columnconfigure(1, weight=0)   # cd
        self.grid_columnconfigure(2, weight=3)   # progress bar
        self.grid_columnconfigure(3, weight=0)   # time
        self.grid_columnconfigure(4, weight=0)   # buttons

        # name
        self.lb_name = tk.Label(self, text=self.sk.name, font=FONT_B,
                                bg=ROW, fg=TX, anchor=tk.W)
        self.lb_name.grid(row=0, column=0, padx=(8,4), pady=6, sticky=tk.W)

        # cd label
        self.lb_cd = tk.Label(self, text=f"{self.sk.cd}s", font=FONT,
                              bg=ROW, fg=MUTED, cursor="hand2")
        self.lb_cd.grid(row=0, column=1, padx=2, pady=6)
        self.lb_cd.bind("<Double-Button-1>", lambda e: self.cb["edit_cd"](self.idx))

        # progress bar canvas
        self.cv = tk.Canvas(self, bg=BG, height=16, highlightthickness=0)
        self.cv.grid(row=0, column=2, padx=6, pady=6, sticky=tk.EW)
        self._bar_fill = None   # filled rect item ID
        self._bar_bg   = None   # background rect item ID

        # time
        self.lb_time = tk.Label(self, text=f"{self.sk.cd:.1f}", font=FONT_B,
                                bg=ROW, fg=GREEN)
        self.lb_time.grid(row=0, column=3, padx=4, pady=6)

        # button frame (start + reset)
        bf = tk.Frame(self, bg=ROW)
        bf.grid(row=0, column=4, padx=(4,8), pady=4, sticky=tk.E)

        self.btn_start = tk.Label(bf, text="▶", font=FONT_S,
                                  cursor="hand2", bg=GREEN, fg="#fff",
                                  width=2, padx=4, pady=1)
        self.btn_start.pack(side=tk.LEFT, padx=1)
        self.btn_start.bind("<Button-1>", lambda e: self.cb["start"](self.idx))

        self.btn_reset = tk.Label(bf, text="🔄", font=FONT,
                                  cursor="hand2", bg=ROW, fg=RED, width=2)
        self.btn_reset.pack(side=tk.LEFT, padx=1)
        self.btn_reset.bind("<Button-1>", lambda e: self.cb["reset"](self.idx))

    def refresh(self, skill):
        """Only update dynamic parts: bar, time, button, row color"""
        self.sk = skill
        s = skill

        # time + bar — only if running
        if not s.run:
            return  # nothing dynamic to update

        r = max(0.0, s.left)
        ratio = r / s.cd if s.cd > 0 else 0.0
        if r <= 3:    color = RED
        elif r <= 8:  color = YELLOW
        else:         color = GREEN
        self.lb_time.config(text=f"{r:.1f}", fg=color)
        self._bar(ratio, color)

        # start button state
        self.btn_start.config(text="⏸", bg="#5c3d1a", fg=YELLOW)

        # row highlight
        if s.left <= 3:
            bg = "#1a1010"
            self.config(bg=bg)
            self.lb_name.config(bg=bg)
            self.lb_time.config(bg=bg)
            self.btn_start.master.config(bg=bg)

    def refresh_static(self, skill):
        """Update name, cd, button initial state after config changes"""
        self.sk = skill; s = skill
        self.lb_name.config(text=s.name)
        self.lb_cd.config(text=f"{s.cd}s")
        if s.run:
            self.lb_time.config(text=f"{s.left:.1f}")
            self.btn_start.config(text="⏸", bg="#5c3d1a", fg=YELLOW)
        else:
            self.lb_time.config(text=f"{s.cd:.1f}", fg=GREEN)
            self._bar(1.0, GREEN)
            self.btn_start.config(text="▶", bg=GREEN, fg="#fff")
        self.config(bg=ROW)
        self.lb_name.config(bg=ROW)
        self.lb_time.config(bg=ROW)
        self.btn_start.master.config(bg=ROW)

    def _bar(self, ratio, color):
        """Update progress bar — no flicker, uses coords()"""
        w = self.cv.winfo_width()
        h = self.cv.winfo_height()
        if w < 4 or h < 2:
            return

        fw = int(w * max(0.0, min(1.0, ratio)))

        # Create rects if first time
        if self._bar_bg is None:
            self._bar_bg = self.cv.create_rectangle(0, 0, w, h,
                                                     fill=BG, outline="")
        if self._bar_fill is None:
            self._bar_fill = self.cv.create_rectangle(0, 0, max(fw, 0), h,
                                                       fill=color, outline="")

        # Update coords (no delete → no flicker)
        self.cv.coords(self._bar_bg, 0, 0, w, h)
        self.cv.coords(self._bar_fill, 0, 0, max(fw, 1), h)
        self.cv.itemconfig(self._bar_fill, fill=color)


# ============================================================
# Config dialog — edit skills, load/save presets
# ============================================================
class ConfigDialog(tk.Toplevel):
    def __init__(self, parent, timer, base_dir):
        super().__init__(parent)
        self.timer    = timer
        self.base_dir = base_dir
        self.presets_dir = os.path.join(base_dir, "presets")
        os.makedirs(self.presets_dir, exist_ok=True)

        self.title("配置 BOSS 方案")
        self.geometry("480x420")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent)

        self._skills  = []
        self._cur_file = None
        self._build()
        self._load_presets()

        self.update_idletasks()
        x = parent.winfo_x() + parent.winfo_width() // 2 - 240
        y = parent.winfo_y() + 20
        self.geometry(f"+{x}+{y}")
        self.grab_set()

    def _build(self):
        # preset selector
        top = tk.Frame(self, bg=PANEL)
        top.pack(fill=tk.X, padx=10, pady=(10, 0))

        tk.Label(top, text="BOSS方案:", font=FONT_S, fg=MUTED, bg=PANEL).pack(side=tk.LEFT, padx=4, pady=6)

        self.file_var = tk.StringVar()
        self.file_menu = ttk.Combobox(top, textvariable=self.file_var,
                                      font=FONT_S, width=18, state="readonly")
        self.file_menu.pack(side=tk.LEFT, padx=4)
        self.file_menu.bind("<<ComboboxSelected>>", lambda e: self._on_select())

        self.name_var = tk.StringVar()
        tk.Entry(top, textvariable=self.name_var, font=FONT_S,
                 bg=ACC, fg=TX, insertbackground=TX, width=12).pack(side=tk.LEFT, padx=4)

        # skill list
        mid = tk.Frame(self, bg=BG)
        mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)

        cv = tk.Canvas(mid, bg=BG, highlightthickness=0)
        sb = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=cv.yview)
        self.inner = tk.Frame(cv, bg=BG)
        self.inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=self.inner, anchor=tk.NW)
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # header
        hf = tk.Frame(self.inner, bg=HEAD)
        hf.pack(fill=tk.X, pady=2)
        for t, w in [("技能名称", 18), ("冷却(秒)", 8), ("", 6)]:
            tk.Label(hf, text=t, font=FONT_S, fg=MUTED, bg=HEAD, width=w).pack(side=tk.LEFT, padx=2)

        self._row_frames = []

        # buttons
        bot = tk.Frame(self, bg=PANEL)
        bot.pack(fill=tk.X, padx=10, pady=8)

        self._btn(bot, "+ 添加技能", BLUE,  self._add).pack(side=tk.LEFT, padx=3)
        self._btn(bot, "💾 保存并应用", GREEN, self._save).pack(side=tk.LEFT, padx=3)
        self._btn(bot, "🗑 删除方案", RED,   self._delete).pack(side=tk.LEFT, padx=3)
        self._btn(bot, "关闭", ACC,  self.destroy).pack(side=tk.RIGHT, padx=3)

    def _btn(self, parent, text, color, cmd):
        b = tk.Label(parent, text=text, font=FONT_S, cursor="hand2",
                     bg=color, fg="#fff", padx=8, pady=2)
        b.bind("<Button-1>", lambda e: cmd())
        return b

    def _load_presets(self):
        files = sorted(f for f in os.listdir(self.presets_dir) if f.endswith(".json"))
        self.file_menu["values"] = files
        if files:
            self.file_var.set(files[0])
            self._on_select()

    def _on_select(self):
        fn = self.file_var.get()
        if not fn:
            return
        path = os.path.join(self.presets_dir, fn)
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.name_var.set(data.get("name", ""))
        self._skills = [Skill.from_dict(s) for s in data.get("skills", [])]
        self._cur_file = fn
        self._show()

    def _show(self):
        for fr in self._row_frames:
            fr.destroy()
        self._row_frames.clear()

        for i, s in enumerate(self._skills):
            fr = tk.Frame(self.inner, bg=ROW)
            fr.pack(fill=tk.X, padx=2, pady=1)
            self._row_frames.append(fr)

            sv = tk.StringVar(value=s.name)
            tk.Entry(fr, textvariable=sv, font=FONT_S,
                     bg=ACC, fg=TX, insertbackground=TX, width=16).pack(side=tk.LEFT, padx=2, pady=2)

            cv_ = tk.IntVar(value=s.cd)
            tk.Spinbox(fr, textvariable=cv_, from_=1, to=60, font=FONT_S,
                       bg=ACC, fg=TX, width=4, buttonbackground=ACC).pack(side=tk.LEFT, padx=2, pady=2)

            tk.Label(fr, text="✕", font=FONT, cursor="hand2", bg=ROW,
                     fg=RED, width=2).pack(side=tk.LEFT, padx=(8, 2))
            fr.winfo_children()[-1].bind("<Button-1>", lambda e, idx=i: self._del(idx))

            fr._name_var = sv
            fr._cd_var   = cv_

    def _add(self):
        self._skills.append(Skill("新技能", 30))
        self._show()

    def _del(self, idx):
        if 0 <= idx < len(self._skills):
            self._skills.pop(idx)
            self._show()

    def _save(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入 BOSS 名称", parent=self)
            return

        # Read values from UI
        for i, fr in enumerate(self._row_frames):
            if i < len(self._skills):
                s = self._skills[i]
                nm = fr._name_var.get().strip()
                if nm:
                    s.name = nm
                try:
                    s.cd = max(1, min(60, fr._cd_var.get()))
                except:
                    pass

        fn = f"{name}.json"
        data = {"name": name, "skills": [s.to_dict() for s in self._skills]}
        path = os.path.join(self.presets_dir, fn)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        # Apply to timer
        self.timer.reset_all()
        self.timer.load(self._skills)

        # Refresh preset list
        files = sorted(f for f in os.listdir(self.presets_dir) if f.endswith(".json"))
        self.file_menu["values"] = files
        self.file_var.set(fn)

        self.destroy()

    def _delete(self):
        fn = self.file_var.get()
        if not fn:
            return
        if messagebox.askyesno("确认", f"删除方案「{fn}」？", parent=self):
            path = os.path.join(self.presets_dir, fn)
            if os.path.exists(path):
                os.remove(path)
            files = sorted(f for f in os.listdir(self.presets_dir) if f.endswith(".json"))
            self.file_menu["values"] = files
            if files:
                self.file_var.set(files[0])
                self._on_select()
            else:
                self.file_var.set("")
                self._skills = []
                self._show()


# ============================================================
# Settings dialog
# ============================================================
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, voice, cb_topmost):
        super().__init__(parent)
        self.voice = voice
        self.title("设置")
        self.geometry("380x300")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent)
        self._build(cb_topmost)

        self.update_idletasks()
        x = parent.winfo_x() + parent.winfo_width() // 2 - 190
        y = parent.winfo_y() + 30
        self.geometry(f"+{x}+{y}")

    def _build(self, cb_topmost):
        pad = {"padx": 16, "pady": 6}

        tk.Label(self, text="⚙ 设置", font=FONT_B, fg=TX, bg=BG).pack(anchor=tk.W, **pad)

        # Voice
        vf = tk.LabelFrame(self, text="🔊 语音", font=FONT_S, fg=MUTED, bg=BG, padx=12, pady=6)
        vf.pack(fill=tk.X, padx=14, pady=4)

        ve = tk.BooleanVar(value=self.voice.enabled)
        tk.Checkbutton(vf, text="启用语音播报", variable=ve, font=FONT_S, fg=TX, bg=BG,
                       selectcolor=BG, activebackground=BG,
                       command=lambda: setattr(self.voice, 'enabled', ve.get())).pack(anchor=tk.W)

        rf = tk.Frame(vf, bg=BG); rf.pack(fill=tk.X, pady=2)
        tk.Label(rf, text="语速:", font=FONT_S, fg=MUTED, bg=BG, width=5).pack(side=tk.LEFT)
        rv = tk.IntVar(value=self.voice.rate)
        tk.Scale(rf, from_=-10, to=10, orient=tk.HORIZONTAL, variable=rv, bg=BG, fg=TX,
                 highlightthickness=0, troughcolor=ACC,
                 command=lambda v: setattr(self.voice, 'rate', int(float(v)))).pack(side=tk.LEFT, fill=tk.X, expand=True)

        rf2 = tk.Frame(vf, bg=BG); rf2.pack(fill=tk.X, pady=2)
        tk.Label(rf2, text="音量:", font=FONT_S, fg=MUTED, bg=BG, width=5).pack(side=tk.LEFT)
        vv = tk.IntVar(value=self.voice.volume)
        tk.Scale(rf2, from_=0, to=100, orient=tk.HORIZONTAL, variable=vv, bg=BG, fg=TX,
                 highlightthickness=0, troughcolor=ACC,
                 command=lambda v: setattr(self.voice, 'volume', int(float(v)))).pack(side=tk.LEFT, fill=tk.X, expand=True)

        tk.Button(vf, text="📢 测试语音", font=FONT_S, bg=ACC, fg=TX, relief=tk.FLAT,
                  cursor="hand2", command=lambda: self.voice.speak(f"{APP} 语音测试")).pack(anchor=tk.W, pady=4)

        # Window
        df = tk.LabelFrame(self, text="🖥 窗口", font=FONT_S, fg=MUTED, bg=BG, padx=12, pady=6)
        df.pack(fill=tk.X, padx=14, pady=4)

        tv = tk.BooleanVar(value=True)
        tk.Checkbutton(df, text="窗口置顶", variable=tv, font=FONT_S, fg=TX, bg=BG,
                       selectcolor=BG, activebackground=BG,
                       command=lambda: cb_topmost(tv.get())).pack(anchor=tk.W)


# ============================================================
# Main application
# ============================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP)
        self.geometry("820x500")
        self.minsize(560, 300)
        self.configure(bg=BG)

        self.base  = data_dir()
        self.timer = Timer(0.25)  # 4 fps — smooth enough, no flicker
        self.voice = Voice()

        self.timer.on_frame = lambda: self.after(0, self._refresh)
        self.timer.on_done  = lambda s: self.after(0, lambda: self.voice.speak(s.name))
        self.timer.on_state = lambda s: self.after(0, lambda: self._on_state(s))

        self.rows      = []
        self.cur_file  = None
        self._topmost  = True
        self._refresh_pending = False

        self._build()
        self._init_data()
        self.protocol("WM_DELETE_WINDOW", self._close)

    # ---- build UI ----
    def _build(self):
        # title bar — no fixed height, let content determine
        top = tk.Frame(self, bg=HEAD)
        top.pack(fill=tk.X)

        self.lb_title = tk.Label(top, text=f"🔥 {APP}", font=FONT_B,
                                 fg=YELLOW, bg=HEAD)
        self.lb_title.pack(side=tk.LEFT, padx=12, pady=6)

        tr = tk.Frame(top, bg=HEAD); tr.pack(side=tk.RIGHT, padx=8, pady=6)
        self.btn_pin = tk.Label(tr, text="📌", font=("", 14), cursor="hand2",
                                bg=HEAD, fg=GREEN, width=3)
        self.btn_pin.pack(side=tk.LEFT)
        self.btn_pin.bind("<Button-1>", lambda e: self._toggle_topmost())

        self.btn_voice = tk.Label(tr, text="🔊", font=("", 14), cursor="hand2",
                                  bg=HEAD, fg=GREEN, width=3)
        self.btn_voice.pack(side=tk.LEFT)
        self.btn_voice.bind("<Button-1>", lambda e: self._toggle_voice())

        self.btn_cfg = tk.Label(tr, text="⚙", font=("", 14), cursor="hand2",
                                bg=HEAD, fg=MUTED, width=3)
        self.btn_cfg.pack(side=tk.LEFT)
        self.btn_cfg.bind("<Button-1>", lambda e: self._open_settings())

        # Thin separator
        tk.Frame(self, bg=ACC, height=1).pack(fill=tk.X)

        # column header — grid for alignment
        hf = tk.Frame(self, bg=BG)
        hf.pack(fill=tk.X, padx=4, pady=(2, 0))
        hf.grid_columnconfigure(0, weight=1)
        hf.grid_columnconfigure(1, weight=0)
        hf.grid_columnconfigure(2, weight=3)
        hf.grid_columnconfigure(3, weight=0)
        hf.grid_columnconfigure(4, weight=0)
        for col, (t, pad) in enumerate([
            ("技能名称", (8, 4)), ("冷却", (2,)), ("倒计时进度", (6,)), ("剩余", (4,)), ("操作", (4,8))
        ]):
            tk.Label(hf, text=t, font=FONT_S, fg=MUTED, bg=BG).grid(
                row=0, column=col, padx=pad, pady=2, sticky=tk.W)

        # scrollable skill list
        self.canvas = tk.Canvas(self, bg=BG, highlightthickness=0)
        self.scroll = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.list_fr = tk.Frame(self.canvas, bg=BG)
        self.list_fr.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.list_fr, anchor=tk.NW)
        self.canvas.configure(yscrollcommand=self.scroll.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
        self.scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.bind_all("<MouseWheel>",
            lambda e: self.canvas.yview_scroll(int(-e.delta / 120), "units"))

        # bottom bar — no fixed height
        bot = tk.Frame(self, bg=PANEL)
        bot.pack(fill=tk.X, side=tk.BOTTOM, padx=4, pady=4)

        bl = tk.Frame(bot, bg=PANEL); bl.pack(side=tk.LEFT, padx=8, pady=8)
        self._btn_label(bl, "🔧 配置", BLUE, self._open_config).pack(side=tk.LEFT, padx=3)
        self._btn_label(bl, "💾 保存", ACC,  self._save).pack(side=tk.LEFT, padx=3)

        br = tk.Frame(bot, bg=PANEL); br.pack(side=tk.RIGHT, padx=8, pady=8)
        self.btn_start = self._btn_label(br, "▶ 全部开始", GREEN, self._start)
        self.btn_start.pack(side=tk.LEFT, padx=3)
        self._btn_label(br, "⏸ 暂停", "#5c3d1a", self.timer.pause_all).pack(side=tk.LEFT, padx=3)
        self._btn_label(br, "🔄 重置", "#5c1a1a", self.timer.reset_all).pack(side=tk.LEFT, padx=3)

        self.lb_state = tk.Label(bot, text="🟢 就绪", font=FONT_S, bg=PANEL, fg=MUTED)
        self.lb_state.pack(side=tk.RIGHT, padx=10, pady=8)

    def _btn_label(self, parent, text, color, cmd):
        b = tk.Label(parent, text=text, font=FONT_S, cursor="hand2",
                     bg=color, fg="#fff", padx=10, pady=3)
        b.bind("<Button-1>", lambda e: cmd())
        return b

    # ---- data init ----
    def _init_data(self):
        pd = os.path.join(self.base, "presets")
        os.makedirs(pd, exist_ok=True)

        demo_path = os.path.join(pd, "demo_boss.json")
        if not os.path.exists(demo_path):
            self._write_json(demo_path, {
                "name": "炎狱之王",
                "skills": [
                    {"name": "火焰吐息", "cooldown": 30},
                    {"name": "暗影冲锋", "cooldown": 20},
                    {"name": "陨石坠落", "cooldown": 90},
                    {"name": "冰霜禁锢", "cooldown": 20},
                ]
            })

        tpl_path = os.path.join(pd, "new_boss_template.json")
        if not os.path.exists(tpl_path):
            self._write_json(tpl_path, {
                "name": "新BOSS",
                "skills": [{"name": "技能1", "cooldown": 30}, {"name": "技能2", "cooldown": 20}]
            })

        # Load demo preset
        data = self._read_json(demo_path)
        if data:
            skills = [Skill.from_dict(s) for s in data.get("skills", [])]
            self.timer.load(skills)
            self.lb_title.config(text=f"🔥 {data.get('name', APP)}")
            self.cur_file = "demo_boss.json"

        self._rebuild()
        self._refresh()

    def _write_json(self, path, data):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def _read_json(self, path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)

    # ---- row management ----
    def _rebuild(self):
        for r in self.rows:
            r.destroy()
        self.rows.clear()

        cb = {
            "start":    self._on_start,
            "reset":    self._on_reset,
            "edit_cd":  self._on_edit_cd,
        }
        for i, s in enumerate(self.timer.skills):
            self.rows.append(SkillRow(self.list_fr, i, s, cb))

        self.list_fr.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _refresh(self):
        if self._refresh_pending:
            return
        self._refresh_pending = True
        # Only update rows with running skills (bars + time only)
        for i, row in enumerate(self.rows):
            if i < len(self.timer.skills):
                s = self.timer.skills[i]
                if s.run:  # only refresh active timers
                    row.refresh(s)
        self._refresh_pending = False

    # ---- state ----
    def _on_state(self, st):
        if st == "running":
            self.lb_state.config(text="🔴 运行中", fg=RED)
            self.btn_start.config(text="⏸ 全部暂停", bg="#5c3d1a")
        elif st == "paused":
            self.lb_state.config(text="🟡 已暂停", fg=YELLOW)
            self.btn_start.config(text="▶ 继续", bg=GREEN)
        else:
            self.lb_state.config(text="🟢 就绪", fg=GREEN)
            self.btn_start.config(text="▶ 全部开始", bg=GREEN)

    def _start(self):
        s = self.timer.state
        if s == "running":
            self.timer.pause_all()
        elif s == "paused":
            self.timer.resume_all()
        else:
            self.timer.start_all()

    # ---- per-skill events ----
    def _on_start(self, i):
        if i < len(self.timer.skills):
            s = self.timer.skills[i]
            if s.run:
                self.timer.pause_one(i)
                # Use static refresh to update button
                if i < len(self.rows):
                    self.rows[i].refresh_static(s)
            else:
                self.timer.start_one(i)
                # running refresh will handle it

    def _on_reset(self, i):
        self.timer.reset_one(i)
        if i < len(self.rows) and i < len(self.timer.skills):
            self.rows[i].refresh_static(self.timer.skills[i])

    def _on_edit_cd(self, i):
        if i < len(self.timer.skills):
            s = self.timer.skills[i]
            v = simpledialog.askinteger("编辑冷却", f"「{s.name}」冷却时间 (1-60秒):",
                                        initialvalue=s.cd, minvalue=1, maxvalue=60, parent=self)
            if v:
                self.timer.update(i, cd=v)
                self._refresh()

    # ---- save ----
    def _save(self):
        if not self.timer.skills:
            return
        n = simpledialog.askstring("保存方案", "BOSS 名称:", parent=self)
        if n and n.strip():
            n = n.strip()
            pd = os.path.join(self.base, "presets")
            os.makedirs(pd, exist_ok=True)
            path = os.path.join(pd, f"{n}.json")
            self._write_json(path, {
                "name": n,
                "skills": [s.to_dict() for s in self.timer.skills]
            })
            self.lb_title.config(text=f"🔥 {n}")
            self.cur_file = f"{n}.json"

    # ---- config dialog ----
    def _open_config(self):
        ConfigDialog(self, self.timer, self.base)
        self._rebuild()
        self._refresh()

    # ---- settings ----
    def _open_settings(self):
        SettingsDialog(self, self.voice, self._set_topmost)

    # ---- toggles ----
    def _toggle_topmost(self):
        self._topmost = not self._topmost
        self._set_topmost(self._topmost)

    def _set_topmost(self, on):
        self._topmost = on
        self.btn_pin.config(fg=GREEN if on else MUTED)
        try:
            self.attributes('-topmost', on)
        except:
            pass

    def _toggle_voice(self):
        self.voice.enabled = not self.voice.enabled
        self.btn_voice.config(
            fg=GREEN if self.voice.enabled else DIM,
            text="🔊" if self.voice.enabled else "🔇")

    def _close(self):
        self.timer.stop()
        self.destroy()


if __name__ == "__main__":
    app = App()
    app.mainloop()
