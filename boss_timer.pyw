#!/usr/bin/env pythonw
# -*- coding: utf-8 -*-
"""BOSS技能倒计时器 v2.0 - 零外部依赖 | pip install pyinstaller && pyinstaller --onefile --windowed boss_timer.pyw"""
import tkinter as tk
from tkinter import messagebox, simpledialog, filedialog
import json, os, time, threading, ctypes, sys, platform

FONT = ("Microsoft YaHei", 11)
FONT_B = ("Microsoft YaHei", 11, "bold")
FONT_S = ("Microsoft YaHei", 10)
C = {"bg":"#0f0f1a","panel":"#161630","header":"#1a1a35","row":"#12122a","row_warn":"#1a0f0f","border":"#2a2a4a","text":"#e5e5e5","muted":"#a0a0a0","dim":"#555","green":"#22c55e","yellow":"#f59e0b","red":"#ef4444","blue":"#3b82f6","purple":"#8b5cf6"}
DIR_ARROW = {"up":"▲","down":"▼","left":"◀","right":"▶","none":"—"}
DIR_CN = {"up":"向上","down":"向下","left":"向左","right":"向右","none":""}
DIR_CLR = {"up":C["yellow"],"down":C["blue"],"left":C["green"],"right":C["red"],"none":C["dim"]}
DIR_BG = {"up":"#3d2a0a","down":"#0a1a3d","left":"#0a2a0a","right":"#3d0a0a","none":"#1a1a1a"}
IS_WIN = platform.system() == "Windows"
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in dir() else os.getcwd()
PRESETS_DIR = os.path.join(BASE_DIR, "presets")

# ===================== 数据模型 =====================
class Skill:
    def __init__(self, name="技能", cooldown=30, direction="none", enabled=True):
        self.name = name; self.cooldown = cooldown; self.direction = direction
        self.enabled = enabled; self.remaining = float(cooldown); self.running = False

class BossPreset:
    def __init__(self, name="", skills=None, notes="", game=""):
        self.name = name; self.skills = skills or []; self.notes = notes; self.game = game

# ===================== 语音引擎 =====================
class VoiceEngine:
    def __init__(self):
        self.enabled = True; self.rate = 2; self.volume = 100; self.early_warning = 0
        self._voice = None
        if IS_WIN:
            try:
                import win32com.client
                self._voice = win32com.client.Dispatch("SAPI.SpVoice")
            except: pass
    def speak(self, text):
        if not self.enabled or not text: return
        def _do():
            try:
                if self._voice:
                    self._voice.Rate = self.rate; self._voice.Volume = self.volume
                    self._voice.Speak(text, 1)  # async
                elif platform.system() == "Darwin":
                    os.system(f'say -v Tingting "{text}" &')
            except: pass
        threading.Thread(target=_do, daemon=True).start()
    def speak_skill(self, name, direction="none"):
        d = DIR_CN.get(direction, "")
        self.speak(f"{name} {d}" if d else name)
    def speak_warning(self, name, remaining):
        self.speak(f"准备 {name}")
    def test_speak(self): self.speak("BOSS技能倒计时器 语音测试")

# ===================== 方案管理 =====================
class PresetManager:
    def __init__(self):
        os.makedirs(PRESETS_DIR, exist_ok=True)
        self._ensure_defaults()
    def _ensure_defaults(self):
        if not os.path.exists(os.path.join(PRESETS_DIR, "demo_boss.json")):
            demo = BossPreset("炎狱之王", game="演示副本", notes="P1技能轴",
                skills=[Skill("火焰吐息",30,"up"), Skill("暗影冲锋",20,"left"),
                        Skill("陨石坠落",90,"down"), Skill("冰霜禁锢",20,"none")])
            self.save(demo, "demo_boss.json")
        if not os.path.exists(os.path.join(PRESETS_DIR, "new_boss_template.json")):
            tpl = BossPreset("新BOSS", skills=[Skill("技能1",30), Skill("技能2",20)])
            self.save(tpl, "new_boss_template.json")
    def save(self, preset, filename=None):
        fp = os.path.join(PRESETS_DIR, filename or f"{preset.name}.json")
        data = {"name":preset.name,"game":preset.game,"notes":preset.notes,
            "skills":[{"name":s.name,"cooldown":s.cooldown,"direction":s.direction,"enabled":s.enabled} for s in preset.skills]}
        with open(fp,'w',encoding='utf-8') as f: json.dump(data,f,ensure_ascii=False,indent=2)
    def load(self, filename):
        fp = os.path.join(PRESETS_DIR, filename)
        if not os.path.exists(fp): return None
        with open(fp,'r',encoding='utf-8') as f:
            d = json.load(f)
        return BossPreset(d.get("name",""), d.get("game",""), d.get("notes",""),
            [Skill(s["name"],s["cooldown"],s.get("direction","none"),s.get("enabled",True)) for s in d.get("skills",[])])
    def list(self):
        return sorted([f for f in os.listdir(PRESETS_DIR) if f.endswith('.json')]) if os.path.exists(PRESETS_DIR) else []
    def delete(self, filename):
        fp = os.path.join(PRESETS_DIR, filename)
        if os.path.exists(fp): os.remove(fp)

# ===================== 倒计时引擎 =====================
class TimerEngine:
    def __init__(self, interval=0.1):
        self.skills = []; self.state = "idle"; self.interval = interval
        self._running = False; self._thread = None; self._lock = threading.Lock()
        self._warn_tracker = {}
        self.on_tick = lambda s: None; self.on_complete = lambda i,s: None
        self.on_state = lambda s: None
    def set_skills(self, skills):
        with self._lock:
            self.skills = [Skill(s.name, s.cooldown, s.direction, s.enabled) for s in skills]
    def add_skill(self, name="新技能", cd=30, d="none"):
        with self._lock:
            s = Skill(name, cd, d); self.skills.append(s); return len(self.skills)-1
    def remove_skill(self, i):
        with self._lock:
            if 0 <= i < len(self.skills): self.skills.pop(i)
    def start(self):
        with self._lock:
            if self.state == "running": return
            self.state = "running"
            for s in self.skills:
                if s.enabled: s.running = True; s.remaining = s.remaining or float(s.cooldown)
            self._warn_tracker.clear()
            self._start_loop(); self.on_state("running")
    def pause(self):
        with self._lock:
            if self.state != "running": return
            self.state = "paused"; self._stop_loop()
            for s in self.skills: s.running = False
            self.on_state("paused")
    def resume(self):
        with self._lock:
            if self.state != "paused": return
            self.state = "running"
            for s in self.skills:
                if s.enabled and not s.running: s.running = True
            self._warn_tracker.clear(); self._start_loop(); self.on_state("running")
    def reset(self):
        with self._lock:
            self._stop_loop(); self.state = "idle"; self._warn_tracker.clear()
            for s in self.skills: s.remaining = float(s.cooldown); s.running = False
            self.on_state("idle"); self.on_tick(self.skills)
    def toggle_skill(self, i):
        with self._lock:
            if 0 <= i < len(self.skills):
                s = self.skills[i]
                if not s.enabled: return
                s.running = not s.running
                if s.running and s.remaining <= 0: s.remaining = float(s.cooldown)
    def reset_skill(self, i):
        with self._lock:
            if 0 <= i < len(self.skills):
                s = self.skills[i]; s.remaining = float(s.cooldown)
                s.running = (self.state == "running")
                self._warn_tracker.pop(i, None)
    def toggle_enabled(self, i):
        with self._lock:
            if 0 <= i < len(self.skills):
                s = self.skills[i]; s.enabled = not s.enabled
                if not s.enabled: s.running = False; s.remaining = float(s.cooldown)
                elif self.state == "running": s.running = True; s.remaining = float(s.cooldown)
    def _start_loop(self):
        if self._running: return
        self._running = True; self._thread = threading.Thread(target=self._loop, daemon=True); self._thread.start()
    def _stop_loop(self): self._running = False
    def _loop(self):
        while self._running:
            t0 = time.time()
            with self._lock:
                if self.state != "running": break
                completed = []
                for i, s in enumerate(self.skills):
                    if s.running and s.enabled:
                        s.remaining = max(0, s.remaining - self.interval)
                        if s.remaining <= 0:
                            s.remaining = 0; completed.append((i, s))
                for i, s in completed:
                    self.on_complete(i, s); s.remaining = float(s.cooldown)
                    if self.state == "running": s.running = True
                self.on_tick(self.skills)
            time.sleep(max(0, self.interval - (time.time()-t0)))
        self._running = False
    def stop(self): self._stop_loop()
    def check_early_warning(self, warn_sec):
        for i, s in enumerate(self.skills):
            if s.running and s.enabled and s.remaining <= warn_sec:
                last = self._warn_tracker.get(i, 999)
                if last > warn_sec: self._warn_tracker[i] = s.remaining; return i, s
                self._warn_tracker[i] = s.remaining
        return None, None

# ===================== 窗口叠加层 =====================
def make_overlay(win, topmost=True, click_through=False, opacity=0.85):
    if not IS_WIN: return
    try:
        hwnd = win.winfo_id()
        GWL_EXSTYLE = -20
        WS_EX_LAYERED = 0x80000; WS_EX_TRANSPARENT = 0x20; WS_EX_TOPMOST = 0x8
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        if topmost: style |= WS_EX_TOPMOST
        if click_through: style |= WS_EX_TRANSPARENT
        style |= WS_EX_LAYERED
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
        ctypes.windll.user32.SetLayeredWindowAttributes(hwnd, 0, int(opacity*255), 2)
    except: pass

def set_click_through(win, enable):
    if not IS_WIN: return
    try:
        hwnd = win.winfo_id(); GWL_EXSTYLE = -20; WS_EX_TRANSPARENT = 0x20
        style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        style = (style | WS_EX_TRANSPARENT) if enable else (style & ~WS_EX_TRANSPARENT)
        ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, style)
    except: pass

def set_opacity(win, opacity):
    if not IS_WIN: return
    try:
        hwnd = win.winfo_id()
        ctypes.windll.user32.SetLayeredWindowAttributes(hwnd, 0, int(opacity*255), 2)
    except: pass

# ===================== 技能行组件 =====================
class SkillRow(tk.Frame):
    def __init__(self, parent, index, skill, cb):
        super().__init__(parent, bg=C["row"], height=50)
        self.idx = index; self.skill = skill; self.cb = cb
        self.pack_propagate(False); self.pack(fill=tk.X, padx=2, pady=1)
        self._build()
    def _build(self):
        s = self.skill
        # 启用指示器
        self.dot = tk.Label(self, text="●", font=FONT, cursor="hand2", bg=C["row"])
        self.dot.pack(side=tk.LEFT, padx=(6,4), pady=10)
        self.dot.bind("<Button-1>", lambda e: self.cb("toggle_enabled", self.idx))
        # 名称
        self.lbl_name = tk.Label(self, text=s.name, font=FONT_B, bg=C["row"], fg=C["text"], width=10, anchor=tk.W)
        self.lbl_name.pack(side=tk.LEFT, padx=4, pady=10)
        self.lbl_name.bind("<Double-Button-1>", lambda e: self.cb("edit_name", self.idx))
        # 方向
        self.lbl_dir = tk.Label(self, text=DIR_ARROW.get(s.direction,"—"), font=("",14), cursor="hand2", width=3)
        self.lbl_dir.pack(side=tk.LEFT, padx=4, pady=8)
        self.lbl_dir.bind("<Button-1>", lambda e: self.cb("cycle_dir", self.idx))
        # 冷却
        self.lbl_cd = tk.Label(self, text=f"{s.cooldown}s", font=FONT, bg=C["row"], fg=C["muted"], width=5)
        self.lbl_cd.pack(side=tk.LEFT, padx=4, pady=10)
        self.lbl_cd.bind("<Double-Button-1>", lambda e: self.cb("edit_cd", self.idx))
        # 进度条
        self.canvas = tk.Canvas(self, bg=C["bg"], height=20, highlightthickness=1, highlightbackground=C["border"])
        self.canvas.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=10)
        # 时间
        self.lbl_time = tk.Label(self, text="30.0s", font=FONT_B, bg=C["row"], fg=C["green"], width=7)
        self.lbl_time.pack(side=tk.LEFT, padx=4, pady=10)
        # 操作按钮
        self.btn_pause = tk.Label(self, text="⏯", font=FONT, cursor="hand2", bg=C["row"], fg=C["muted"], width=3)
        self.btn_pause.pack(side=tk.LEFT, padx=1, pady=10)
        self.btn_pause.bind("<Button-1>", lambda e: self.cb("toggle_pause", self.idx))
        self.btn_reset = tk.Label(self, text="🔄", font=FONT, cursor="hand2", bg=C["row"], fg=C["muted"], width=3)
        self.btn_reset.pack(side=tk.LEFT, padx=1, pady=10)
        self.btn_reset.bind("<Button-1>", lambda e: self.cb("reset", self.idx))
    def update(self, skill):
        self.skill = skill; s = skill
        # 启用
        self.dot.config(text="●" if s.enabled else "○", fg=C["green"] if s.enabled else C["dim"])
        # 名称
        nc = C["dim"] if not s.enabled else C["text"]
        self.lbl_name.config(text=s.name, fg=nc)
        # 方向
        dc = C["dim"] if not s.enabled else DIR_CLR.get(s.direction, C["dim"])
        bgc = "#1a1a1a" if not s.enabled else DIR_BG.get(s.direction, "#1a1a1a")
        self.lbl_dir.config(text=DIR_ARROW.get(s.direction,"—"), fg=dc, bg=bgc)
        # 冷却
        self.lbl_cd.config(text=f"{s.cooldown}s", fg=C["dim"] if not s.enabled else C["muted"])
        # 进度条和时间
        if not s.enabled:
            self.lbl_time.config(text="--", fg=C["dim"])
            self._draw_bar(0, C["dim"])
        else:
            r = max(0, s.remaining)
            ratio = r / s.cooldown if s.cooldown > 0 else 0
            if r <= 5: tc = C["red"]
            elif r <= 10: tc = C["yellow"]
            else: tc = C["green"]
            self.lbl_time.config(text=f"{r:.1f}s", fg=tc)
            self._draw_bar(ratio, tc)
        # 按钮
        bc = C["dim"] if not s.enabled else C["muted"]
        self.btn_pause.config(fg=bc); self.btn_reset.config(fg=bc)
        # 行背景
        warn = s.running and s.enabled and s.remaining <= 5
        row_bg = C["row_warn"] if warn else C["row"]
        self.config(bg=row_bg)
        for w in [self.dot, self.lbl_name, self.lbl_time, self.btn_pause, self.btn_reset]:
            try: w.config(bg=row_bg)
            except: pass
    def _draw_bar(self, ratio, color):
        self.canvas.delete("all")
        w = self.canvas.winfo_width(); h = self.canvas.winfo_height()
        if w < 5: return
        fw = int(w * ratio)
        if fw > 0: self.canvas.create_rectangle(0, 0, fw, h, fill=color, outline="")
        if fw < w: self.canvas.create_rectangle(fw, 0, w, h, fill=C["bg"], outline="")

# ===================== 设置弹窗 =====================
class SettingsDialog(tk.Toplevel):
    def __init__(self, parent, voice, callbacks):
        super().__init__(parent); self.voice = voice; self.cb = callbacks
        self.title("设置"); self.geometry("420x350"); self.resizable(False,False)
        self.configure(bg=C["bg"]); self.transient(parent)
        self._build()
        self.update_idletasks()
        self.geometry(f"+{parent.winfo_x()+parent.winfo_width()//2-210}+{parent.winfo_y()+30}")
    def _build(self):
        pad = {"padx":16,"pady":6}
        tk.Label(self, text="⚙ 语音与显示设置", font=FONT_B, fg=C["text"], bg=C["bg"]).pack(anchor=tk.W, **pad)
        # 语音
        vf = tk.LabelFrame(self, text="🔊 语音播报", font=FONT_S, fg=C["muted"], bg=C["bg"], padx=12, pady=8)
        vf.pack(fill=tk.X, padx=14, pady=4)
        self.voice_var = tk.BooleanVar(value=self.voice.enabled)
        tk.Checkbutton(vf, text="启用语音播报", variable=self.voice_var, font=FONT_S, fg=C["text"], bg=C["bg"],
                       selectcolor=C["bg"], activebackground=C["bg"], activeforeground=C["text"],
                       command=lambda: setattr(self.voice, 'enabled', self.voice_var.get())).pack(anchor=tk.W)
        # 语速
        rf = tk.Frame(vf, bg=C["bg"]); rf.pack(fill=tk.X, pady=2)
        tk.Label(rf, text="语速:", font=FONT_S, fg=C["muted"], bg=C["bg"], width=5).pack(side=tk.LEFT)
        self.rate_var = tk.IntVar(value=self.voice.rate)
        tk.Scale(rf, from_=-10, to=10, orient=tk.HORIZONTAL, variable=self.rate_var, bg=C["bg"], fg=C["text"],
                 highlightthickness=0, troughcolor=C["border"], command=lambda v: setattr(self.voice, 'rate', int(v))).pack(side=tk.LEFT, fill=tk.X, expand=True)
        # 预警
        wf = tk.Frame(vf, bg=C["bg"]); wf.pack(fill=tk.X, pady=2)
        tk.Label(wf, text="预警:", font=FONT_S, fg=C["muted"], bg=C["bg"], width=5).pack(side=tk.LEFT)
        self.warn_var = tk.IntVar(value=self.voice.early_warning)
        for v, t in [(0,"关"),(3,"3s"),(5,"5s"),(10,"10s")]:
            tk.Radiobutton(wf, text=t, variable=self.warn_var, value=v, font=FONT_S, fg=C["muted"], bg=C["bg"],
                           selectcolor=C["bg"], activebackground=C["bg"], activeforeground=C["text"],
                           command=lambda v=v: setattr(self.voice, 'early_warning', v)).pack(side=tk.LEFT, padx=3)
        # 测试
        tk.Button(vf, text="📢 测试语音", font=FONT_S, bg=C["border"], fg=C["text"], relief=tk.FLAT,
                  cursor="hand2", command=self.voice.test_speak, padx=12, pady=3).pack(anchor=tk.W, pady=4)
        # 显示
        df = tk.LabelFrame(self, text="🖥 显示设置", font=FONT_S, fg=C["muted"], bg=C["bg"], padx=12, pady=8)
        df.pack(fill=tk.X, padx=14, pady=4)
        self.top_var = tk.BooleanVar(value=True)
        tk.Checkbutton(df, text="窗口置顶", variable=self.top_var, font=FONT_S, fg=C["text"], bg=C["bg"],
                       selectcolor=C["bg"], activebackground=C["bg"],
                       command=lambda: self.cb.get("topmost", lambda v:None)(self.top_var.get())).pack(anchor=tk.W)
        self.ct_var = tk.BooleanVar(value=False)
        tk.Checkbutton(df, text="🖱 点击穿透", variable=self.ct_var, font=FONT_S, fg=C["text"], bg=C["bg"],
                       selectcolor=C["bg"], activebackground=C["bg"],
                       command=lambda: self.cb.get("click", lambda v:None)(self.ct_var.get())).pack(anchor=tk.W)
        of = tk.Frame(df, bg=C["bg"]); of.pack(fill=tk.X, pady=2)
        tk.Label(of, text="透明度:", font=FONT_S, fg=C["muted"], bg=C["bg"], width=5).pack(side=tk.LEFT)
        self.op_var = tk.IntVar(value=85)
        self.lbl_op = tk.Label(of, text="85%", font=FONT_S, fg=C["muted"], bg=C["bg"], width=4)
        self.lbl_op.pack(side=tk.RIGHT)
        tk.Scale(of, from_=30, to=100, orient=tk.HORIZONTAL, variable=self.op_var, bg=C["bg"], fg=C["text"],
                 highlightthickness=0, troughcolor=C["border"],
                 command=lambda v: [self.lbl_op.config(text=f"{int(v)}%"),
                                    self.cb.get("opacity",lambda v:None)(int(v)/100)]).pack(side=tk.LEFT, fill=tk.X, expand=True)

# ===================== 主窗口 =====================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME); self.geometry("780x520"); self.minsize(600,320)
        self.configure(bg=C["bg"])
        self.timer = TimerEngine(0.1); self.voice = VoiceEngine(); self.presets = PresetManager()
        self.timer.on_tick = lambda s: self.after(0, self._refresh)
        self.timer.on_complete = lambda i,s: self.after(0, lambda: self._on_complete(i,s))
        self.timer.on_state = lambda s: self.after(0, lambda: self._on_state(s))
        self.rows = []; self.cur_preset = None; self._warn_tracker = {}
        self._click_thru = False; self._topmost = True
        self._build_ui()
        self._load_preset("demo_boss.json")
        self.after(300, lambda: make_overlay(self, True, False, 0.85))
        self.protocol("WM_DELETE_WINDOW", self._on_close)
    # ---- UI 构建 ----
    def _build_ui(self):
        # 标题栏
        tb = tk.Frame(self, bg="#1a1030", height=36); tb.pack(fill=tk.X); tb.pack_propagate(False)
        self.lbl_boss = tk.Label(tb, text=f"🔥 {APP_NAME}", font=FONT_B, fg=C["yellow"], bg="#1a1030")
        self.lbl_boss.pack(side=tk.LEFT, padx=12, pady=4)
        bf = tk.Frame(tb, bg="#1a1030"); bf.pack(side=tk.RIGHT, padx=8)
        self.btn_pin = tk.Label(bf, text="📌", font=("",13), cursor="hand2", bg="#1a1030", fg=C["green"], width=3)
        self.btn_pin.pack(side=tk.LEFT); self.btn_pin.bind("<Button-1>", lambda e: self._toggle_topmost())
        self.btn_clk = tk.Label(bf, text="🖱", font=("",13), cursor="hand2", bg="#1a1030", fg=C["muted"], width=3)
        self.btn_clk.pack(side=tk.LEFT); self.btn_clk.bind("<Button-1>", lambda e: self._toggle_click())
        self.btn_cfg = tk.Label(bf, text="⚙", font=("",13), cursor="hand2", bg="#1a1030", fg=C["muted"], width=3)
        self.btn_cfg.pack(side=tk.LEFT); self.btn_cfg.bind("<Button-1>", lambda e: self._open_settings())
        # 表头
        hf = tk.Frame(self, bg=C["header"], height=26); hf.pack(fill=tk.X, padx=4, pady=(4,0)); hf.pack_propagate(False)
        for t, w in [("启用",5),("技能名称",10),("方向",4),("冷却",4),("倒计时进度",30),("剩余",6),("操作",6)]:
            tk.Label(hf, text=t, font=FONT_S, fg=C["muted"], bg=C["header"], width=w).pack(side=tk.LEFT, padx=2, pady=3)
        # 技能列表
        self.canvas = tk.Canvas(self, bg=C["bg"], highlightthickness=0)
        self.scrollbar = tk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        self.skill_frame = tk.Frame(self.canvas, bg=C["bg"])
        self.skill_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0,0), window=self.skill_frame, anchor=tk.NW)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-e.delta/120),"units"))
        # 底部栏
        bb = tk.Frame(self, bg=C["panel"], height=68); bb.pack(fill=tk.X, side=tk.BOTTOM, padx=4, pady=4); bb.pack_propagate(False)
        lf = tk.Frame(bb, bg=C["panel"]); lf.pack(side=tk.LEFT, padx=8, pady=10)
        tk.Label(lf, text="+ 添加技能", font=FONT_S, bg=C["blue"], fg=C["text"], cursor="hand2", padx=12, pady=4).pack(side=tk.LEFT, padx=4)
        # 方案下拉
        self.preset_var = tk.StringVar(value="选择方案...")
        presets_list = self.presets.list()
        self.preset_menu = tk.OptionMenu(lf, self.preset_var, presets_list[0] if presets_list else "无方案", *presets_list,
            command=lambda f: self._load_preset(f))
        self.preset_menu.config(font=FONT_S, bg=C["bg"], fg=C["muted"], highlightthickness=0, width=18)
        self.preset_menu.pack(side=tk.LEFT, padx=6)
        tk.Label(lf, text="💾 保存", font=FONT_S, bg=C["border"], fg=C["text"], cursor="hand2", padx=10, pady=4).pack(side=tk.LEFT, padx=4)
        # 右侧控制
        rf = tk.Frame(bb, bg=C["panel"]); rf.pack(side=tk.RIGHT, padx=8, pady=10)
        self.btn_start = tk.Label(rf, text="▶ 全部开始", font=FONT_B, bg=C["green"], fg="#fff", cursor="hand2", padx=14, pady=5)
        self.btn_start.pack(side=tk.LEFT, padx=3)
        self.btn_start.bind("<Button-1>", lambda e: self._start_pause())
        tk.Label(rf, text="⏸ 暂停", font=FONT_B, bg="#5c3d1a", fg=C["yellow"], cursor="hand2", padx=12, pady=5).pack(side=tk.LEFT, padx=3)
        tk.Label(rf, text="🔄 重置", font=FONT_B, bg="#5c1a1a", fg=C["red"], cursor="hand2", padx=12, pady=5).pack(side=tk.LEFT, padx=3)
        self.lbl_voice = tk.Label(bb, text="🔊", font=("",15), bg=C["panel"], fg=C["green"], cursor="hand2")
        self.lbl_voice.pack(side=tk.RIGHT, padx=6, pady=16)
        self.lbl_voice.bind("<Button-1>", lambda e: self._toggle_voice())
        self.lbl_state = tk.Label(bb, text="🟢 准备就绪", font=FONT_S, bg=C["panel"], fg=C["muted"])
        self.lbl_state.pack(side=tk.RIGHT, padx=12, pady=16)
    # ---- 行管理 ----
    def _rebuild_rows(self):
        for r in self.rows: r.destroy()
        self.rows.clear()
        cbs = {"toggle_enabled": self._on_te,"toggle_pause": self._on_tp,"reset": self._on_rs,
               "cycle_dir": self._on_cd,"edit_name": self._on_en,"edit_cd": self._on_ec}
        for i, s in enumerate(self.timer.skills):
            r = SkillRow(self.skill_frame, i, s, cbs); self.rows.append(r)
        self.skill_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    def _refresh(self):
        skills = self.timer.skills
        for i, row in enumerate(self.rows):
            if i < len(skills): row.update(skills[i])
        self._check_warn()
    def _check_warn(self):
        ws = self.voice.early_warning
        if ws > 0:
            idx, s = self.timer.check_early_warning(ws)
            if s: self.voice.speak_warning(s.name, ws)
    # ---- 事件 ----
    def _on_complete(self, idx, skill):
        self.voice.speak_skill(skill.name, skill.direction)
    def _on_state(self, state):
        if state == "running":
            self.lbl_state.config(text="🔴 运行中", fg=C["red"])
            self.btn_start.config(text="⏸ 全部暂停", bg="#5c3d1a", fg=C["yellow"])
        elif state == "paused":
            self.lbl_state.config(text="🟡 已暂停", fg=C["yellow"])
            self.btn_start.config(text="▶ 继续", bg=C["green"], fg="#fff")
        else:
            self.lbl_state.config(text="🟢 准备就绪", fg=C["green"])
            self.btn_start.config(text="▶ 全部开始", bg=C["green"], fg="#fff")
    def _start_pause(self):
        s = self.timer.state
        if s == "running": self.timer.pause()
        elif s == "paused": self.timer.resume()
        else: self.timer.start()
    def _on_te(self, i): self.timer.toggle_enabled(i); self._refresh()
    def _on_tp(self, i): self.timer.toggle_skill(i); self._refresh()
    def _on_rs(self, i): self.timer.reset_skill(i); self._refresh()
    def _on_cd(self, i):
        s = self.timer.skills[i]; dirs = ["up","down","left","right","none"]
        cur = dirs.index(s.direction) if s.direction in dirs else 4
        s.direction = dirs[(cur+1)%5]; self._refresh()
    def _on_en(self, i):
        s = self.timer.skills[i]
        n = simpledialog.askstring("编辑技能名", "技能名称:", initialvalue=s.name, parent=self)
        if n and n.strip(): s.name = n.strip(); self._refresh()
    def _on_ec(self, i):
        s = self.timer.skills[i]
        v = simpledialog.askinteger("编辑冷却", "冷却时间 (1-60秒):", initialvalue=s.cooldown, minvalue=1, maxvalue=60, parent=self)
        if v: s.cooldown = v; s.remaining = float(v); self._refresh()
    def _toggle_topmost(self):
        self._topmost = not self._topmost
        self.btn_pin.config(fg=C["green"] if self._topmost else C["muted"])
        try: self.attributes('-topmost', self._topmost)
        except: pass
    def _toggle_click(self):
        self._click_thru = not self._click_thru
        self.btn_clk.config(fg=C["green"] if self._click_thru else C["muted"])
        set_click_through(self, self._click_thru)
    def _toggle_voice(self):
        self.voice.enabled = not self.voice.enabled
        self.lbl_voice.config(fg=C["green"] if self.voice.enabled else C["dim"],
                              text="🔊" if self.voice.enabled else "🔇")
    def _open_settings(self):
        cbs = {"topmost": lambda v: [setattr(self,'_topmost',v), self.btn_pin.config(fg=C["green"] if v else C["muted"]), self.attributes('-topmost', v)],
               "click": lambda v: [setattr(self,'_click_thru',v), self.btn_clk.config(fg=C["green"] if v else C["muted"]), set_click_through(self, v)],
               "opacity": lambda v: set_opacity(self, v)}
        SettingsDialog(self, self.voice, cbs)
    def _load_preset(self, fn):
        if not fn or fn == "选择方案...": return
        p = self.presets.load(fn)
        if p:
            self.timer.reset(); self.timer.set_skills(p.skills)
            self.cur_preset = fn; self.lbl_boss.config(text=f"🔥 {p.name}")
            self._rebuild_rows(); self._refresh()
            self._update_menu()
    def _save_preset(self):
        skills = self.timer.skills
        n = simpledialog.askstring("保存方案", "BOSS 名称:", parent=self)
        if not n or not n.strip(): return
        p = BossPreset(n.strip(), skills=[Skill(s.name,s.cooldown,s.direction,s.enabled) for s in skills])
        self.presets.save(p, f"{n.strip()}.json")
        self.cur_preset = f"{n.strip()}.json"; self.lbl_boss.config(text=f"🔥 {n.strip()}")
        self._update_menu(); messagebox.showinfo("成功", f"方案已保存: {n.strip()}.json")
    def _update_menu(self):
        menu = self.preset_menu["menu"]; menu.delete(0,"end")
        for f in self.presets.list():
            menu.add_command(label=f, command=lambda v=f: [self.preset_var.set(v), self._load_preset(v)])
    def _on_close(self):
        self.timer.stop(); self.destroy()

if __name__ == "__main__":
    App().mainloop()
