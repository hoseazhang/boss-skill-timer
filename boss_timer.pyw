#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BOSS Skill Cooldown Timer v3.3
Minimal, robust, tested step by step.
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json, os, time, threading, platform, sys

# ── Config ──────────────────────────────────────────────────────
APP_NAME = "BOSS技能倒计时器"

FONT = ("Microsoft YaHei", 11) if platform.system() == "Windows" else ("", 13)
FONT_B = (FONT[0], FONT[1], "bold")
FONT_S = (FONT[0], FONT[1] - 1)

C = {"bg":"#0f0f1a","fg":"#e5e5e5","dim":"#555","green":"#22c55e","red":"#ef4444","yellow":"#f59e0b","blue":"#3b82f6","row":"#12122a","panel":"#161630","header":"#1a1a35","border":"#2a2a4a","muted":"#a0a0b0","warn":"#1a0f0f"}

IS_WIN = platform.system() == "Windows"
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "presets")


# ── Skill ───────────────────────────────────────────────────────
class Skill:
    def __init__(self, name="技能", cd=30, on=True):
        self.name = name; self.cd = max(1, min(60, int(cd)))
        self.on = bool(on); self.left = float(self.cd); self.run = False
    def reset(self): self.left = float(self.cd); self.run = False
    def to_d(self): return {"name": self.name, "cooldown": self.cd, "enabled": self.on}
    @classmethod
    def from_d(cls, d): return cls(d.get("name","技能"), d.get("cooldown",30), d.get("enabled",True))


# ── Voice ────────────────────────────────────────────────────────
class Voice:
    def __init__(self):
        self.on = True; self.rate = 2; self.vol = 100; self.warn = 0; self._v = None
        if IS_WIN:
            try:
                import win32com.client; self._v = win32com.client.Dispatch("SAPI.SpVoice")
            except: pass
    def say(self, text):
        if not self.on or not text: return
        def _go():
            try:
                if IS_WIN and self._v:
                    self._v.Rate = self.rate; self._v.Volume = self.vol; self._v.Speak(text, 1)
                elif platform.system() == "Darwin":
                    os.system(f'say -v Tingting "{text}" &')
            except: pass
        threading.Thread(target=_go, daemon=True).start()
    def test(self): self.say(f"{APP_NAME} 语音测试")


# ── Timer ────────────────────────────────────────────────────────
class Timer:
    def __init__(self, iv=0.1):
        self._s = []; self._iv = iv; self._lk = threading.Lock()
        self._go = False; self._th = None; self._w = {}
        self.state = "idle"  # idle | running | paused
        self.on_tick = lambda s: None
        self.on_done = lambda i, s: None
        self.on_state = lambda s: None

    def get(self):
        with self._lk: return list(self._s)

    def set(self, skills):
        with self._lk: self._s = [Skill(s.name, s.cd, s.on) for s in skills]

    def add(self, name="新技能", cd=30):
        with self._lk: s = Skill(name, cd); self._s.append(s); return len(self._s)-1

    def rm(self, i):
        with self._lk:
            if 0 <= i < len(self._s): self._s.pop(i)

    def start_all(self):
        with self._lk:
            if self.state == "running": return
            self.state = "running"
            for s in self._s:
                if s.on: s.run = True
                if s.left <= 0: s.left = float(s.cd)
            self._w.clear(); self._start(); self.on_state("running")

    def pause_all(self):
        with self._lk:
            if self.state != "running": return
            self.state = "paused"; self._stop()
            for s in self._s: s.run = False
            self.on_state("paused")

    def resume_all(self):
        with self._lk:
            if self.state != "paused": return
            self.state = "running"
            for s in self._s:
                if s.on: s.run = True
            self._w.clear(); self._start(); self.on_state("running")

    def reset_all(self):
        with self._lk:
            self._stop(); self.state = "idle"; self._w.clear()
            for s in self._s: s.reset()
            self.on_state("idle"); self.on_tick(self._s)

    def start_one(self, i):
        with self._lk:
            if 0 <= i < len(self._s):
                s = self._s[i]
                if not s.on: return
                s.run = True; s.left = float(s.cd)
                self._w.pop(i, None); self._start()

    def reset_one(self, i):
        with self._lk:
            if 0 <= i < len(self._s):
                s = self._s[i]; s.left = float(s.cd)
                s.run = (self.state == "running" and s.on)
                self._w.pop(i, None)

    def toggle_on(self, i):
        with self._lk:
            if 0 <= i < len(self._s):
                s = self._s[i]; s.on = not s.on
                if not s.on: s.run = False; s.left = float(s.cd)
                elif self.state == "running": s.run = True; s.left = float(s.cd)

    def upd(self, i, name=None, cd=None):
        with self._lk:
            if 0 <= i < len(self._s):
                s = self._s[i]
                if name: s.name = name
                if cd: s.cd = max(1, min(60, int(cd))); s.left = float(s.cd)

    def warn_check(self, ws):
        if ws <= 0: return None
        with self._lk:
            for i, s in enumerate(self._s):
                if s.run and s.on and 0 < s.left <= ws:
                    l = self._w.get(i, 999)
                    if l > ws: self._w[i] = s.left; return s
                    self._w[i] = s.left
        return None

    def _start(self):
        if self._go: return
        self._go = True; self._th = threading.Thread(target=self._loop, daemon=True); self._th.start()

    def _stop(self): self._go = False

    def _loop(self):
        while self._go:
            t0 = time.time()
            with self._lk:
                done = []
                for i, s in enumerate(self._s):
                    if s.run and s.on:
                        s.left = max(0.0, s.left - self._iv)
                        if s.left <= 0: s.left = 0.0; s.run = False; done.append((i, s))
                for i, s in done:
                    self.on_done(i, s); s.left = float(s.cd)
                self.on_tick(self._s)
                if self.state != "running" and not any(s.run for s in self._s):
                    self._go = False; break
            time.sleep(max(0.0, self._iv - (time.time() - t0)))
        self._go = False


# ── Row ──────────────────────────────────────────────────────────
class Row(tk.Frame):
    def __init__(self, p, i, s, cb):
        super().__init__(p, bg=C["row"], height=48)
        self.idx = i; self.skill = s; self.cb = cb
        self.pack_propagate(False); self.pack(fill=tk.X, padx=2, pady=1)
        # enable dot
        self.dot = tk.Label(self, text="●", font=FONT, cursor="hand2", bg=C["row"], fg=C["green"])
        self.dot.pack(side=tk.LEFT, padx=(6,4), pady=8)
        self.dot.bind("<Button-1>", lambda e: cb["on"](self.idx))
        # name
        self.ln = tk.Label(self, text=s.name, font=FONT_B, bg=C["row"], fg=C["fg"], width=12, anchor=tk.W)
        self.ln.pack(side=tk.LEFT, padx=4, pady=8)
        # cd label
        self.lc = tk.Label(self, text=f"{s.cd}s", font=FONT, bg=C["row"], fg=C["muted"], width=5)
        self.lc.pack(side=tk.LEFT, padx=4, pady=8)
        # progress bar
        self.cv = tk.Canvas(self, bg=C["bg"], height=18, highlightthickness=0)
        self.cv.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6, pady=8)
        # time
        self.lt = tk.Label(self, text=f"{s.cd:.1f}s", font=FONT_B, bg=C["row"], fg=C["green"], width=7)
        self.lt.pack(side=tk.LEFT, padx=4, pady=8)
        # start
        self.bs = tk.Label(self, text="▶ 开始", font=FONT_S, cursor="hand2", bg=C["green"], fg="#fff", padx=6, pady=1)
        self.bs.pack(side=tk.LEFT, padx=2, pady=8)
        self.bs.bind("<Button-1>", lambda e: cb["start"](self.idx))
        # reset
        self.br = tk.Label(self, text="🔄", font=FONT, cursor="hand2", bg=C["row"], fg=C["red"], width=3)
        self.br.pack(side=tk.LEFT, padx=2, pady=8)
        self.br.bind("<Button-1>", lambda e: cb["reset"](self.idx))

    def rf(self, s):
        self.skill = s
        d = C["dim"]; m = C["muted"]; on = s.on
        self.dot.config(text="●" if on else "○", fg=C["green"] if on else d)
        self.ln.config(text=s.name, fg=d if not on else C["fg"])
        self.lc.config(text=f"{s.cd}s", fg=d if not on else m)
        if not on:
            self.lt.config(text="--", fg=d); self._bar(0, d)
        elif not s.run:
            self.lt.config(text=f"{s.cd:.1f}s", fg=C["green"]); self._bar(1, C["green"])
        else:
            r = max(0.0, s.left); rt = r / s.cd if s.cd else 0
            cl = C["red"] if r <= 5 else (C["yellow"] if r <= 10 else C["green"])
            self.lt.config(text=f"{r:.1f}s", fg=cl); self._bar(rt, cl)
        if not on: self.bs.config(bg=d, fg=C["bg"], text="▶")
        elif s.run: self.bs.config(bg="#5c3d1a", fg=C["yellow"], text="⏸")
        else: self.bs.config(bg=C["green"], fg="#fff", text="▶")
        nb = C["warn"] if (s.run and on and s.left <= 5) else C["row"]
        self.config(bg=nb)
        for w in (self.dot, self.ln, self.lt): w.config(bg=nb)

    def _bar(self, rt, cl):
        self.cv.delete("all")
        w = self.cv.winfo_width(); h = self.cv.winfo_height()
        if w < 5: return
        fw = int(w * max(0.0, min(1.0, rt)))
        if fw > 0: self.cv.create_rectangle(0, 0, fw, h, fill=cl, outline="")
        if fw < w: self.cv.create_rectangle(fw, 0, w, h, fill=C["bg"], outline="")


# ── Config Dialog ────────────────────────────────────────────────
class ConfigDialog(tk.Toplevel):
    def __init__(self, parent, timer, presets):
        super().__init__(parent)
        self.timer = timer; self.presets = presets
        self.title("配置方案"); self.geometry("500x450"); self.resizable(False, False)
        self.configure(bg=C["bg"]); self.transient(parent)
        self._skills = []; self._build()
        self.update_idletasks()
        x = parent.winfo_x() + parent.winfo_width()//2 - 250
        y = parent.winfo_y() + 20
        self.geometry(f"+{x}+{y}"); self.grab_set()

    def _build(self):
        # top: preset selection
        pf = tk.Frame(self, bg=C["panel"]); pf.pack(fill=tk.X, padx=10, pady=(10,0))
        tk.Label(pf, text="方案:", font=FONT_S, fg=C["muted"], bg=C["panel"]).pack(side=tk.LEFT, padx=4)
        flist = self._list_files()
        self.fvar = tk.StringVar(value=flist[0] if flist else "")
        self.fm = tk.OptionMenu(pf, self.fvar, flist[0] if flist else "", *flist, command=self._load)
        self.fm.config(font=FONT_S, bg=C["bg"], fg=C["muted"], highlightthickness=0, width=22)
        self.fm.pack(side=tk.LEFT, padx=4)
        self.name_var = tk.StringVar()
        tk.Entry(pf, textvariable=self.name_var, font=FONT_S, bg=C["border"], fg=C["fg"],
                 insertbackground=C["fg"], width=14).pack(side=tk.LEFT, padx=4)
        # skill list
        sf = tk.Frame(self, bg=C["bg"]); sf.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        cv = tk.Canvas(sf, bg=C["bg"], highlightthickness=0)
        sb = tk.Scrollbar(sf, command=cv.yview)
        self.inner = tk.Frame(cv, bg=C["bg"])
        self.inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0, 0), window=self.inner, anchor=tk.NW)
        cv.configure(yscrollcommand=sb.set); cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._row_frames = []
        # bottom buttons
        bf = tk.Frame(self, bg=C["panel"]); bf.pack(fill=tk.X, padx=10, pady=8)
        self._btn(bf, "+ 添加", C["blue"], self._add).pack(side=tk.LEFT, padx=4)
        self._btn(bf, "💾 保存并应用", C["green"], self._save).pack(side=tk.LEFT, padx=4)
        self._btn(bf, "关闭", C["border"], self.destroy).pack(side=tk.RIGHT, padx=4)
        if flist: self._load(flist[0])

    def _btn(self, p, t, c, cmd):
        b = tk.Label(p, text=t, font=FONT_S, cursor="hand2", bg=c, fg="#fff", padx=10, pady=3)
        b.bind("<Button-1>", lambda e: cmd()); return b

    def _list_files(self):
        if not os.path.exists(DATA_DIR): return []
        return sorted(f for f in os.listdir(DATA_DIR) if f.endswith(".json"))

    def _load(self, fn):
        if not fn: return
        path = os.path.join(DATA_DIR, fn)
        if not os.path.exists(path): return
        with open(path, "r", encoding="utf-8") as f: d = json.load(f)
        self._skills = [Skill.from_d(s) for s in d.get("skills", [])]
        self.name_var.set(d.get("name", "")); self._show()

    def _show(self):
        for f in self._row_frames: f.destroy()
        self._row_frames.clear()
        for i, s in enumerate(self._skills):
            fr = tk.Frame(self.inner, bg=C["row"]); fr.pack(fill=tk.X, padx=2, pady=1)
            self._row_frames.append(fr)
            sv = tk.StringVar(value=s.name)
            tk.Entry(fr, textvariable=sv, font=FONT_S, bg=C["border"], fg=C["fg"],
                     insertbackground=C["fg"], width=16).pack(side=tk.LEFT, padx=2, pady=3)
            cv_ = tk.IntVar(value=s.cd)
            tk.Spinbox(fr, textvariable=cv_, from_=1, to=60, font=FONT_S,
                       bg=C["border"], fg=C["fg"], width=5, buttonbackground=C["border"]).pack(side=tk.LEFT, padx=2, pady=3)
            ev = tk.BooleanVar(value=s.on)
            tk.Checkbutton(fr, variable=ev, bg=C["row"], fg=C["fg"],
                           selectcolor=C["row"], activebackground=C["row"]).pack(side=tk.LEFT, padx=4)
            tk.Label(fr, text="✕", font=FONT, cursor="hand2", bg=C["row"], fg=C["red"], width=2).pack(side=tk.LEFT, padx=4)
            fr.winfo_children()[-1].bind("<Button-1>", lambda e, idx=i: self._rm(idx))
            fr._nv = sv; fr._cv = cv_; fr._ev = ev

    def _add(self):
        self._skills.append(Skill("新技能", 30)); self._show()

    def _rm(self, idx):
        if 0 <= idx < len(self._skills): self._skills.pop(idx); self._show()

    def _save(self):
        name = self.name_var.get().strip()
        if not name: messagebox.showwarning("提示", "请输入方案名称", parent=self); return
        for i, fr in enumerate(self._row_frames):
            if i < len(self._skills):
                s = self._skills[i]
                s.name = fr._nv.get().strip() or s.name
                try: s.cd = max(1, min(60, fr._cv.get()))
                except: pass
                s.on = fr._ev.get()
        fn = f"{name}.json"
        data = {"name": name, "skills": [s.to_d() for s in self._skills]}
        path = os.path.join(DATA_DIR, fn)
        with open(path, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
        self.timer.reset_all(); self.timer.set(self._skills)
        self.destroy()


# ── Main App ─────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME); self.geometry("760x500"); self.minsize(500, 280)
        self.configure(bg=C["bg"])

        self.timer = Timer(0.1); self.voice = Voice()
        self.timer.on_tick = lambda s: self.after(0, self._refresh)
        self.timer.on_done = lambda i, s: self.after(0, lambda: self.voice.say(s.name))
        self.timer.on_state = lambda s: self.after(0, lambda: self._set_state(s))

        self.rows = []; self._top = True; self._clk = False
        self._build()
        self._init_data()
        self.protocol("WM_DELETE_WINDOW", self._close)

    def _build(self):
        # title
        tb = tk.Frame(self, bg=C["header"], height=34)
        tb.pack(fill=tk.X); tb.pack_propagate(False)
        self.lbl_title = tk.Label(tb, text=f"🔥 {APP_NAME}", font=FONT_B, fg=C["yellow"], bg=C["header"])
        self.lbl_title.pack(side=tk.LEFT, padx=12, pady=4)
        tr = tk.Frame(tb, bg=C["header"]); tr.pack(side=tk.RIGHT, padx=8)
        self.btn_pin = tk.Label(tr, text="📌", font=("", 13), cursor="hand2", bg=C["header"], fg=C["green"], width=3)
        self.btn_pin.pack(side=tk.LEFT); self.btn_pin.bind("<Button-1>", lambda e: self._tog_top())
        self.btn_clk = tk.Label(tr, text="🖱", font=("", 13), cursor="hand2", bg=C["header"], fg=C["muted"], width=3)
        self.btn_clk.pack(side=tk.LEFT); self.btn_clk.bind("<Button-1>", lambda e: self._tog_clk())
        self.btn_set = tk.Label(tr, text="⚙", font=("", 13), cursor="hand2", bg=C["header"], fg=C["muted"], width=3)
        self.btn_set.pack(side=tk.LEFT); self.btn_set.bind("<Button-1>", lambda e: self._open_settings())

        # header row
        hf = tk.Frame(self, bg=C["bg"], height=24)
        hf.pack(fill=tk.X, padx=4, pady=(4, 0)); hf.pack_propagate(False)
        for t, w in [("启用", 5), ("技能名称", 14), ("冷却", 5), ("倒计时进度", 38), ("剩余", 6), ("操作", 12)]:
            tk.Label(hf, text=t, font=FONT_S, fg=C["muted"], bg=C["bg"], width=w).pack(side=tk.LEFT, padx=1, pady=2)

        # skill list area
        self.cv = tk.Canvas(self, bg=C["bg"], highlightthickness=0)
        self.sb = tk.Scrollbar(self, command=self.cv.yview)
        self.fr = tk.Frame(self.cv, bg=C["bg"])
        self.fr.bind("<Configure>", lambda e: self.cv.configure(scrollregion=self.cv.bbox("all")))
        self.cv.create_window((0, 0), window=self.fr, anchor=tk.NW)
        self.cv.configure(yscrollcommand=self.sb.set)
        self.cv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=4)
        self.sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.cv.bind_all("<MouseWheel>", lambda e: self.cv.yview_scroll(int(-e.delta/120), "units"))

        # bottom bar
        bb = tk.Frame(self, bg=C["panel"], height=64)
        bb.pack(fill=tk.X, side=tk.BOTTOM, padx=4, pady=4); bb.pack_propagate(False)

        bl = tk.Frame(bb, bg=C["panel"]); bl.pack(side=tk.LEFT, padx=8, pady=10)
        self._make_btn(bl, "🔧 配置", C["blue"], self._open_config).pack(side=tk.LEFT, padx=4)
        self._make_btn(bl, "💾 保存", C["border"], self._save).pack(side=tk.LEFT, padx=4)

        br = tk.Frame(bb, bg=C["panel"]); br.pack(side=tk.RIGHT, padx=8, pady=10)
        self.btn_start = self._make_btn(br, "▶ 全部开始", C["green"], self._start)
        self.btn_start.pack(side=tk.LEFT, padx=3)
        self._make_btn(br, "⏸ 暂停", "#5c3d1a", self.timer.pause_all).pack(side=tk.LEFT, padx=3)
        self._make_btn(br, "🔄 重置", "#5c1a1a", self.timer.reset_all).pack(side=tk.LEFT, padx=3)

        self.lbl_voice = tk.Label(bb, text="🔊", font=("", 15), bg=C["panel"], fg=C["green"], cursor="hand2")
        self.lbl_voice.pack(side=tk.RIGHT, padx=6, pady=14)
        self.lbl_voice.bind("<Button-1>", lambda e: self._tog_voice())

        self.lbl_state = tk.Label(bb, text="🟢 就绪", font=FONT_S, bg=C["panel"], fg=C["muted"])
        self.lbl_state.pack(side=tk.RIGHT, padx=10, pady=14)

    def _make_btn(self, p, text, color, cmd):
        b = tk.Label(p, text=text, font=FONT_S, cursor="hand2", bg=color, fg="#fff", padx=10, pady=3)
        b.bind("<Button-1>", lambda e: cmd()); return b

    def _init_data(self):
        os.makedirs(DATA_DIR, exist_ok=True)
        dp = os.path.join(DATA_DIR, "demo_boss.json")
        if not os.path.exists(dp):
            data = {"name": "炎狱之王", "skills": [
                {"name": "火焰吐息", "cooldown": 30}, {"name": "暗影冲锋", "cooldown": 20},
                {"name": "陨石坠落", "cooldown": 90}, {"name": "冰霜禁锢", "cooldown": 20},
            ]}
            with open(dp, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
        tp = os.path.join(DATA_DIR, "new_boss_template.json")
        if not os.path.exists(tp):
            data = {"name": "新BOSS", "skills": [{"name": "技能1", "cooldown": 30}, {"name": "技能2", "cooldown": 20}]}
            with open(tp, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
        # Load demo
        with open(dp, "r", encoding="utf-8") as f: d = json.load(f)
        skills = [Skill.from_d(s) for s in d.get("skills", [])]
        self.timer.set(skills); self.lbl_title.config(text=f"🔥 {d['name']}")
        self._rebuild(); self._refresh()

    def _rebuild(self):
        for r in self.rows: r.destroy()
        self.rows.clear()
        cb = {"on": self._on_on, "start": self._on_start, "reset": self._on_reset}
        for i, s in enumerate(self.timer.get()):
            self.rows.append(Row(self.fr, i, s, cb))
        self.fr.update_idletasks()
        self.cv.configure(scrollregion=self.cv.bbox("all"))

    def _refresh(self):
        sk = self.timer.get()
        for i, r in enumerate(self.rows):
            if i < len(sk): r.rf(sk[i])
        ws = self.voice.warn
        if ws > 0:
            s = self.timer.warn_check(ws)
            if s: self.voice.say(f"准备 {s.name}")

    def _set_state(self, st):
        if st == "running":
            self.lbl_state.config(text="🔴 运行中", fg=C["red"])
            self.btn_start.config(text="⏸ 全部暂停", bg="#5c3d1a")
        elif st == "paused":
            self.lbl_state.config(text="🟡 已暂停", fg=C["yellow"])
            self.btn_start.config(text="▶ 继续", bg=C["green"])
        else:
            self.lbl_state.config(text="🟢 就绪", fg=C["green"])
            self.btn_start.config(text="▶ 全部开始", bg=C["green"])

    def _start(self):
        s = self.timer.state
        if s == "running": self.timer.pause_all()
        elif s == "paused": self.timer.resume_all()
        else: self.timer.start_all()

    def _on_on(self, i): self.timer.toggle_on(i); self._refresh()
    def _on_start(self, i):
        sk = self.timer.get()
        if i < len(sk):
            s = sk[i]
            if not s.on: return
            if s.run: s.run = False; self._refresh()
            else: self.timer.start_one(i); self._refresh()
    def _on_reset(self, i): self.timer.reset_one(i); self._refresh()

    def _save(self):
        sk = self.timer.get()
        if not sk: return
        n = simpledialog.askstring("保存方案", "BOSS 名称:", parent=self)
        if n and n.strip():
            n = n.strip()
            data = {"name": n, "skills": [s.to_d() for s in sk]}
            with open(os.path.join(DATA_DIR, f"{n}.json"), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            self.lbl_title.config(text=f"🔥 {n}")

    def _open_config(self):
        ConfigDialog(self, self.timer, None)
        self._rebuild(); self._refresh()

    def _tog_top(self):
        self._top = not self._top
        self.btn_pin.config(fg=C["green"] if self._top else C["muted"])
        try: self.attributes('-topmost', self._top)
        except: pass

    def _tog_clk(self):
        self._clk = not self._clk
        self.btn_clk.config(fg=C["green"] if self._clk else C["muted"])

    def _tog_voice(self):
        self.voice.on = not self.voice.on
        self.lbl_voice.config(fg=C["green"] if self.voice.on else C["dim"],
                              text="🔊" if self.voice.on else "🔇")

    def _open_settings(self):
        dlg = tk.Toplevel(self)
        dlg.title("设置"); dlg.geometry("400x340"); dlg.resizable(False, False)
        dlg.configure(bg=C["bg"]); dlg.transient(self)
        tk.Label(dlg, text="⚙ 语音与显示设置", font=FONT_B, fg=C["fg"], bg=C["bg"]).pack(anchor=tk.W, padx=16, pady=8)
        # voice
        vf = tk.LabelFrame(dlg, text="🔊 语音", font=FONT_S, fg=C["muted"], bg=C["bg"], padx=12, pady=6)
        vf.pack(fill=tk.X, padx=12, pady=4)
        ve = tk.BooleanVar(value=self.voice.on)
        tk.Checkbutton(vf, text="启用语音", variable=ve, font=FONT_S, fg=C["fg"], bg=C["bg"],
                       selectcolor=C["bg"], activebackground=C["bg"],
                       command=lambda: setattr(self.voice, 'on', ve.get())).pack(anchor=tk.W)
        rf = tk.Frame(vf, bg=C["bg"]); rf.pack(fill=tk.X, pady=2)
        tk.Label(rf, text="语速:", font=FONT_S, fg=C["muted"], bg=C["bg"], width=5).pack(side=tk.LEFT)
        rv = tk.IntVar(value=self.voice.rate)
        tk.Scale(rf, from_=-10, to=10, orient=tk.HORIZONTAL, variable=rv, bg=C["bg"], fg=C["fg"],
                 highlightthickness=0, troughcolor=C["border"],
                 command=lambda v: setattr(self.voice, 'rate', int(float(v)))).pack(side=tk.LEFT, fill=tk.X, expand=True)
        wf = tk.Frame(vf, bg=C["bg"]); wf.pack(fill=tk.X, pady=2)
        tk.Label(wf, text="预警:", font=FONT_S, fg=C["muted"], bg=C["bg"], width=5).pack(side=tk.LEFT)
        wv = tk.IntVar(value=self.voice.warn)
        for v, t in [(0, "关"), (3, "3s"), (5, "5s"), (10, "10s")]:
            tk.Radiobutton(wf, text=t, variable=wv, value=v, font=FONT_S, fg=C["muted"], bg=C["bg"],
                           selectcolor=C["bg"], activebackground=C["bg"],
                           command=lambda vv=v: setattr(self.voice, 'warn', vv)).pack(side=tk.LEFT, padx=3)
        tk.Button(vf, text="📢 测试", font=FONT_S, bg=C["border"], fg=C["fg"], relief=tk.FLAT,
                  cursor="hand2", command=self.voice.test).pack(anchor=tk.W, pady=4)
        # display
        df = tk.LabelFrame(dlg, text="🖥 显示", font=FONT_S, fg=C["muted"], bg=C["bg"], padx=12, pady=6)
        df.pack(fill=tk.X, padx=12, pady=4)
        tv = tk.BooleanVar(value=self._top)
        tk.Checkbutton(df, text="窗口置顶", variable=tv, font=FONT_S, fg=C["fg"], bg=C["bg"],
                       selectcolor=C["bg"], activebackground=C["bg"],
                       command=lambda: [setattr(self, '_top', tv.get()), self.attributes('-topmost', tv.get()),
                                        self.btn_pin.config(fg=C["green"] if tv.get() else C["muted"])]
                       ).pack(anchor=tk.W)
        dlg.update_idletasks()
        x = self.winfo_x() + self.winfo_width()//2 - 200
        y = self.winfo_y() + 30
        dlg.geometry(f"+{x}+{y}")

    def _close(self):
        self.timer.stop(); self.destroy()


if __name__ == "__main__":
    app = App(); app.mainloop()
