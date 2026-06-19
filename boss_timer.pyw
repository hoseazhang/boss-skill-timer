#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
BOSS Skill Cooldown Timer v3.2
Cross-platform zero-dependency tkinter app.
Build: pip install pyinstaller && pyinstaller --onefile --windowed boss_timer.pyw
"""
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json, os, time, threading, platform, sys, traceback

# ── Constants ───────────────────────────────────────────────────
APP_NAME    = "BOSS技能倒计时器"
APP_VERSION = "3.2"

FONT_NORMAL = ("Microsoft YaHei", 11) if platform.system() == "Windows" else ("PingFang SC", 13)
FONT_BOLD   = (FONT_NORMAL[0], FONT_NORMAL[1], "bold")
FONT_SMALL  = (FONT_NORMAL[0], FONT_NORMAL[1] - 1)

C = {
    "bg":"#0f0f1a","panel":"#161630","header":"#1a1a35",
    "row":"#12122a","row_warn":"#1a0f0f","border":"#2a2a4a",
    "text":"#e5e5e5","muted":"#a0a0b0","dim":"#555",
    "green":"#22c55e","yellow":"#f59e0b","red":"#ef4444","blue":"#3b82f6","purple":"#8b5cf6",
}

IS_WIN = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PRESETS_DIR = os.path.join(BASE_DIR, "presets")

def log_err(fn):
    """Decorator: log and show exceptions."""
    def wrapper(*a, **kw):
        try: return fn(*a, **kw)
        except Exception as e:
            msg = f"[{APP_NAME}] Error in {fn.__name__}:\n{traceback.format_exc()}"
            print(msg)
            try: messagebox.showerror("错误", str(e))
            except: pass
    return wrapper


# ── Skill Data ───────────────────────────────────────────────────
class Skill:
    __slots__ = ("name","cooldown","enabled","remaining","running")
    def __init__(self, name="技能", cooldown=30, enabled=True):
        self.name=name; self.cooldown=max(1,min(60,int(cooldown)))
        self.enabled=bool(enabled); self.remaining=float(self.cooldown); self.running=False
    def reset(self): self.remaining=float(self.cooldown); self.running=False
    def to_dict(self): return {"name":self.name,"cooldown":self.cooldown,"enabled":self.enabled}
    @classmethod
    def from_dict(cls,d): return cls(d.get("name","技能"),d.get("cooldown",30),d.get("enabled",True))


# ── Voice ───────────────────────────────────────────────────────
class VoiceEngine:
    def __init__(self):
        self.enabled=True; self.rate=2; self.volume=100; self.early_warning=0; self._voice=None
        if IS_WIN:
            try:
                import win32com.client
                self._voice=win32com.client.Dispatch("SAPI.SpVoice")
            except: pass
    def speak(self,text):
        if not self.enabled or not text: return
        def _run():
            try:
                if IS_WIN and self._voice:
                    self._voice.Rate=self.rate; self._voice.Volume=self.volume; self._voice.Speak(text,1)
                elif IS_MAC: os.system(f'say -v Tingting "{text}" &')
            except: pass
        threading.Thread(target=_run,daemon=True).start()
    def speak_skill(self,name): self.speak(name)
    def speak_warning(self,name): self.speak(f"准备 {name}")
    def test(self): self.speak(f"{APP_NAME} 语音测试")


# ── Presets ─────────────────────────────────────────────────────
class PresetManager:
    def __init__(self):
        os.makedirs(PRESETS_DIR,exist_ok=True); self._ensure_defaults()
    def _ensure_defaults(self):
        dp=os.path.join(PRESETS_DIR,"demo_boss.json")
        tp=os.path.join(PRESETS_DIR,"new_boss_template.json")
        if not os.path.exists(dp):
            self._write(dp,{"name":"炎狱之王","game":"演示副本","notes":"P1技能轴",
                "skills":[{"name":"火焰吐息","cooldown":30},{"name":"暗影冲锋","cooldown":20},
                          {"name":"陨石坠落","cooldown":90},{"name":"冰霜禁锢","cooldown":20}]})
        if not os.path.exists(tp):
            self._write(tp,{"name":"新BOSS","skills":[{"name":"技能1","cooldown":30},{"name":"技能2","cooldown":20}]})
    def _write(self,path,data):
        with open(path,"w",encoding="utf-8") as f: json.dump(data,f,ensure_ascii=False,indent=2)
    def list(self):
        if not os.path.exists(PRESETS_DIR): return []
        return sorted(f for f in os.listdir(PRESETS_DIR) if f.endswith(".json"))
    def load(self,fn):
        p=os.path.join(PRESETS_DIR,fn)
        if not os.path.exists(p): return None
        with open(p,"r",encoding="utf-8") as f: d=json.load(f)
        return {"filename":fn,"name":d.get("name",""),"game":d.get("game",""),"notes":d.get("notes",""),
                "skills":[Skill.from_dict(s) for s in d.get("skills",[])]}
    def save(self,fn,name,skills,game="",notes=""):
        self._write(os.path.join(PRESETS_DIR,fn),{"name":name,"game":game,"notes":notes,
                    "skills":[s.to_dict() for s in skills]})
    def delete(self,fn):
        p=os.path.join(PRESETS_DIR,fn)
        if os.path.exists(p): os.remove(p)


# ── Timer ───────────────────────────────────────────────────────
class TimerEngine:
    def __init__(self,iv=0.1):
        self._sk=[]; self._iv=iv; self._lk=threading.Lock()
        self._tr=False; self._th=None; self._wt={}; self.state="idle"
        self.on_tick=lambda s:None; self.on_done=lambda i,s:None; self.on_state=lambda s:None
    def get(self):
        with self._lk: return list(self._sk)
    def set(self,skills):
        with self._lk: self._sk=[Skill(s.name,s.cooldown,s.enabled) for s in skills]
    def add(self,name="新技能",cd=30):
        with self._lk: s=Skill(name,cd); self._sk.append(s); return len(self._sk)-1
    def rm(self,i):
        with self._lk:
            if 0<=i<len(self._sk): self._sk.pop(i)
    # global
    def start_all(self):
        with self._lk:
            if self.state=="running": return
            self.state="running"
            for s in self._sk:
                if s.enabled: s.running=True
                if s.remaining<=0: s.remaining=float(s.cooldown)
            self._wt.clear(); self._go(); self.on_state("running")
    def pause_all(self):
        with self._lk:
            if self.state!="running": return
            self.state="paused"; self._stop()
            for s in self._sk: s.running=False
            self.on_state("paused")
    def resume_all(self):
        with self._lk:
            if self.state!="paused": return
            self.state="running"
            for s in self._sk:
                if s.enabled: s.running=True
            self._wt.clear(); self._go(); self.on_state("running")
    def reset_all(self):
        with self._lk:
            self._stop(); self.state="idle"; self._wt.clear()
            for s in self._sk: s.reset()
            self.on_state("idle"); self.on_tick(self._sk)
    # per-skill
    def start_one(self,i):
        with self._lk:
            if 0<=i<len(self._sk):
                s=self._sk[i]
                if not s.enabled: return
                s.running=True; s.remaining=float(s.cooldown)
                self._wt.pop(i,None); self._go()
    def reset_one(self,i):
        with self._lk:
            if 0<=i<len(self._sk):
                s=self._sk[i]; s.remaining=float(s.cooldown)
                s.running=(self.state=="running" and s.enabled)
                self._wt.pop(i,None)
    def toggle_on(self,i):
        with self._lk:
            if 0<=i<len(self._sk):
                s=self._sk[i]; s.enabled=not s.enabled
                if not s.enabled: s.running=False; s.remaining=float(s.cooldown)
                elif self.state=="running": s.running=True; s.remaining=float(s.cooldown)
    def upd(self,i,name=None,cd=None):
        with self._lk:
            if 0<=i<len(self._sk):
                s=self._sk[i]
                if name is not None: s.name=name
                if cd is not None: s.cooldown=max(1,min(60,int(cd))); s.remaining=float(s.cooldown)
    # internal
    def _go(self):
        if self._tr: return
        self._tr=True; self._th=threading.Thread(target=self._loop,daemon=True); self._th.start()
    def _stop(self): self._tr=False
    def _loop(self):
        while self._tr:
            t0=time.time()
            with self._lk:
                done=[]
                for i,s in enumerate(self._sk):
                    if s.running and s.enabled:
                        s.remaining=max(0.0,s.remaining-self._iv)
                        if s.remaining<=0: s.remaining=0.0; s.running=False; done.append((i,s))
                for i,s in done:
                    self.on_done(i,s); s.remaining=float(s.cooldown)
                self.on_tick(self._sk)
                if self.state!="running" and not any(s.running for s in self._sk):
                    self._tr=False; break
            time.sleep(max(0.0,self._iv-(time.time()-t0)))
        self._tr=False
    def check_warn(self,ws):
        if ws<=0: return None,None
        with self._lk:
            for i,s in enumerate(self._sk):
                if s.running and s.enabled and 0<s.remaining<=ws:
                    l=self._wt.get(i,999)
                    if l>ws: self._wt[i]=s.remaining; return i,s
                    self._wt[i]=s.remaining
        return None,None
    def stop(self): self._stop()


# ── Overlay (Windows only) ──────────────────────────────────────
class Overlay:
    @staticmethod
    def apply(win,top=True,clk=False,op=0.85):
        if not IS_WIN: return
        try:
            import ctypes; h=win.winfo_id(); u=ctypes.windll.user32
            s=u.GetWindowLongW(h,-20); s|=0x80000
            if top: s|=0x8
            if clk: s|=0x20
            u.SetWindowLongW(h,-20,s); u.SetLayeredWindowAttributes(h,0,int(op*255),2)
        except Exception as e: print(f"Overlay.apply error: {e}")
    @staticmethod
    def set_clk(win,on):
        if not IS_WIN: return
        try:
            import ctypes; h=win.winfo_id(); u=ctypes.windll.user32
            s=u.GetWindowLongW(h,-20); s=(s|0x20) if on else (s&~0x20); u.SetWindowLongW(h,-20,s)
        except: pass
    @staticmethod
    def set_op(win,op):
        if not IS_WIN: return
        try:
            import ctypes; h=win.winfo_id()
            ctypes.windll.user32.SetLayeredWindowAttributes(h,0,int(op*255),2)
        except: pass


# ── Skill Row ───────────────────────────────────────────────────
class SkillRow(tk.Frame):
    def __init__(self,p,i,s,cb):
        super().__init__(p,bg=C["row"],height=52); self.idx=i; self.skill=s; self.cb=cb
        self.pack_propagate(False); self.pack(fill=tk.X,padx=2,pady=1); self._b()
    def _b(self):
        s=self.skill; bg=C["row"]
        self.dot=tk.Label(self,text="●",font=FONT_NORMAL,cursor="hand2",bg=bg,fg=C["green"])
        self.dot.pack(side=tk.LEFT,padx=(6,4),pady=10)
        self.dot.bind("<Button-1>",lambda e:self.cb("on")(self.idx))
        self.ln=tk.Label(self,text=s.name,font=FONT_BOLD,bg=bg,fg=C["text"],width=12,anchor=tk.W)
        self.ln.pack(side=tk.LEFT,padx=4,pady=10)
        self.lc=tk.Label(self,text=f"{s.cooldown}s",font=FONT_NORMAL,bg=bg,fg=C["muted"],width=5)
        self.lc.pack(side=tk.LEFT,padx=4,pady=10)
        self.cv=tk.Canvas(self,bg=C["bg"],height=20,highlightthickness=1,highlightbackground=C["border"])
        self.cv.pack(side=tk.LEFT,fill=tk.X,expand=True,padx=6,pady=10)
        self.lt=tk.Label(self,text=f"{s.cooldown:.1f}s",font=FONT_BOLD,bg=bg,fg=C["green"],width=7)
        self.lt.pack(side=tk.LEFT,padx=4,pady=10)
        self.bs=tk.Label(self,text="▶ 开始",font=FONT_SMALL,cursor="hand2",bg=C["green"],fg="#fff",padx=8,pady=2)
        self.bs.pack(side=tk.LEFT,padx=2,pady=10)
        self.bs.bind("<Button-1>",lambda e:self.cb("start")(self.idx))
        self.br=tk.Label(self,text="🔄 重置",font=FONT_SMALL,cursor="hand2",bg="#5c1a1a",fg=C["red"],padx=8,pady=2)
        self.br.pack(side=tk.LEFT,padx=2,pady=10)
        self.br.bind("<Button-1>",lambda e:self.cb("reset")(self.idx))
    def rf(self,skill):
        self.skill=skill; s=skill
        d=C["dim"]; m=C["muted"]; en=s.enabled
        self.dot.config(text="●" if en else "○",fg=C["green"] if en else d)
        self.ln.config(text=s.name,fg=d if not en else C["text"])
        self.lc.config(text=f"{s.cooldown}s",fg=d if not en else m)
        if not en:
            self.lt.config(text="--",fg=d); self._bar(0,d)
        elif not s.running:
            self.lt.config(text=f"{s.cooldown:.1f}s",fg=C["green"]); self._bar(1,C["green"])
        else:
            r=max(0.0,s.remaining); rt=r/s.cooldown if s.cooldown else 0
            cl=C["red"] if r<=5 else (C["yellow"] if r<=10 else C["green"])
            self.lt.config(text=f"{r:.1f}s",fg=cl); self._bar(rt,cl)
        if not en:
            self.bs.config(bg=d,fg=C["bg"],text="▶ 开始")
        elif s.running:
            self.bs.config(bg="#5c3d1a",fg=C["yellow"],text="⏸ 暂停")
        else:
            self.bs.config(bg=C["green"],fg="#fff",text="▶ 开始")
        nb=C["row_warn"] if (s.running and en and s.remaining<=5) else C["row"]
        self.config(bg=nb)
        for w in(self.dot,self.ln,self.lt):
            try: w.config(bg=nb)
            except: pass
    def _bar(self,rt,cl):
        self.cv.delete("all"); w=self.cv.winfo_width(); h=self.cv.winfo_height()
        if w<5: return
        fw=int(w*max(0.0,min(1.0,rt)))
        if fw>0: self.cv.create_rectangle(0,0,fw,h,fill=cl,outline="")
        if fw<w: self.cv.create_rectangle(fw,0,w,h,fill=C["bg"],outline="")


# ── Config Dialog ───────────────────────────────────────────────
class ConfigDialog(tk.Toplevel):
    """Full config: edit skill names, cooldowns, add/remove skills, save/load presets."""
    def __init__(self,parent,timer,presets,on_config_done):
        super().__init__(parent)
        self.timer=timer; self.presets=presets; self.done=on_config_done
        self.title("⚙ 配置方案"); self.geometry("520x480"); self.resizable(False,False)
        self.configure(bg=C["bg"]); self.transient(parent)
        self._skills=[]; self._build()
        self.update_idletasks()
        x=parent.winfo_x()+parent.winfo_width()//2-260; y=parent.winfo_y()+20
        self.geometry(f"+{x}+{y}")
        self.grab_set()
    def _build(self):
        p={"padx":12,"pady":4}
        # ── preset selector ──
        pf=tk.Frame(self,bg=C["panel"]); pf.pack(fill=tk.X,padx=10,pady=(10,0))
        tk.Label(pf,text="方案:",font=FONT_SMALL,fg=C["muted"],bg=C["panel"]).pack(side=tk.LEFT,**p)
        flist=self.presets.list()
        self.fvar=tk.StringVar(value=flist[0] if flist else "")
        self.fmenu=tk.OptionMenu(pf,self.fvar,flist[0] if flist else "",*flist,command=self._on_load)
        self.fmenu.config(font=FONT_SMALL,bg=C["bg"],fg=C["muted"],highlightthickness=0,width=20)
        self.fmenu.pack(side=tk.LEFT,padx=4)
        self.ename=tk.StringVar()
        tk.Entry(pf,textvariable=self.ename,font=FONT_SMALL,bg=C["border"],fg=C["text"],
                 insertbackground=C["text"],width=16).pack(side=tk.LEFT,padx=4)

        # ── skill list ──
        sf=tk.Frame(self,bg=C["bg"]); sf.pack(fill=tk.BOTH,expand=True,padx=10,pady=6)
        cv=tk.Canvas(sf,bg=C["bg"],highlightthickness=0); sb=tk.Scrollbar(sf,command=cv.yview)
        self.sf_inner=tk.Frame(cv,bg=C["bg"])
        self.sf_inner.bind("<Configure>",lambda e:cv.configure(scrollregion=cv.bbox("all")))
        cv.create_window((0,0),window=self.sf_inner,anchor=tk.NW)
        cv.configure(yscrollcommand=sb.set); cv.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        sb.pack(side=tk.RIGHT,fill=tk.Y)
        # header
        hf=tk.Frame(self.sf_inner,bg=C["header"]); hf.pack(fill=tk.X,pady=2)
        for t,w in[("技能名",28),("冷却(s)",8),("启用",5),("",6)]:
            tk.Label(hf,text=t,font=FONT_SMALL,fg=C["muted"],bg=C["header"],width=w).pack(side=tk.LEFT,padx=2)
        # rows built in _show_skills
        self._row_frames=[]

        # ── buttons ──
        bf=tk.Frame(self,bg=C["panel"]); bf.pack(fill=tk.X,padx=10,pady=8)
        tk.Label(bf,text="+ 添加",font=FONT_SMALL,bg=C["blue"],fg=C["text"],cursor="hand2",
                 padx=12,pady=4).pack(side=tk.LEFT,padx=4)
        bf.winfo_children()[-1].bind("<Button-1>",lambda e:self._add_skill())
        tk.Label(bf,text="💾 保存",font=FONT_SMALL,bg=C["green"],fg="#fff",cursor="hand2",
                 padx=12,pady=4).pack(side=tk.LEFT,padx=4)
        bf.winfo_children()[-1].bind("<Button-1>",lambda e:self._save())
        tk.Label(bf,text="🗑 删除方案",font=FONT_SMALL,bg="#5c1a1a",fg=C["red"],cursor="hand2",
                 padx=12,pady=4).pack(side=tk.LEFT,padx=4)
        bf.winfo_children()[-1].bind("<Button-1>",lambda e:self._delete_preset())
        tk.Label(bf,text="关闭",font=FONT_SMALL,bg=C["border"],fg=C["text"],cursor="hand2",
                 padx=16,pady=4).pack(side=tk.RIGHT,padx=4)
        bf.winfo_children()[-1].bind("<Button-1>",lambda e:self.destroy())

        if flist: self._on_load(flist[0])

    def _on_load(self,fn):
        if not fn: return
        d=self.presets.load(fn)
        if d:
            self._skills=d["skills"]
            self.ename.set(d["name"])
            self._show_skills()

    def _show_skills(self):
        for f in self._row_frames: f.destroy()
        self._row_frames.clear()
        for i,s in enumerate(self._skills):
            fr=tk.Frame(self.sf_inner,bg=C["row"]); fr.pack(fill=tk.X,padx=2,pady=1)
            self._row_frames.append(fr)
            # name
            sv=tk.StringVar(value=s.name)
            tk.Entry(fr,textvariable=sv,font=FONT_SMALL,bg=C["border"],fg=C["text"],
                     insertbackground=C["text"],width=18).pack(side=tk.LEFT,padx=2,pady=4)
            # cooldown
            cv=tk.IntVar(value=s.cooldown)
            tk.Spinbox(fr,textvariable=cv,from_=1,to=60,font=FONT_SMALL,
                       bg=C["border"],fg=C["text"],width=5,buttonbackground=C["border"]).pack(side=tk.LEFT,padx=2,pady=4)
            # enabled
            ev=tk.BooleanVar(value=s.enabled)
            tk.Checkbutton(fr,variable=ev,bg=C["row"],fg=C["text"],
                           selectcolor=C["row"],activebackground=C["row"]).pack(side=tk.LEFT,padx=4)
            # delete
            tk.Label(fr,text="✕",font=FONT_NORMAL,cursor="hand2",bg=C["row"],fg=C["red"],width=2).pack(side=tk.LEFT,padx=4)
            fr.winfo_children()[-1].bind("<Button-1>",lambda e,idx=i: self._remove_skill(idx))
            # store refs
            fr._name_var=sv; fr._cd_var=cv; fr._en_var=ev

    def _add_skill(self):
        self._skills.append(Skill("新技能",30))
        self._show_skills()

    def _remove_skill(self,idx):
        if 0<=idx<len(self._skills):
            self._skills.pop(idx); self._show_skills()

    def _save(self):
        name=self.ename.get().strip()
        if not name: messagebox.showwarning("提示","请输入方案名称",parent=self); return
        # collect from UI
        for i,fr in enumerate(self._row_frames):
            if i<len(self._skills):
                s=self._skills[i]
                s.name=fr._name_var.get().strip() or s.name
                try: s.cooldown=max(1,min(60,fr._cd_var.get()))
                except: pass
                s.enabled=fr._en_var.get()
        fn=f"{name}.json"
        self.presets.save(fn,name,self._skills)
        self.done(name,self._skills)  # signal main app
        self.destroy()

    def _delete_preset(self):
        fn=self.fvar.get()
        if not fn: return
        if messagebox.askyesno("确认删除",f"删除方案「{fn}」?",parent=self):
            self.presets.delete(fn)
            flist=self.presets.list()
            menu=self.fmenu["menu"]; menu.delete(0,"end")
            for f in flist: menu.add_command(label=f,command=lambda v=f: self.fvar.set(v) or self._on_load(v))
            if flist: self.fvar.set(flist[0]); self._on_load(flist[0])
            else: self.fvar.set(""); self.ename.set(""); self._skills=[]; self._show_skills()


# ── Settings ────────────────────────────────────────────────────
class SettingsDialog(tk.Toplevel):
    def __init__(self,parent,voice,cb):
        super().__init__(parent); self.voice=voice; self.cb=cb
        self.title("设置"); self.geometry("420x380"); self.resizable(False,False)
        self.configure(bg=C["bg"]); self.transient(parent); self._build()
        self.update_idletasks()
        x=parent.winfo_x()+parent.winfo_width()//2-210; y=parent.winfo_y()+30
        self.geometry(f"+{x}+{y}")
    def _build(self):
        pad={"padx":16,"pady":6}
        tk.Label(self,text="⚙ 语音与显示设置",font=FONT_BOLD,fg=C["text"],bg=C["bg"]).pack(anchor=tk.W,**pad)
        vf=tk.LabelFrame(self,text="🔊 语音播报",font=FONT_SMALL,fg=C["muted"],bg=C["bg"],padx=12,pady=8)
        vf.pack(fill=tk.X,padx=14,pady=4)
        self.ve=tk.BooleanVar(value=self.voice.enabled)
        tk.Checkbutton(vf,text="启用语音播报",variable=self.ve,font=FONT_SMALL,fg=C["text"],bg=C["bg"],
                       selectcolor=C["bg"],activebackground=C["bg"],
                       command=lambda:setattr(self.voice,'enabled',self.ve.get())).pack(anchor=tk.W)
        rf=tk.Frame(vf,bg=C["bg"]); rf.pack(fill=tk.X,pady=2)
        tk.Label(rf,text="语速:",font=FONT_SMALL,fg=C["muted"],bg=C["bg"],width=5).pack(side=tk.LEFT)
        self.rv=tk.IntVar(value=self.voice.rate)
        tk.Scale(rf,from_=-10,to=10,orient=tk.HORIZONTAL,variable=self.rv,bg=C["bg"],fg=C["text"],
                 highlightthickness=0,troughcolor=C["border"],
                 command=lambda v:setattr(self.voice,'rate',int(float(v)))).pack(side=tk.LEFT,fill=tk.X,expand=True)
        wf=tk.Frame(vf,bg=C["bg"]); wf.pack(fill=tk.X,pady=2)
        tk.Label(wf,text="预警:",font=FONT_SMALL,fg=C["muted"],bg=C["bg"],width=5).pack(side=tk.LEFT)
        self.wv=tk.IntVar(value=self.voice.early_warning)
        for v,t in[(0,"关"),(3,"3s"),(5,"5s"),(10,"10s")]:
            tk.Radiobutton(wf,text=t,variable=self.wv,value=v,font=FONT_SMALL,fg=C["muted"],bg=C["bg"],
                           selectcolor=C["bg"],activebackground=C["bg"],
                           command=lambda vv=v:setattr(self.voice,'early_warning',vv)).pack(side=tk.LEFT,padx=3)
        tk.Button(vf,text="📢 测试语音",font=FONT_SMALL,bg=C["border"],fg=C["text"],relief=tk.FLAT,
                  cursor="hand2",command=self.voice.test,padx=12,pady=3).pack(anchor=tk.W,pady=4)
        df=tk.LabelFrame(self,text="🖥 显示设置",font=FONT_SMALL,fg=C["muted"],bg=C["bg"],padx=12,pady=8)
        df.pack(fill=tk.X,padx=14,pady=4)
        self.tv=tk.BooleanVar(value=True)
        tk.Checkbutton(df,text="窗口置顶",variable=self.tv,font=FONT_SMALL,fg=C["text"],bg=C["bg"],
                       selectcolor=C["bg"],activebackground=C["bg"],
                       command=lambda:self.cb("top")(self.tv.get())).pack(anchor=tk.W)
        self.cv=tk.BooleanVar(value=False)
        tk.Checkbutton(df,text="🖱 点击穿透",variable=self.cv,font=FONT_SMALL,fg=C["text"],bg=C["bg"],
                       selectcolor=C["bg"],activebackground=C["bg"],
                       command=lambda:self.cb("clk")(self.cv.get())).pack(anchor=tk.W)
        of=tk.Frame(df,bg=C["bg"]); of.pack(fill=tk.X,pady=2)
        tk.Label(of,text="透明度:",font=FONT_SMALL,fg=C["muted"],bg=C["bg"],width=5).pack(side=tk.LEFT)
        self.ov=tk.IntVar(value=85); self.lo=tk.Label(of,text="85%",font=FONT_SMALL,fg=C["muted"],bg=C["bg"],width=4)
        self.lo.pack(side=tk.RIGHT)
        tk.Scale(of,from_=30,to=100,orient=tk.HORIZONTAL,variable=self.ov,bg=C["bg"],fg=C["text"],
                 highlightthickness=0,troughcolor=C["border"],
                 command=lambda v:(self.lo.config(text=f"{int(float(v))}%"),self.cb("op")(int(float(v))/100))
                 ).pack(side=tk.LEFT,fill=tk.X,expand=True)


# ── Main App ────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_NAME); self.geometry("780x520"); self.minsize(600,320)
        self.configure(bg=C["bg"])
        self.timer=TimerEngine(0.1); self.voice=VoiceEngine(); self.presets=PresetManager()
        self.timer.on_tick=lambda s:self.after(0,self._refresh)
        self.timer.on_done=lambda i,s:self.after(0,lambda:self._on_done(i,s))
        self.timer.on_state=lambda s:self.after(0,lambda:self._on_state(s))
        self.rows=[]; self.cur_file=None; self._top=True; self._clk=False
        self._build(); self._load_preset("demo_boss.json")
        # NO automatic overlay — was causing blank window on Windows
        self.protocol("WM_DELETE_WINDOW",self._close)

    def _build(self):
        # title bar
        tb=tk.Frame(self,bg="#1a1030",height=36); tb.pack(fill=tk.X); tb.pack_propagate(False)
        self.lbl_boss=tk.Label(tb,text=f"🔥 {APP_NAME}",font=FONT_BOLD,fg=C["yellow"],bg="#1a1030")
        self.lbl_boss.pack(side=tk.LEFT,padx=12,pady=4)
        bf=tk.Frame(tb,bg="#1a1030"); bf.pack(side=tk.RIGHT,padx=8)
        self.btn_pin=tk.Label(bf,text="📌",font=("",13),cursor="hand2",bg="#1a1030",fg=C["green"],width=3)
        self.btn_pin.pack(side=tk.LEFT); self.btn_pin.bind("<Button-1>",lambda e:self._tog_top())
        self.btn_clk=tk.Label(bf,text="🖱",font=("",13),cursor="hand2",bg="#1a1030",fg=C["muted"],width=3)
        self.btn_clk.pack(side=tk.LEFT); self.btn_clk.bind("<Button-1>",lambda e:self._tog_clk())
        self.btn_cfg=tk.Label(bf,text="⚙",font=("",13),cursor="hand2",bg="#1a1030",fg=C["muted"],width=3)
        self.btn_cfg.pack(side=tk.LEFT); self.btn_cfg.bind("<Button-1>",lambda e:self._open_settings())
        # header
        hf=tk.Frame(self,bg=C["header"],height=26); hf.pack(fill=tk.X,padx=4,pady=(4,0)); hf.pack_propagate(False)
        for t,w in[("启用",5),("技能名称",14),("冷却",5),("倒计时进度",38),("剩余",6),("操作",14)]:
            tk.Label(hf,text=t,font=FONT_SMALL,fg=C["muted"],bg=C["header"],width=w).pack(side=tk.LEFT,padx=2,pady=3)
        # scrollable list
        self.cv_list=tk.Canvas(self,bg=C["bg"],highlightthickness=0)
        self.sb_list=tk.Scrollbar(self,command=self.cv_list.yview)
        self.fr_list=tk.Frame(self.cv_list,bg=C["bg"])
        self.fr_list.bind("<Configure>",lambda e:self.cv_list.configure(scrollregion=self.cv_list.bbox("all")))
        self.cv_list.create_window((0,0),window=self.fr_list,anchor=tk.NW)
        self.cv_list.configure(yscrollcommand=self.sb_list.set)
        self.cv_list.pack(side=tk.LEFT,fill=tk.BOTH,expand=True,padx=4)
        self.sb_list.pack(side=tk.RIGHT,fill=tk.Y)
        self.cv_list.bind_all("<MouseWheel>",lambda e:self.cv_list.yview_scroll(int(-e.delta/120),"units"))
        # bottom bar
        bb=tk.Frame(self,bg=C["panel"],height=72); bb.pack(fill=tk.X,side=tk.BOTTOM,padx=4,pady=4); bb.pack_propagate(False)
        lf=tk.Frame(bb,bg=C["panel"]); lf.pack(side=tk.LEFT,padx=8,pady=12)
        tk.Label(lf,text="+ 添加",font=FONT_SMALL,bg=C["blue"],fg=C["text"],cursor="hand2",padx=12,pady=4).pack(side=tk.LEFT,padx=4)
        lf.winfo_children()[-1].bind("<Button-1>",lambda e:self._add_skill())
        tk.Label(lf,text="🔧 配置",font=FONT_SMALL,bg=C["purple"],fg=C["text"],cursor="hand2",padx=12,pady=4).pack(side=tk.LEFT,padx=4)
        lf.winfo_children()[-1].bind("<Button-1>",lambda e:self._open_config())
        tk.Label(lf,text="💾 保存",font=FONT_SMALL,bg=C["border"],fg=C["text"],cursor="hand2",padx=10,pady=4).pack(side=tk.LEFT,padx=4)
        lf.winfo_children()[-1].bind("<Button-1>",lambda e:self._save())
        # global controls
        rf=tk.Frame(bb,bg=C["panel"]); rf.pack(side=tk.RIGHT,padx=8,pady=12)
        self.btn_start=tk.Label(rf,text="▶ 全部开始",font=FONT_BOLD,bg=C["green"],fg="#fff",cursor="hand2",padx=14,pady=5)
        self.btn_start.pack(side=tk.LEFT,padx=3)
        self.btn_start.bind("<Button-1>",lambda e:self._start_pause())
        tk.Label(rf,text="⏸ 暂停",font=FONT_BOLD,bg="#5c3d1a",fg=C["yellow"],cursor="hand2",padx=12,pady=5).pack(side=tk.LEFT,padx=3)
        rf.winfo_children()[-1].bind("<Button-1>",lambda e:self.timer.pause_all())
        tk.Label(rf,text="🔄 重置",font=FONT_BOLD,bg="#5c1a1a",fg=C["red"],cursor="hand2",padx=12,pady=5).pack(side=tk.LEFT,padx=3)
        rf.winfo_children()[-1].bind("<Button-1>",lambda e:self.timer.reset_all())
        self.lbl_voice=tk.Label(bb,text="🔊",font=("",15),bg=C["panel"],fg=C["green"],cursor="hand2")
        self.lbl_voice.pack(side=tk.RIGHT,padx=6,pady=18)
        self.lbl_voice.bind("<Button-1>",lambda e:self._tog_voice())
        self.lbl_state=tk.Label(bb,text="🟢 准备就绪",font=FONT_SMALL,bg=C["panel"],fg=C["muted"])
        self.lbl_state.pack(side=tk.RIGHT,padx=12,pady=18)

    # ── rows ──
    def _rebuild(self):
        for r in self.rows: r.destroy()
        self.rows.clear()
        cb={"on":self._on_en,"start":self._on_start,"reset":self._on_reset}
        for i,s in enumerate(self.timer.get()): self.rows.append(SkillRow(self.fr_list,i,s,cb))
        self.fr_list.update_idletasks()
        self.cv_list.configure(scrollregion=self.cv_list.bbox("all"))

    def _refresh(self):
        sk=self.timer.get()
        for i,r in enumerate(self.rows):
            if i<len(sk): r.rf(sk[i])
        self._check_warn()

    def _check_warn(self):
        w=self.voice.early_warning
        if w>0:
            _,s=self.timer.check_warn(w)
            if s: self.voice.speak_warning(s.name)

    # ── events ──
    def _on_done(self,i,s): self.voice.speak_skill(s.name)
    def _on_state(self,st):
        if st=="running": self.lbl_state.config(text="🔴 运行中",fg=C["red"]); self.btn_start.config(text="⏸ 全部暂停",bg="#5c3d1a",fg=C["yellow"])
        elif st=="paused": self.lbl_state.config(text="🟡 已暂停",fg=C["yellow"]); self.btn_start.config(text="▶ 继续",bg=C["green"],fg="#fff")
        else: self.lbl_state.config(text="🟢 准备就绪",fg=C["green"]); self.btn_start.config(text="▶ 全部开始",bg=C["green"],fg="#fff")
    def _start_pause(self):
        s=self.timer.state
        if s=="running": self.timer.pause_all()
        elif s=="paused": self.timer.resume_all()
        else: self.timer.start_all()
    def _on_en(self,i): self.timer.toggle_on(i); self._refresh()
    def _on_start(self,i):
        sk=self.timer.get()
        if i<len(sk):
            s=sk[i]
            if not s.enabled: return
            if s.running: s.running=False; self._refresh()
            else: self.timer.start_one(i); self._refresh()
    def _on_reset(self,i): self.timer.reset_one(i); self._refresh()
    def _add_skill(self): self.timer.add(); self._rebuild(); self._refresh()

    def _open_config(self):
        def done(name,skills):
            self.timer.reset_all(); self.timer.set(skills)
            self.cur_file=f"{name}.json"; self.lbl_boss.config(text=f"🔥 {name}")
            self._rebuild(); self._refresh()
        ConfigDialog(self,self.timer,self.presets,done)

    def _save(self):
        sk=self.timer.get()
        if not sk: messagebox.showwarning("提示","没有技能",parent=self); return
        n=simpledialog.askstring("保存方案","BOSS 名称:",parent=self)
        if n and n.strip():
            n=n.strip(); self.presets.save(f"{n}.json",n,sk)
            self.cur_file=f"{n}.json"; self.lbl_boss.config(text=f"🔥 {n}")

    def _load_preset(self,fn):
        if not fn or fn=="无方案": return
        d=self.presets.load(fn)
        if d:
            self.timer.reset_all(); self.timer.set(d["skills"])
            self.cur_file=fn; self.lbl_boss.config(text=f"🔥 {d['name']}")
            self._rebuild(); self._refresh()

    # ── toggles ──
    def _tog_top(self):
        self._top=not self._top; self.btn_pin.config(fg=C["green"] if self._top else C["muted"])
        try: self.attributes('-topmost',self._top)
        except: pass
    def _tog_clk(self):
        self._clk=not self._clk; self.btn_clk.config(fg=C["green"] if self._clk else C["muted"])
        Overlay.set_clk(self,self._clk)
    def _tog_voice(self):
        self.voice.enabled=not self.voice.enabled
        self.lbl_voice.config(fg=C["green"] if self.voice.enabled else C["dim"],
                              text="🔊" if self.voice.enabled else "🔇")
    def _open_settings(self):
        SettingsDialog(self,self.voice,{
            "top":lambda v:(setattr(self,'_top',v),self.btn_pin.config(fg=C["green"] if v else C["muted"]),self.attributes('-topmost',v)),
            "clk":lambda v:(setattr(self,'_clk',v),self.btn_clk.config(fg=C["green"] if v else C["muted"]),Overlay.set_clk(self,v)),
            "op":lambda v:Overlay.set_op(self,v)
        })
    def _close(self):
        self.timer.stop(); self.destroy()


if __name__=="__main__":
    app=App(); app.mainloop()
