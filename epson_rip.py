import subprocess, os, sys, tempfile, shutil, threading, json, time, math
from pathlib import Path
from PIL import Image, ImageTk, ImageOps, ImageDraw
Image.MAX_IMAGE_PIXELS = None
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog

SCREENRIP_VERSION = "3.1.0"
CONFIG_PATH = Path.home() / ".screenrip_config.json"
DEFAULT_CONFIG = {
    "save_folder": "",
    "recent_files": [],
    "presets": {
        "1430 Fine Detail": {
            "lpi":65,"angle":22.5,"shape":"0 round","dpi":1440,"sat":200,
            "mode":"Multi Black","drop":"1 small","wcor":0.0,"lcor":0.0,
            "quality":"High","media_type":"Sheet","media_w":13.0,"media_h":19.0,
            "paper_type":"Matte Film","enhance":"None",
            "nup_cols":2,"nup_rows":2,"nup_gap":0.25,"shirt":"White"
        },
        "T3170 Wide Format": {
            "lpi":45,"angle":22.5,"shape":"0 round","dpi":720,"sat":200,
            "mode":"Multi Black","drop":"1 small","wcor":0.0,"lcor":0.0,
            "quality":"High","media_type":"Roll","media_w":24.0,"media_h":36.0,
            "paper_type":"Matte Film","enhance":"None",
            "nup_cols":2,"nup_rows":3,"nup_gap":0.25,"shirt":"White"
        },
    },
    "last_preset":"1430 Fine Detail"
}
BG="#2d2d2d"; BG2="#383838"; BG3="#222222"; BG4="#1a1a1a"
FG="#e0e0e0"; FG2="#aaaaaa"; FG3="#666666"
GREEN="#00c853"; GREEN2="#00e676"; DARKGREEN="#003d18"; BORDER="#1a1a1a"

def load_config():
    try:
        if CONFIG_PATH.exists():
            with open(CONFIG_PATH) as f:
                c=json.load(f)
                for k,v in DEFAULT_CONFIG.items():
                    if k not in c: c[k]=v
                return c
    except: pass
    return dict(DEFAULT_CONFIG)

def save_config(cfg):
    try:
        with open(CONFIG_PATH,"w") as f: json.dump(cfg,f,indent=2)
    except: pass

def find_gs():
    for g in ["/opt/homebrew/bin/gs","/usr/local/bin/gs","gs"]:
        if shutil.which(g): return g
    return None

def gs_env():
    e=os.environ.copy(); e["OBJC_DISABLE_INITIALIZE_FORK_SAFETY"]="YES"; return e

def render_pdf(gs,pdf,out,dpi=1440,lpi=55,angle=22.5,shape="0 round",
               saturation=200,mode="Multi Black",wcor=0.0,lcor=0.0,quality="High",cb=None):
    tmp_in=tempfile.mktemp(suffix=".pdf"); tmp_out=tempfile.mktemp(suffix=".tiff")
    shutil.copy2(pdf,tmp_in)
    q_flags={"High":["-dGraphicsAlphaBits=4","-dTextAlphaBits=4"],"Medium":[],"Draft":[]}.get(quality,[])
    cmd=([gs,"-dBATCH","-dNOPAUSE","-dSAFER","-sDEVICE=tiffgray",
          f"-r{dpi}",f"-dDefaultScreenFrequency={lpi}",f"-dDefaultScreenAngle={angle}",
          f"-sOutputFile={tmp_out}"]+q_flags+[tmp_in])
    if cb: cb("Rendering with Ghostscript...")
    r=subprocess.run(cmd,capture_output=True,text=True,env=gs_env())
    try: os.remove(tmp_in)
    except: pass
    if r.returncode!=0: raise RuntimeError(r.stderr[-800:])
    if cb: cb("Processing image...")
    img=Image.open(tmp_out).convert("L"); img.load()
    try:
        import numpy as np; arr=np.array(img)
        dark_pct=(arr<128).sum()/arr.size
    except:
        px=list(img.getdata()); dark_pct=sum(1 for p in px if p<128)/len(px)
    if dark_pct>0.5: img=ImageOps.invert(img)
    from PIL import ImageEnhance
    if saturation!=100: img=ImageEnhance.Brightness(img).enhance(saturation/100.0)
    if wcor!=0.0 or lcor!=0.0:
        w,h=img.size
        img=img.resize((int(w*(1+wcor/100)),int(h*(1+lcor/100))),Image.LANCZOS)
    img.save(out,dpi=(dpi,dpi))
    try: os.remove(tmp_out)
    except: pass
    return img

def get_pdf_size(gs,pdf):
    try:
        tmp=tempfile.mktemp(suffix=".pdf"); shutil.copy2(pdf,tmp)
        r=subprocess.run([gs,"-dBATCH","-dNOPAUSE","-dSAFER","-sDEVICE=bbox",tmp],
                         capture_output=True,text=True,env=gs_env())
        os.remove(tmp)
        for line in r.stderr.splitlines():
            if line.startswith("%%BoundingBox:"):
                p=line.split()
                return round((int(p[3])-int(p[1]))/72,2),round((int(p[4])-int(p[2]))/72,2)
    except: pass
    return None,None

def estimate_density(img):
    try:
        import numpy as np; arr=np.array(img)
        return round(100*(arr<128).sum()/arr.size,1)
    except:
        px=list(img.getdata())
        return round(100*sum(1 for p in px if p<128)/len(px),1)

def simulate_halftone(img_patch,lpi=55,angle=22.5,shape="0 round",dpi=300):
    w,h=img_patch.size; result=Image.new("L",(w,h),255); draw=ImageDraw.Draw(result)
    cell_px=dpi/lpi; rad=math.radians(angle); cos_a=math.cos(rad); sin_a=math.sin(rad)
    for iy in range(-1,int(h/cell_px)+3):
        for ix in range(-1,int(w/cell_px)+3):
            cx_r=ix*cell_px; cy_r=iy*cell_px
            cx=cx_r*cos_a-cy_r*sin_a; cy=cx_r*sin_a+cy_r*cos_a
            sx=max(0,min(w-1,int(cx))); sy=max(0,min(h-1,int(cy)))
            gray=img_patch.getpixel((sx,sy)); darkness=(255-gray)/255.0
            max_r=cell_px*0.5*0.95; r=max_r*math.sqrt(darkness)
            if r<0.5: continue
            s=shape.lower()
            if "ellipse" in s: draw.ellipse([cx-r,cy-r*0.6,cx+r,cy+r*0.6],fill=0)
            elif "line" in s: draw.rectangle([cx-r,cy-r*0.3,cx+r,cy+r*0.3],fill=0)
            elif "diamond" in s: draw.polygon([(cx,cy-r),(cx+r,cy),(cx,cy+r),(cx-r,cy)],fill=0)
            else: draw.ellipse([cx-r,cy-r,cx+r,cy+r],fill=0)
    return result

def nest_images(images,sheet_w_in,sheet_h_in,dpi,gap_in=0.25,cols=2,rows=2):
    if not images: return None
    sw=int(sheet_w_in*dpi); sh=int(sheet_h_in*dpi); gp=int(gap_in*dpi)
    cw=(sw-gp*(cols+1))//cols; ch=(sh-gp*(rows+1))//rows
    sheet=Image.new("L",(sw,sh),255)
    for i,img in enumerate(images):
        if i>=cols*rows: break
        col=i%cols; row=i//cols
        thumb=img.copy(); thumb.thumbnail((cw,ch),Image.LANCZOS)
        x=gp+col*(cw+gp)+(cw-thumb.width)//2; y=gp+row*(ch+gp)+(ch-thumb.height)//2
        sheet.paste(thumb,(x,y))
    return sheet

def make_test_strip():
    img=Image.new("L",(800,200),255); draw=ImageDraw.Draw(img)
    for i in range(10):
        gray=int(255*i/9); bw=80
        draw.rectangle([i*bw,0,(i+1)*bw,170],fill=gray)
        draw.rectangle([i*bw,170,(i+1)*bw,200],fill=200)
        draw.text((i*bw+4,176),str(int(100*i/9))+"%",fill=0)
    return img

def list_printers():
    try:
        r=subprocess.run(["lpstat","-a"],capture_output=True,text=True)
        return [l.split()[0] for l in r.stdout.splitlines() if l.strip()] or ["Epson_1430"]
    except: return ["Epson_1430"]

def mk_btn(parent,text,cmd,bg=BG2,fg=FG2,bold=False):
    f=("Helvetica Neue",10,"bold") if bold else ("Helvetica Neue",10)
    return tk.Button(parent,text=text,command=cmd,bg=bg,fg=fg,font=f,
                     relief="flat",padx=10,pady=5,activebackground="#444",
                     activeforeground=FG,cursor="hand2")

def mk_green(parent,text,cmd):
    return tk.Button(parent,text=text,command=cmd,bg=GREEN,fg="#000",
                     font=("Helvetica Neue",10,"bold"),relief="flat",padx=10,pady=5,
                     activebackground=GREEN2,activeforeground="#000",cursor="hand2")

def mk_icon(parent,text,cmd):
    return tk.Button(parent,text=text,command=cmd,bg=BG3,fg=FG2,
                     font=("Helvetica Neue",12),relief="flat",padx=8,pady=4,
                     activebackground=BG2,activeforeground=FG,cursor="hand2",
                     highlightthickness=1,highlightbackground="#333")

def mk_spin(parent,var,frm,to,inc=1,w=7,fmt=None):
    kw=dict(textvariable=var,from_=frm,to=to,increment=inc,width=w,
            bg=BG3,fg=FG,font=("Helvetica Neue",10),buttonbackground=BG2,
            insertbackground=FG,relief="flat",highlightthickness=1,
            highlightbackground="#444",highlightcolor=GREEN)
    if fmt: kw["format"]=fmt
    return tk.Spinbox(parent,**kw)

def mk_combo(parent,var,values,width=16):
    return ttk.Combobox(parent,textvariable=var,values=values,state="readonly",width=width)

def rl(parent,text,r,w=16):
    tk.Label(parent,text=text,bg=BG,fg=FG2,font=("Helvetica Neue",10),
             anchor="e",width=w).grid(row=r,column=0,sticky="e",padx=(0,10),pady=5)

class FloatingPanel(tk.Toplevel):
    def __init__(self,master,title,w,h,x,y):
        super().__init__(master)
        self.title(""); self.configure(bg=BG2)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.resizable(True,True)
        self.protocol("WM_DELETE_WINDOW",self._hide)
        tb=tk.Frame(self,bg=BG3,height=28); tb.pack(fill="x")
        tb.pack_propagate(False)
        tk.Button(tb,text="×",bg=BG3,fg=FG3,font=("Helvetica Neue",12),
                  relief="flat",padx=6,pady=0,
                  activebackground="#555",command=self._hide).pack(side="left",padx=(4,0))
        tk.Label(tb,text=title,bg=BG3,fg=FG2,
                 font=("Helvetica Neue",10)).pack(side="left",padx=8)
        tb.bind("<ButtonPress-1>",self._ds)
        tb.bind("<B1-Motion>",self._dm)
        self._dx=self._dy=0; self.visible=True
    def _ds(self,e): self._dx=e.x; self._dy=e.y
    def _dm(self,e):
        self.geometry(f"+{self.winfo_x()+e.x-self._dx}+{self.winfo_y()+e.y-self._dy}")
    def _hide(self): self.withdraw(); self.visible=False
    def show(self): self.deiconify(); self.visible=True

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.gs=find_gs(); self.cfg=load_config()
        self.pdf=None; self.tiff=None; self.photo=None; self.nav_photo=None
        self.queue=[]; self.queue_tiffs={}
        self._last_img=None; self._ht_photo=None; self._ht_after=None; self._zoom=1.0; self._pan_x=None; self._pan_y=None
        self.title("ScreenRIP"); self.configure(bg=BG)
        self._style(); self._build()
        self.after(200,self._init_panels)
        if not self.gs:
            messagebox.showwarning("Ghostscript Missing","brew install ghostscript")

    def _style(self):
        s=ttk.Style(self); s.theme_use("clam")
        s.configure("TNotebook",background=BG,borderwidth=0,tabmargins=[0,0,0,0])
        s.configure("TNotebook.Tab",background=BG,foreground=FG3,
                    font=("Helvetica Neue",12),padding=[20,8],borderwidth=0)
        s.map("TNotebook.Tab",background=[("selected",BG)],foreground=[("selected",GREEN)])
        for w in ["TFrame","TLabel","TLabelframe"]:
            s.configure(w,background=BG,foreground=FG)
        s.configure("TCheckbutton",background=BG,foreground=FG2,font=("Helvetica Neue",10))
        s.map("TCheckbutton",background=[("active",BG)],
              foreground=[("selected",GREEN),("active",FG)])
        self.option_add("*TCombobox*Listbox.background",BG3)
        self.option_add("*TCombobox*Listbox.foreground",FG)
        self.option_add("*TCombobox*Listbox.selectBackground",DARKGREEN)
        self.option_add("*TCombobox*Listbox.font",("Helvetica Neue",10))

    def _build(self):
        self.columnconfigure(0,weight=0,minsize=460)
        self.columnconfigure(1,weight=1); self.rowconfigure(0,weight=1)

        # LEFT
        left=tk.Frame(self,bg=BG,width=460)
        left.grid(row=0,column=0,sticky="nsew"); left.grid_propagate(False)
        left.columnconfigure(0,weight=1); left.rowconfigure(2,weight=1)

        # Header
        hdr=tk.Frame(left,bg=BG4,height=58); hdr.grid(row=0,column=0,sticky="ew")
        hdr.grid_propagate(False); hdr.columnconfigure(1,weight=1)
        logo=tk.Frame(hdr,bg=BG4); logo.grid(row=0,column=0,padx=14,pady=10)
        tk.Label(logo,text="Screen",font=("Helvetica Neue",20,"bold"),bg=BG4,fg=FG).pack(side="left")
        tk.Label(logo,text="RIP",font=("Helvetica Neue",20,"bold"),bg=BG4,fg=GREEN).pack(side="left")
        tk.Label(logo,text="®",font=("Helvetica Neue",11),bg=BG4,fg=FG2).pack(side="left",pady=(6,0))
        pr_f=tk.Frame(hdr,bg=BG4); pr_f.grid(row=0,column=1,padx=14,pady=14,sticky="e")
        tk.Label(pr_f,text="Select Printer:",bg=BG4,fg=FG2,
                 font=("Helvetica Neue",10)).pack(side="left",padx=(0,6))
        self.prn=tk.StringVar(); pl=list_printers(); self.prn.set(pl[0] if pl else "")
        self.pcb=ttk.Combobox(pr_f,textvariable=self.prn,values=pl,state="readonly",width=20)
        self.pcb.pack(side="left",padx=(0,4))
        tk.Button(pr_f,text="↺",command=self._refresh_printers,bg=BG4,fg=GREEN,
                  font=("Helvetica Neue",13),relief="flat",activebackground="#333").pack(side="left")

        # Preset bar
        pbar=tk.Frame(left,bg=BG2,height=36); pbar.grid(row=1,column=0,sticky="ew")
        pbar.grid_propagate(False); pbar.columnconfigure(1,weight=1)
        tk.Label(pbar,text="Preset:",bg=BG2,fg=FG2,
                 font=("Helvetica Neue",10)).grid(row=0,column=0,padx=(12,6),pady=7)
        self.preset_var=tk.StringVar(value=self.cfg.get("last_preset","1430 Fine Detail"))
        self.preset_cb=ttk.Combobox(pbar,textvariable=self.preset_var,
                                     values=list(self.cfg["presets"].keys()),state="readonly",width=22)
        self.preset_cb.grid(row=0,column=1,sticky="ew",pady=5,padx=(0,4))
        self.preset_cb.bind("<<ComboboxSelected>>",lambda e:self._load_preset())
        for t,c in [("Save",self._save_preset),("Delete",self._delete_preset)]:
            tk.Button(pbar,text=t,command=c,bg=BG2,fg=FG2,font=("Helvetica Neue",9),
                      relief="flat",padx=6,activebackground="#444"
                      ).grid(row=0,column=2 if t=="Save" else 3,pady=5,padx=(0,4))

        # Tabs
        nb=ttk.Notebook(left); nb.grid(row=2,column=0,sticky="nsew")

        # PRINT tab
        pt=tk.Frame(nb,bg=BG,padx=18,pady=12); nb.add(pt,text="Print"); pt.columnconfigure(1,weight=1)
        rl(pt,"Quality:",0); self.quality=tk.StringVar(value="High")
        mk_combo(pt,self.quality,["Draft","Medium","High"],14).grid(row=0,column=1,sticky="w",pady=4)
        rl(pt,"Type:",1); self.media_type=tk.StringVar(value="Sheet")
        mk_combo(pt,self.media_type,["Sheet","Roll"],14).grid(row=1,column=1,sticky="w",pady=4)
        rl(pt,"Paper Type:",2); self.paper_type=tk.StringVar(value="Matte Film")
        mk_combo(pt,self.paper_type,["Matte Film","Transparent Paper","Transparent Paper Light",
            "Plain paper","Single Weight Matte","Double Weight Matte","Enhanced Matte",
            "Archival Matte","Photo Quality Inkjet","Coated Generic","Premium Glossy 170",
            "Premium Semigloss 170","Premium Glossy","Premium Semigloss","Premium Luster",
            "Photo Paper Glossy","Photo Paper Generic","Enhanced Adhesive Synthetic",
            "Enhanced Low Adhesive Synthetic","Heavy Weight Polyester Banner"],22
            ).grid(row=2,column=1,sticky="w",pady=4)
        rl(pt,"Size:",3)
        szf=tk.Frame(pt,bg=BG); szf.grid(row=3,column=1,sticky="w",pady=4)
        self.media_w=tk.DoubleVar(value=13.0)
        mk_spin(szf,self.media_w,1,60,0.5,6,"%.2f").pack(side="left")
        tk.Label(szf,text="x",bg=BG,fg=FG2,font=("Helvetica Neue",10)).pack(side="left",padx=4)
        self.media_h=tk.DoubleVar(value=19.0)
        mk_spin(szf,self.media_h,1,120,0.5,6,"%.2f").pack(side="left")
        tk.Label(szf,text="in",bg=BG,fg=FG2,font=("Helvetica Neue",10)).pack(side="left",padx=(4,8))
        self.sz_cm=tk.Label(szf,text="",bg=BG,fg=FG3,font=("Menlo",8)); self.sz_cm.pack(side="left")
        self.media_w.trace_add("write",lambda *a:self._upd_sz())
        self.media_h.trace_add("write",lambda *a:self._upd_sz())
        rl(pt,"Quick Size:",4)
        qsf=tk.Frame(pt,bg=BG); qsf.grid(row=4,column=1,sticky="w",pady=4)
        for lbl,(w,h) in [("13x19",(13,19)),("17x22",(17,22)),("24 Roll",(24,36))]:
            tk.Button(qsf,text=lbl,command=lambda ww=w,hh=h:self._set_media(ww,hh),
                      bg=BG2,fg=FG2,font=("Helvetica Neue",9),relief="flat",
                      padx=8,pady=3,activebackground="#444").pack(side="left",padx=(0,4))
        rl(pt,"N-UP:",5)
        nf=tk.Frame(pt,bg=BG); nf.grid(row=5,column=1,sticky="w",pady=4)
        self.nup_on=tk.BooleanVar(value=False)
        ttk.Checkbutton(nf,text="Enable",variable=self.nup_on,command=self._tog_nup).pack(side="left",padx=(0,8))
        tk.Label(nf,text="Cols:",bg=BG,fg=FG2,font=("Helvetica Neue",10)).pack(side="left")
        self.nup_cols=tk.IntVar(value=2)
        self.nup_cols_sp=mk_spin(nf,self.nup_cols,1,10,1,3); self.nup_cols_sp.pack(side="left",padx=(4,8))
        tk.Label(nf,text="Rows:",bg=BG,fg=FG2,font=("Helvetica Neue",10)).pack(side="left")
        self.nup_rows=tk.IntVar(value=2)
        self.nup_rows_sp=mk_spin(nf,self.nup_rows,1,10,1,3); self.nup_rows_sp.pack(side="left",padx=4)
        rl(pt,"Enhance:",6); self.enhance=tk.StringVar(value="None")
        mk_combo(pt,self.enhance,["None","Edge Sharpen","Smooth"],14).grid(row=6,column=1,sticky="w",pady=4)
        rl(pt,"Shirt Color:",7)
        scf=tk.Frame(pt,bg=BG); scf.grid(row=7,column=1,sticky="w",pady=4)
        self.shirt_mode=tk.StringVar(value="White")
        self.btn_white=tk.Button(scf,text="White Shirt",bg="#d0d0d0",fg="#000",
                                  font=("Helvetica Neue",10,"bold"),relief="flat",padx=10,pady=4,
                                  command=lambda:self._set_shirt("White")); self.btn_white.pack(side="left",padx=(0,8))
        self.btn_black=tk.Button(scf,text="Black Shirt",bg=BG2,fg=FG2,
                                  font=("Helvetica Neue",10),relief="flat",padx=10,pady=4,
                                  command=lambda:self._set_shirt("Black")); self.btn_black.pack(side="left")
        tk.Button(pt,text="Mesh Count / LPI Guide...",command=self._show_lpi_guide,bg=BG,fg="#4fc3f7",font=("Helvetica Neue",10),relief="flat",cursor="hand2",activebackground=BG,activeforeground="#00e5ff").grid(row=8,column=0,columnspan=2,sticky="w",pady=(16,0))

        # INKS tab
        it=tk.Frame(nb,bg=BG,padx=18,pady=12); nb.add(it,text="Inks"); it.columnconfigure(1,weight=1)
        rl(it,"Screen Mode:",0,20)
        self.mode=tk.StringVar(value="Multi Black")
        mk_combo(it,self.mode,["Single Black","Multi Black","Single Black (pre-halftoned)",
                               "Multi Black (pre-halftoned)","Single Black (no halftone)",
                               "All Black (no halftone)"],22).grid(row=0,column=1,sticky="w",pady=4)
        rl(it,"Ink Saturation:",1,20)
        satf=tk.Frame(it,bg=BG); satf.grid(row=1,column=1,sticky="w",pady=4)
        self.sat=tk.IntVar(value=200)
        mk_spin(satf,self.sat,50,200,10,5).pack(side="left")
        tk.Label(satf,text="% default",bg=BG,fg=FG2,font=("Helvetica Neue",10)).pack(side="left",padx=(6,0))
        tk.Frame(it,bg=BORDER,height=1).grid(row=2,column=0,columnspan=2,sticky="ew",pady=8)
        rl(it,"Multi Black Cart:",3,20)
        cbf=tk.Frame(it,bg=BG); cbf.grid(row=3,column=1,sticky="w",pady=4)
        self.ck=tk.BooleanVar(value=True); self.cc=tk.BooleanVar(value=True)
        self.cm_=tk.BooleanVar(value=True); self.cy=tk.BooleanVar(value=True)
        for lbl,var in [("Black",self.ck),("Cyan",self.cc),("Magenta",self.cm_),("Yellow",self.cy)]:
            ttk.Checkbutton(cbf,text=lbl,variable=var).pack(side="left",padx=(0,12))
        rl(it,"Drop Size:",4,20); self.drop=tk.StringVar(value="1 small")
        mk_combo(it,self.drop,["1 small","2 medium","3 large"],14).grid(row=4,column=1,sticky="w",pady=4)

        # HALFTONE tab
        ht=tk.Frame(nb,bg=BG,padx=18,pady=12); nb.add(ht,text="Halftone"); ht.columnconfigure(1,weight=1)
        rl(ht,"Screen LPI:",0)
        self.lpi=tk.IntVar(value=55)
        mk_spin(ht,self.lpi,30,100,1,5).grid(row=0,column=1,sticky="w",pady=4)
        self.lpi.trace_add("write",lambda *a:self._sched_ht())
        rl(ht,"Screen Angle:",1)
        self.ang=tk.DoubleVar(value=22.5)
        mk_spin(ht,self.ang,0,90,7.5,6,"%.1f").grid(row=1,column=1,sticky="w",pady=4)
        self.ang.trace_add("write",lambda *a:self._sched_ht())
        rl(ht,"Screen Shape:",2); self.shp=tk.StringVar(value="0 round")
        mk_combo(ht,self.shp,["0 round","1 ellipse","2 line","3 diamond"],14).grid(row=2,column=1,sticky="w",pady=4)
        self.shp.trace_add("write",lambda *a:self._sched_ht())
        rl(ht,"Resolution:",3); self.dpi=tk.IntVar(value=1440)
        mk_combo(ht,self.dpi,[720,1440,2880],10).grid(row=3,column=1,sticky="w",pady=4)
        tk.Frame(ht,bg=BORDER,height=1).grid(row=4,column=0,columnspan=2,sticky="ew",pady=8)
        tk.Label(ht,text="Tip: Use the zoom buttons on the preview to inspect halftone dots.",bg=BG,fg=FG3,font=("Menlo",9),wraplength=300,justify="left").grid(row=5,column=0,columnspan=2,sticky="w",pady=(0,4))
        tk.Label(ht,text="Zoom to 3x-4x after processing to see individual dots.",bg=BG,fg=FG3,font=("Menlo",9),wraplength=300,justify="left").grid(row=6,column=0,columnspan=2,sticky="w",pady=(0,12))
        tk.Button(ht,text="Process Halftone Preview",
                  command=self._process_halftone,
                  bg=GREEN,fg="#000",font=("Helvetica Neue",11,"bold"),
                  relief="flat",padx=12,pady=8,
                  activebackground=GREEN2,activeforeground="#000",cursor="hand2"
                  ).grid(row=7,column=0,columnspan=2,sticky="ew",pady=(0,6))
        self.ht_status=tk.Label(ht,text="Render a PDF first, then process halftone.",
                                 bg=BG,fg=FG3,font=("Menlo",9),wraplength=300,justify="left")
        self.ht_status.grid(row=8,column=0,columnspan=2,sticky="w")

        # LAYOUT tab
        ly=tk.Frame(nb,bg=BG,padx=18,pady=12); nb.add(ly,text="Layout"); ly.columnconfigure(1,weight=1)
        rl(ly,"Width Correction:",0,20)
        wcf=tk.Frame(ly,bg=BG); wcf.grid(row=0,column=1,sticky="w",pady=4)
        self.wcor=tk.DoubleVar(value=0.0)
        mk_spin(wcf,self.wcor,-10,10,0.5,6,"%.1f").pack(side="left")
        tk.Label(wcf,text="% default",bg=BG,fg=FG2,font=("Helvetica Neue",10)).pack(side="left",padx=(6,0))
        rl(ly,"Length Correction:",1,20)
        lcf=tk.Frame(ly,bg=BG); lcf.grid(row=1,column=1,sticky="w",pady=4)
        self.lcor=tk.DoubleVar(value=0.0)
        mk_spin(lcf,self.lcor,-10,10,0.5,6,"%.1f").pack(side="left")
        tk.Label(lcf,text="% default",bg=BG,fg=FG2,font=("Helvetica Neue",10)).pack(side="left",padx=(6,0))
        tk.Frame(ly,bg=BORDER,height=1).grid(row=2,column=0,columnspan=2,sticky="ew",pady=8)
        rl(ly,"Auto-Save To:",3,20)
        svf=tk.Frame(ly,bg=BG); svf.grid(row=3,column=1,sticky="ew",pady=4); svf.columnconfigure(0,weight=1)
        self.save_lbl=tk.Label(svf,text=self.cfg.get("save_folder") or "Not set",
                                bg=BG3,fg=FG2,font=("Menlo",9),anchor="w",padx=6,pady=3)
        self.save_lbl.grid(row=0,column=0,sticky="ew",pady=(0,4))
        mk_btn(svf,"Choose Folder...",self._pick_save_folder).grid(row=1,column=0,sticky="w")
        rl(ly,"N-UP Gap (in):",4,20); self.nup_gap=tk.DoubleVar(value=0.25)
        mk_spin(ly,self.nup_gap,0,2,0.125,6,"%.3f").grid(row=4,column=1,sticky="w",pady=4)

        # SUPPORT tab
        sp=tk.Frame(nb,bg=BG,padx=18,pady=12); nb.add(sp,text="Support"); sp.columnconfigure(0,weight=1)
        tk.Label(sp,text="ScreenRIP Support",bg=BG,fg=GREEN,
                 font=("Helvetica Neue",13,"bold")).grid(row=0,column=0,sticky="w",pady=(0,10))

        # Version info
        ver_f=tk.Frame(sp,bg=BG2); ver_f.grid(row=1,column=0,sticky="ew",pady=(0,8))
        ver_f.columnconfigure(1,weight=1)
        tk.Label(ver_f,text="Version:",bg=BG2,fg=FG2,font=("Helvetica Neue",10),
                 width=12,anchor="w").grid(row=0,column=0,padx=(10,4),pady=6)
        self.ver_lbl=tk.Label(ver_f,text=SCREENRIP_VERSION,bg=BG2,fg=GREEN,
                               font=("Helvetica Neue",10,"bold"))
        self.ver_lbl.grid(row=0,column=1,sticky="w")
        mk_btn(ver_f,"Check for Updates",self._check_updates).grid(row=0,column=2,padx=(0,10),pady=4)

        tk.Frame(sp,bg=BORDER,height=1).grid(row=2,column=0,sticky="ew",pady=(0,10))

        # Paste update section
        tk.Label(sp,text="Paste Update from Claude",bg=BG,fg=GREEN,
                 font=("Helvetica Neue",11,"bold")).grid(row=3,column=0,sticky="w",pady=(0,4))
        tk.Label(sp,text="Copy updated code from Claude, paste below, click Apply Update.",
                 bg=BG,fg=FG3,font=("Menlo",9)).grid(row=4,column=0,sticky="w",pady=(0,6))
        self.paste_txt=tk.Text(sp,height=8,bg=BG3,fg=FG,font=("Menlo",9),
                                insertbackground=FG,relief="flat",
                                highlightthickness=1,highlightbackground="#444",
                                highlightcolor=GREEN)
        self.paste_txt.grid(row=5,column=0,sticky="ew",pady=(0,6))
        paste_btns=tk.Frame(sp,bg=BG); paste_btns.grid(row=6,column=0,sticky="w",pady=(0,10))
        mk_btn(paste_btns,"Apply Update",self._apply_paste_update).pack(side="left",padx=(0,8))
        mk_btn(paste_btns,"Clear",lambda:self.paste_txt.delete("1.0","end")).pack(side="left")

        tk.Frame(sp,bg=BORDER,height=1).grid(row=7,column=0,sticky="ew",pady=(0,10))

        # Install info
        tk.Label(sp,text="Dependencies",bg=BG,fg=GREEN,
                 font=("Helvetica Neue",11,"bold")).grid(row=8,column=0,sticky="w",pady=(0,6))
        for i,(t,b) in enumerate([
            ("Ghostscript","brew install ghostscript"),
            ("Pillow","pip3 install pillow --break-system-packages"),
            ("NumPy","pip3 install numpy --break-system-packages"),
            ("Launch","OBJC_DISABLE_INITIALIZE_FORK_SAFETY=YES python3 ~/Desktop/ScreenRIP/epson_rip.py")]):
            f=tk.Frame(sp,bg=BG2); f.grid(row=9+i,column=0,sticky="ew",pady=(0,3))
            f.columnconfigure(1,weight=1)
            tk.Label(f,text=t,bg=BG2,fg=GREEN,font=("Helvetica Neue",10,"bold"),
                     width=12,anchor="w").grid(row=0,column=0,padx=(8,4),pady=5)
            tk.Label(f,text=b,bg=BG2,fg=FG2,font=("Menlo",9),anchor="w").grid(row=0,column=1,padx=(0,8),sticky="ew")
        tk.Label(sp,text="Version " + SCREENRIP_VERSION + "  •  Built with Ghostscript + Pillow",
                 bg=BG,fg=FG3,font=("Menlo",8)).grid(row=14,column=0,sticky="w",pady=(12,0))

        # Bottom bar
        bot=tk.Frame(left,bg=BG4); bot.grid(row=3,column=0,sticky="ew")
        bot.columnconfigure(0,weight=1)
        self.stv=tk.StringVar(value="Ready")
        tk.Label(bot,textvariable=self.stv,bg=BG4,fg=FG3,
                 font=("Menlo",9),anchor="w").grid(row=0,column=0,sticky="ew",padx=12,pady=(6,0))
        act=tk.Frame(bot,bg=BG4); act.grid(row=1,column=0,sticky="ew",padx=10,pady=8)
        act.columnconfigure(0,weight=1)
        la=tk.Frame(act,bg=BG4); la.grid(row=0,column=0,columnspan=2,sticky="ew",pady=(0,4))
        la.columnconfigure(0,weight=1)
        row1=tk.Frame(la,bg=BG4); row1.grid(row=0,column=0,sticky="w")
        mk_btn(row1,"Revert",self._load_preset).pack(side="left",padx=(0,6))
        mk_green(row1,"Apply",self._save_preset).pack(side="left",padx=(0,6))
        mk_btn(row1,"Test Strip",self._test_strip).pack(side="left",padx=(0,6))
        mk_btn(row1,"Print All Queue",self._print_all).pack(side="left",padx=(0,6))
        self.pbtn=mk_green(row1,"Print Film Positive",self._print)
        self.pbtn.pack(side="left")
        self.pbtn.config(state="disabled",bg="#1a3a1a",fg="#444")
        self.prog=ttk.Progressbar(left,mode="indeterminate")
        self.prog.grid(row=4,column=0,sticky="ew")

        # RIGHT
        right=tk.Frame(self,bg=BG3); right.grid(row=0,column=1,sticky="nsew")
        right.columnconfigure(0,weight=1); right.rowconfigure(1,weight=1)
        pvhdr=tk.Frame(right,bg=BG4,height=30); pvhdr.grid(row=0,column=0,sticky="ew")
        pvhdr.grid_propagate(False); pvhdr.columnconfigure(1,weight=1)
        tk.Label(pvhdr,text="Proof Positive® Preview",bg=BG4,fg=FG2,
                 font=("Helvetica Neue",10)).grid(row=0,column=0,padx=12,pady=6,sticky="w")
        self.info_var=tk.StringVar(value="")
        tk.Label(pvhdr,textvariable=self.info_var,bg=BG4,fg=GREEN,
                 font=("Menlo",9)).grid(row=0,column=1,sticky="e",padx=12)
        self.cv=tk.Canvas(right,bg=BG3,highlightthickness=0)
        self.cv.grid(row=1,column=0,sticky="nsew")
        self.cv.create_text(400,400,
            text="Proof Positive® Preview\n\nRender a PDF to see the film positive",
            fill="#2a2a2a",font=("Helvetica Neue",14),justify="center")
        # Toolbar
        tb=tk.Frame(right,bg=BG4,height=42); tb.grid(row=2,column=0,sticky="ew")
        tb.grid_propagate(False); tb.columnconfigure(1,weight=1)
        lico=tk.Frame(tb,bg=BG4); lico.grid(row=0,column=0,padx=8,pady=6)
        mk_icon(lico,"✥",self._tog_pan).pack(side="left",padx=(0,2))
        mk_icon(lico,"⊕",self._zoom_in).pack(side="left",padx=(0,2))
        mk_icon(lico,"⊖",self._zoom_out).pack(side="left",padx=(0,6))
        tk.Label(lico,text="Zoom:",bg=BG4,fg=FG3,font=("Menlo",9)).pack(side="left")
        self._zoom_var=tk.DoubleVar(value=1.0)
        self._zoom_slider=tk.Scale(lico,from_=1.0,to=8.0,resolution=0.1,
            orient="horizontal",variable=self._zoom_var,
            command=self._on_zoom_slider,
            bg=BG4,fg=FG,troughcolor=BG3,activebackground=GREEN,
            highlightthickness=0,length=180,showvalue=True,
            sliderrelief="flat")
        self._zoom_slider.pack(side="left",padx=(4,6))
        mk_btn(lico,"Fit",self._zoom_fit).pack(side="left",padx=(0,2))
        mk_btn(lico,"Max",self._zoom_max).pack(side="left",padx=(0,2))
        tk.Label(nf,text="Density:",bg=BG4,fg=FG3,font=("Menlo",9)).pack(side="left")
        self.density_var=tk.StringVar(value="--")
        tk.Label(lico,textvariable=self.density_var,bg=BG4,fg=GREEN,
                 font=("Menlo",9,"bold")).pack(side="left",padx=(4,0))
        rico=tk.Frame(tb,bg=BG4); rico.grid(row=0,column=2,padx=8,pady=6,sticky="e")
        mk_btn(rico,"Reprocess",self._render).pack(side="left",padx=(0,4))
        mk_btn(rico,"Clear",self._clear_preview).pack(side="left",padx=(0,4))
        mk_btn(rico,"Save TIFF",self._save_tiff).pack(side="left",padx=(0,4))
        mk_green(rico,"Print",self._print).pack(side="left")
        self.geometry("1200x900")
        self._load_preset(); self._upd_sz(); self._tog_nup()

    def _init_panels(self):
        sidebar=tk.Frame(self,bg=BG2,width=260)
        sidebar.grid(row=0,column=2,sticky="nsew")
        sidebar.grid_propagate(False)
        sidebar.columnconfigure(0,weight=1)
        self.columnconfigure(2,weight=0,minsize=260)
        tk.Label(sidebar,text="NAVIGATION",bg=BG2,fg=FG3,
                 font=("Helvetica Neue",8,"bold")).pack(pady=(10,4))
        self.nav_cv=tk.Canvas(sidebar,bg="#111",highlightthickness=1,
                               highlightbackground="#333",width=230,height=200)
        self.nav_cv.pack(padx=10)
        self.nav_cv.create_text(85,92,text="no preview",fill="#222",font=("Menlo",8))
        self.size_var=tk.StringVar(value="--")
        tk.Label(sidebar,textvariable=self.size_var,bg=BG2,fg=FG2,
                 font=("Menlo",9),justify="center").pack(pady=(4,0))
        tk.Frame(sidebar,bg=BORDER,height=1).pack(fill="x",padx=10,pady=8)
        tk.Label(sidebar,text="INK DENSITY",bg=BG2,fg=FG3,
                 font=("Helvetica Neue",8,"bold")).pack()
        self.side_density=tk.StringVar(value="--")
        tk.Label(sidebar,textvariable=self.side_density,bg=BG2,fg=GREEN,
                 font=("Helvetica Neue",18,"bold")).pack(pady=(2,6))
        tk.Frame(sidebar,bg=BORDER,height=1).pack(fill="x",padx=10,pady=(0,8))
        tk.Label(sidebar,text="PRINT QUEUE",bg=BG2,fg=FG3,
                 font=("Helvetica Neue",8,"bold")).pack()
        qf=tk.Frame(sidebar,bg=BG2); qf.pack(fill="x",padx=8,pady=(4,0))
        qf.columnconfigure(0,weight=1)
        self.queue_lb=tk.Listbox(qf,bg=BG3,fg=FG,font=("Menlo",8),height=5,
                                  selectmode="single",highlightthickness=1,
                                  highlightbackground="#444",highlightcolor=GREEN,
                                  bd=0,selectbackground=DARKGREEN,selectforeground=GREEN2)
        self.queue_lb.grid(row=0,column=0,sticky="ew",pady=(0,4))
        self.queue_lb.bind("<<ListboxSelect>>",self._on_queue_select)
        qb=tk.Frame(qf,bg=BG2); qb.grid(row=1,column=0,sticky="ew")
        for i in range(4): qb.columnconfigure(i,weight=1)
        for i,(t,c) in enumerate([("Add",self._pick),("Rem",self._remove_queue),
                                   ("Up",self._queue_up),("Dn",self._queue_down)]):
            tk.Button(qb,text=t,command=c,bg=BG3,fg=FG2,font=("Helvetica Neue",8),
                      relief="flat",padx=2,pady=2,activebackground="#444"
                      ).grid(row=0,column=i,sticky="ew",padx=1)
        tk.Frame(sidebar,bg=BORDER,height=1).pack(fill="x",padx=10,pady=8)
        tk.Label(sidebar,text="RECENT FILES",bg=BG2,fg=FG3,
                 font=("Helvetica Neue",8,"bold")).pack()
        rf=tk.Frame(sidebar,bg=BG2); rf.pack(fill="x",padx=8,pady=(4,0))
        rf.columnconfigure(0,weight=1)
        self.recent_lb=tk.Listbox(rf,bg=BG3,fg=FG2,font=("Menlo",8),height=4,
                                   highlightthickness=1,highlightbackground="#444",
                                   highlightcolor=GREEN,bd=0,
                                   selectbackground=DARKGREEN,selectforeground=GREEN2)
        self.recent_lb.grid(row=0,column=0,sticky="ew",pady=(0,4))
        self.recent_lb.bind("<Double-Button-1>",self._open_recent)
        tk.Button(rf,text="Clear Recent",command=self._clear_recent,
                  bg=BG3,fg=FG2,font=("Helvetica Neue",8),relief="flat",
                  padx=4,pady=2,activebackground="#444").grid(row=1,column=0,sticky="w")
        self._refresh_recent()

    def _show_lpi_guide(self):
        w = tk.Toplevel(self)
        w.title("Mesh Count / LPI Guide")
        w.configure(bg=BG)
        w.resizable(False, False)
        tk.Label(w, text="Mesh Count / LPI Reference Guide", bg=BG, fg=GREEN,
                 font=("Helvetica Neue", 13, "bold")).pack(pady=(16, 8), padx=20)
        guide = [
            ("Mesh", "Ink Type", "Manual Press LPI", "Best For"),
            ("86",  "Heavy",     "25-35",  "Glitter, puff, thick plastisol"),
            ("110", "Heavy",     "30-40",  "Discharge, waterbase thick"),
            ("130", "Medium",    "35-45",  "General plastisol, spot color"),
            ("156", "Medium",    "40-50",  "Standard detail work"),
            ("160", "Medium",    "45-50",  "Good all-purpose mesh"),
            ("180", "Fine",      "45-55",  "Finer detail, thinner inks"),
            ("200", "Fine",      "50-55",  "Halftone simulated process"),
            ("230", "Fine",      "50-60",  "High detail, thin inks"),
            ("255", "Fine",      "55-65",  "Very fine detail"),
            ("305", "Extra Fine","55-65",  "Process color, photographic"),
            ("355", "Extra Fine","60-70",  "Extremely fine detail"),
        ]
        f = tk.Frame(w, bg=BG); f.pack(padx=16, pady=(0, 8), fill="x")
        widths = [8, 12, 16, 26]
        for ri, row in enumerate(guide):
            rbg = BG if ri % 2 == 0 else BG2
            for ci, (cell, cw) in enumerate(zip(row, widths)):
                fg = GREEN if ri == 0 else (GREEN if ci == 2 else FG2)
                font = ("Helvetica Neue", 10, "bold") if ri == 0 else ("Menlo", 9)
                tk.Label(f, text=cell, bg=rbg, fg=fg, font=font,
                         width=cw, anchor="w", padx=6, pady=4
                         ).grid(row=ri, column=ci, sticky="ew")
        tk.Label(w, text="Manual press rule of thumb: LPI = Mesh count / 4 to 4.5",
                 bg=BG, fg=FG3, font=("Menlo", 9)).pack(pady=(4, 2))
        tk.Label(w, text="Lower LPI = easier dot hold on manual press. Go higher only with thin ink + good tension.",
                 bg=BG, fg=FG3, font=("Menlo", 9)).pack(pady=(0, 12))
        tk.Button(w, text="Close", command=w.destroy, bg=BG2, fg=FG,
                  font=("Helvetica Neue", 10), relief="flat", padx=16, pady=6,
                  activebackground="#444").pack(pady=(0, 16))

    def _process_halftone(self):
        if not self._last_img:
            messagebox.showinfo("No Image","Render a PDF first.")
            return
        self.ht_status.config(text="Processing halftone...",fg=GREEN)
        self.prog.start(12)
        lpi=self.lpi.get(); angle=self.ang.get(); shape=self.shp.get()
        def work():
            try:
                import math
                from PIL import Image,ImageDraw
                img=self._last_img.copy()
                iw,ih=img.size
                dpi=min(self.dpi.get(),360)
                result=Image.new("L",(iw,ih),255)
                draw=ImageDraw.Draw(result)
                cell_px=dpi/lpi
                rad=math.radians(angle)
                cos_a=math.cos(rad); sin_a=math.sin(rad)
                num_x=int(iw/cell_px)+3; num_y=int(ih/cell_px)+3
                for iy in range(-1,num_y+1):
                    for ix in range(-1,num_x+1):
                        cx_r=ix*cell_px; cy_r=iy*cell_px
                        cx=cx_r*cos_a-cy_r*sin_a
                        cy=cx_r*sin_a+cy_r*cos_a
                        sx=max(0,min(iw-1,int(cx)))
                        sy=max(0,min(ih-1,int(cy)))
                        gray=img.getpixel((sx,sy))
                        darkness=(255-gray)/255.0
                        r=cell_px*0.5*0.9*math.sqrt(darkness)
                        if r<0.5: continue
                        s=shape.lower()
                        if "ellipse" in s:
                            draw.ellipse([cx-r,cy-r*0.6,cx+r,cy+r*0.6],fill=0)
                        elif "line" in s:
                            draw.rectangle([cx-r,cy-r*0.3,cx+r,cy+r*0.3],fill=0)
                        elif "diamond" in s:
                            draw.polygon([(cx,cy-r),(cx+r,cy),(cx,cy+r),(cx-r,cy)],fill=0)
                        else:
                            draw.ellipse([cx-r,cy-r,cx+r,cy+r],fill=0)
                self.after(0,self._show_halftone_result,result)
            except Exception as e:
                self.after(0,self.ht_status.config,{"text":"Error: "+str(e)[:60],"fg":"#ff5252"})
            finally:
                self.after(0,self.prog.stop)
        import threading
        threading.Thread(target=work,daemon=True).start()

    def _show_halftone_result(self,img):
        self._last_img=img
        self._zoom=1.0; self._pan_x=None; self._pan_y=None
        self._display_image(img)
        self.ht_status.config(
            text="Halftone processed! Zoom in with + to see dots.",fg=GREEN)
        self.pbtn.config(state="normal",bg=GREEN,fg="#000")

    def _check_updates(self):
        self.stv.set("Checking for updates...")
        import threading
        def work():
            try:
                import urllib.request, json
                url = "https://raw.githubusercontent.com/screenrip/screenrip/main/version.json"
                try:
                    with urllib.request.urlopen(url, timeout=5) as r:
                        data = json.loads(r.read())
                        latest = data.get("version","unknown")
                except:
                    latest = None
                if latest and latest != SCREENRIP_VERSION:
                    self.after(0, messagebox.showinfo, "Update Available",
                        "Update available: v" + latest + "\nGet it from Claude and paste in Support tab.")
                    self.after(0, self.stv.set, "Update available: v" + latest)
                else:
                    self.after(0, messagebox.showinfo, "Up to Date",
                        "ScreenRIP v" + SCREENRIP_VERSION + " is up to date!")
                    self.after(0, self.stv.set, "Up to date!")
            except:
                self.after(0, self.stv.set, "Could not check for updates")
        threading.Thread(target=work, daemon=True).start()

    def _apply_paste_update(self):
        code = self.paste_txt.get("1.0", "end").strip()
        if not code:
            messagebox.showinfo("Empty", "Paste the updated code from Claude first.")
            return
        if "def " not in code or "import " not in code:
            messagebox.showerror("Invalid", "Does not look like valid Python code.")
            return
        if not messagebox.askyesno("Apply Update", "Replace current code and restart?"):
            return
        try:
            import ast as am
            am.parse(code)
        except SyntaxError as e:
            messagebox.showerror("Syntax Error", str(e))
            return
        import os, shutil, sys, subprocess
        path = os.path.abspath(__file__)
        shutil.copy2(path, path + ".backup")
        open(path, "w").write(code)
        messagebox.showinfo("Updated!", "ScreenRIP updated! Restarting...")
        env = os.environ.copy()
        env["OBJC_DISABLE_INITIALIZE_FORK_SAFETY"] = "YES"
        subprocess.Popen([sys.executable, path], env=env)
        self.destroy()

    def _upd_sz(self):
        try: self.sz_cm.config(text=f"({self.media_w.get()*2.54:.1f}x{self.media_h.get()*2.54:.1f}cm)")
        except: pass
    def _set_media(self,w,h): self.media_w.set(w); self.media_h.set(h)
    def _tog_nup(self):
        on=self.nup_on.get(); st="normal" if on else "disabled"
        self.nup_cols_sp.config(state=st); self.nup_rows_sp.config(state=st)
    def _set_shirt(self,mode):
        self.shirt_mode.set(mode)
        if mode=="White":
            self.btn_white.config(bg="#d0d0d0",fg="#000",font=("Helvetica Neue",10,"bold"))
            self.btn_black.config(bg=BG2,fg=FG2,font=("Helvetica Neue",10))
        else:
            self.btn_white.config(bg=BG2,fg=FG2,font=("Helvetica Neue",10))
            self.btn_black.config(bg="#111",fg=GREEN,font=("Helvetica Neue",10,"bold"))
    def _zoom_in(self):
        steps=[1.0,1.1,1.25,1.5,1.75,2.0,2.5,3.0,4.0,5.0,6.0,8.0]
        for s in steps:
            if s > round(self._zoom,2)+0.05:
                self._zoom=s; break
        self._update_zoom_label(); self._sync_zoom_slider()
        self._redisplay()

    def _zoom_out(self):
        steps=[8.0,6.0,5.0,4.0,3.0,2.5,2.0,1.75,1.5,1.25,1.1,1.0]
        for s in steps:
            if s < round(self._zoom,2)-0.05:
                self._zoom=s; break
        else:
            self._zoom=1.0; self._pan_x=None; self._pan_y=None
        self._update_zoom_label(); self._sync_zoom_slider()
        self._redisplay()

    def _update_zoom_label(self):
        pct=int(self._zoom*100)
        try: self.zoom_lbl.config(text=f"Zoom: {pct}%")
        except: pass

    def _sync_zoom_slider(self):
        try: self._zoom_var.set(round(self._zoom,1))
        except: pass

    def _on_zoom_slider(self,val):
        self._zoom=float(val)
        if self._zoom<=1.0: self._pan_x=None; self._pan_y=None
        self._redisplay()

    def _zoom_fit(self):
        self._zoom=1.0
        self._pan_x=None; self._pan_y=None
        self._redisplay()
    def _zoom_max(self):
        self._zoom=6.0
        self._redisplay()
    def _pan_start(self,e):
        if self._zoom>1.0:
            self._ps_x=e.x; self._ps_y=e.y
            self._pi_x=getattr(self,"_pan_x",None)
            self._pi_y=getattr(self,"_pan_y",None)
    def _pan_move(self,e):
        if self._zoom<=1.0 or not hasattr(self,"_ps_x"): return
        if not self._last_img: return
        iw,ih=self._last_img.size
        cw=self.cv.winfo_width() or 800
        ch=self.cv.winfo_height() or 860
        rw=int(cw/self._zoom); rh=int(ch/self._zoom)
        dx=int((self._ps_x-e.x)/self._zoom)
        dy=int((self._ps_y-e.y)/self._zoom)
        bx=self._pi_x or iw//2
        by=self._pi_y or ih//2
        self._pan_x=max(rw//2,min(iw-rw//2,bx+dx))
        self._pan_y=max(rh//2,min(ih-rh//2,by+dy))
        self._redisplay()

    def _tog_pan(self): pass
    def _redisplay(self):
        if self._last_img: self._display_image(self._last_img)
    def _display_image(self,img):
        self._last_img=img
        cw=self.cv.winfo_width() or 800
        ch=self.cv.winfo_height() or 860
        iw,ih=img.size
        z=self._zoom

        if z<=1.0:
            # Fit: scale whole image to canvas
            scale=min(cw/iw, ch/ih)
            nw=int(iw*scale); nh=int(ih*scale)
            disp=img.resize((nw,nh),Image.LANCZOS)
        else:
            # Zoom: crop a region at actual resolution so dots are visible
            px=getattr(self,"_pan_x",iw//2) or iw//2
            py=getattr(self,"_pan_y",ih//2) or ih//2
            rw=int(cw/z); rh=int(ch/z)
            x0=max(0,min(iw-rw,px-rw//2))
            y0=max(0,min(ih-rh,py-rh//2))
            crop=img.crop((x0,y0,x0+rw,y0+rh))
            disp=crop.resize((cw,ch),Image.LANCZOS)

        self.photo=ImageTk.PhotoImage(disp)
        self.cv.delete("all")
        self.cv.create_image(cw//2,ch//2,image=self.photo,anchor="center")

        # Nav thumbnail
        nav=img.copy(); nav.thumbnail((230,200),Image.LANCZOS)
        self.nav_photo=ImageTk.PhotoImage(nav)
        self.nav_cv.delete("all")
        self.nav_cv.create_image(115,100,image=self.nav_photo,anchor="center")

        # Enable pan when zoomed
        self.cv.bind("<ButtonPress-1>",self._pan_start)
        self.cv.bind("<B1-Motion>",self._pan_move)

    def _update_density(self,d):
        self.density_var.set(f"{d}%")
    def _sched_ht(self): pass
    def _refresh_ht(self):
        self._ht_after=None
        if not self._last_img:
            self.ht_canvas.delete("all")
            self.ht_canvas.create_text(145,145,text="Render a PDF first",fill="#444",font=("Menlo",9))
            return
        lpi=self.lpi.get(); angle=self.ang.get(); shape=self.shp.get(); dpi=min(self.dpi.get(),300)
        def work():
            try:
                img=self._last_img; iw,ih=img.size; ppx=int(3.0*dpi)
                cx,cy=iw//2,ih//2
                x0=max(0,cx-ppx//2); y0=max(0,cy-ppx//2)
                patch=img.crop((x0,y0,min(iw,x0+ppx),min(ih,y0+ppx))).resize((ppx,ppx),Image.LANCZOS)
                ht=simulate_halftone(patch,lpi=lpi,angle=angle,shape=shape,dpi=dpi)
                ht.thumbnail((290,290),Image.LANCZOS)
                self.after(0,self._show_ht,ht)
            except: pass
        threading.Thread(target=work,daemon=True).start()
    def _show_ht(self,img):
        pass
    def _pick(self):
        paths=filedialog.askopenfilenames(title="Choose PDF(s)",filetypes=[("PDF","*.pdf"),("All","*.*")])
        for p in paths:
            if p not in self.queue:
                self.queue.append(p); self.queue_lb.insert("end",Path(p).name)
        if paths: self._load_pdf(paths[0])
    def _load_pdf(self,path):
        self.pdf=path; self.tiff=None
        self.pbtn.config(state="disabled",bg="#1a3a1a",fg="#444")
        self.stv.set("Loaded: "+Path(path).name); self._add_recent(path)
        if self.gs:
            def gi():
                w,h=get_pdf_size(self.gs,path)
                if w and h:
                    self.after(0,self.size_var.set,f"{w}\" x {h}\"\n{w*2.54:.1f}x{h*2.54:.1f}cm")
                    self.after(0,self.info_var.set,f"{w}\" x {h}\"")
            threading.Thread(target=gi,daemon=True).start()
        self.cv.delete("all")
        self.cv.create_text(400,400,text="Click Reprocess to render",fill="#2a2a2a",font=("Helvetica Neue",13))
    def _on_queue_select(self,e):
        sel=self.queue_lb.curselection()
        if sel and sel[0]<len(self.queue): self._load_pdf(self.queue[sel[0]])
    def _remove_queue(self):
        sel=self.queue_lb.curselection()
        if sel:
            i=sel[0]; p=self.queue.pop(i); self.queue_lb.delete(i)
            if p in self.queue_tiffs: del self.queue_tiffs[p]
    def _queue_up(self):
        sel=self.queue_lb.curselection()
        if sel and sel[0]>0:
            i=sel[0]; self.queue[i-1],self.queue[i]=self.queue[i],self.queue[i-1]
            t=self.queue_lb.get(i-1); self.queue_lb.delete(i-1)
            self.queue_lb.insert(i,t); self.queue_lb.selection_set(i-1)
    def _queue_down(self):
        sel=self.queue_lb.curselection()
        if sel and sel[0]<len(self.queue)-1:
            i=sel[0]; self.queue[i],self.queue[i+1]=self.queue[i+1],self.queue[i]
            t=self.queue_lb.get(i+1); self.queue_lb.delete(i+1)
            self.queue_lb.insert(i,t); self.queue_lb.selection_set(i+1)
    def _add_recent(self,path):
        r=self.cfg.get("recent_files",[])
        if path in r: r.remove(path)
        r.insert(0,path); self.cfg["recent_files"]=r[:10]
        save_config(self.cfg); self._refresh_recent()
    def _refresh_recent(self):
        if not hasattr(self,"recent_lb"): return
        self.recent_lb.delete(0,"end")
        for p in self.cfg.get("recent_files",[]): self.recent_lb.insert("end",Path(p).name)
    def _open_recent(self,e):
        sel=self.recent_lb.curselection()
        if sel:
            files=self.cfg.get("recent_files",[]); idx=sel[0]
            if idx<len(files):
                path=files[idx]
                if Path(path).exists():
                    if path not in self.queue:
                        self.queue.append(path); self.queue_lb.insert("end",Path(path).name)
                    self._load_pdf(path)
                else: messagebox.showwarning("Not Found",f"File not found:\n{path}")
    def _clear_recent(self):
        self.cfg["recent_files"]=[]; save_config(self.cfg); self._refresh_recent()
    def _get_settings(self):
        return {"lpi":self.lpi.get(),"angle":self.ang.get(),"shape":self.shp.get(),
                "dpi":self.dpi.get(),"sat":self.sat.get(),"mode":self.mode.get(),
                "drop":self.drop.get(),"wcor":self.wcor.get(),"lcor":self.lcor.get(),
                "quality":self.quality.get(),"media_type":self.media_type.get(),
                "media_w":self.media_w.get(),"media_h":self.media_h.get(),
                "paper_type":self.paper_type.get(),"enhance":self.enhance.get(),
                "nup_cols":self.nup_cols.get(),"nup_rows":self.nup_rows.get(),
                "nup_gap":self.nup_gap.get(),"shirt":self.shirt_mode.get()}
    def _apply_settings(self,s):
        self.lpi.set(s.get("lpi",55)); self.ang.set(s.get("angle",22.5))
        self.shp.set(s.get("shape","0 round")); self.dpi.set(s.get("dpi",1440))
        self.sat.set(s.get("sat",200)); self.mode.set(s.get("mode","Multi Black"))
        self.drop.set(s.get("drop","1 small")); self.wcor.set(s.get("wcor",0.0))
        self.lcor.set(s.get("lcor",0.0)); self.quality.set(s.get("quality","High"))
        self.media_type.set(s.get("media_type","Sheet"))
        self.media_w.set(s.get("media_w",13.0)); self.media_h.set(s.get("media_h",19.0))
        self.paper_type.set(s.get("paper_type","Matte Film"))
        self.enhance.set(s.get("enhance","None"))
        self.nup_cols.set(s.get("nup_cols",2)); self.nup_rows.set(s.get("nup_rows",2))
        self.nup_gap.set(s.get("nup_gap",0.25)); self._set_shirt(s.get("shirt","White"))
    def _load_preset(self):
        name=self.preset_var.get()
        if name in self.cfg["presets"]:
            self._apply_settings(self.cfg["presets"][name])
            self.cfg["last_preset"]=name; save_config(self.cfg)
            self.stv.set("Preset loaded: "+name)
    def _save_preset(self):
        name=simpledialog.askstring("Save Preset","Preset name:",
                                    initialvalue=self.preset_var.get(),parent=self)
        if not name: return
        self.cfg["presets"][name]=self._get_settings()
        self.cfg["last_preset"]=name; save_config(self.cfg)
        self.preset_var.set(name)
        self.preset_cb["values"]=list(self.cfg["presets"].keys())
        self.stv.set("Preset saved: "+name)
    def _delete_preset(self):
        name=self.preset_var.get()
        if name in self.cfg["presets"]:
            if messagebox.askyesno("Delete",f"Delete preset '{name}'?"):
                del self.cfg["presets"][name]; save_config(self.cfg)
                keys=list(self.cfg["presets"].keys())
                self.preset_cb["values"]=keys
                if keys: self.preset_var.set(keys[0])
    def _pick_save_folder(self):
        f=filedialog.askdirectory(title="Choose folder for auto-saved TIFFs")
        if f:
            self.cfg["save_folder"]=f; save_config(self.cfg)
            self.save_lbl.config(text=f,fg=FG)
    def _auto_save(self,img,pdf_path):
        folder=self.cfg.get("save_folder","")
        if not folder: return
        img.save(str(Path(folder)/(Path(pdf_path).stem+"_rip.tiff")))
    def _save_tiff(self):
        if not self._last_img: messagebox.showinfo("Nothing","Render first."); return
        path=filedialog.asksaveasfilename(defaultextension=".tiff",filetypes=[("TIFF","*.tiff")])
        if path: self._last_img.save(path); self.stv.set("Saved: "+Path(path).name)
    def _refresh_printers(self):
        pl=list_printers(); self.pcb["values"]=pl
        if pl: self.prn.set(pl[0])
    def _clear_preview(self):
        self.cv.delete("all")
        self.cv.create_text(400,400,text="Preview cleared",fill="#2a2a2a",font=("Helvetica Neue",13))
        self._last_img=None; self.tiff=None
        self.pbtn.config(state="disabled",bg="#1a3a1a",fg="#444")
        self.density_var.set("--")
        self.nav_cv.delete("all")
        self.nav_cv.create_text(70,82,text="no preview",fill="#222",font=("Menlo",8))
    def _render(self):
        if not self.pdf: messagebox.showinfo("No File","Please add a PDF first."); return
        if not self.gs: messagebox.showerror("No Ghostscript","brew install ghostscript"); return
        self.prog.start(12); self.stv.set("Rendering...")
        def work():
            try:
                out=tempfile.mktemp(suffix=".tiff")
                img=render_pdf(self.gs,self.pdf,out,dpi=self.dpi.get(),lpi=self.lpi.get(),
                               angle=self.ang.get(),shape=self.shp.get(),
                               saturation=self.sat.get(),mode=self.mode.get(),
                               wcor=self.wcor.get(),lcor=self.lcor.get(),
                               quality=self.quality.get(),
                               cb=lambda m:self.after(0,self.stv.set,m))
                if self.shirt_mode.get()=="Black": img=ImageOps.invert(img)
                self.tiff=out; self.queue_tiffs[self.pdf]=out
                density=estimate_density(img); self._auto_save(img,self.pdf)
                self.after(0,self._show,img,density)
            except Exception as e:
                self.after(0,messagebox.showerror,"Render Error",str(e))
                self.after(0,self.stv.set,"Render failed.")
            finally: self.after(0,self.prog.stop)
        threading.Thread(target=work,daemon=True).start()
    def _show(self,img,density=None):
        self._display_image(img)
        self.pbtn.config(state="normal",bg=GREEN,fg="#000")
        if density is not None:
            self._update_density(density)
            self.stv.set(f"Proof Positive ready  |  Ink: {density}%")
        else: self.stv.set("Proof Positive ready")
        self.after(200,self._refresh_ht)
    def _test_strip(self):
        p=self.prn.get()
        if not p: messagebox.showinfo("No Printer","Select a printer."); return
        def work():
            try:
                img=make_test_strip(); tmp=tempfile.mktemp(suffix=".tiff"); img.save(tmp)
                r=subprocess.run(["lpr","-P",p,"-o","ColorModel=Gray",
                                   "-o",f"Resolution={self.dpi.get()}dpi",tmp],
                                  capture_output=True,text=True)
                os.remove(tmp)
                if r.returncode!=0: raise RuntimeError(r.stderr)
                self.after(0,self.stv.set,"Test strip sent!")
                self.after(0,messagebox.showinfo,"Done","Test strip sent!")
            except Exception as e: self.after(0,messagebox.showerror,"Error",str(e))
        threading.Thread(target=work,daemon=True).start()
    def _do_print(self,tiff_path,printer):
        sf=[]
        if self.quality.get()=="High": sf=["-o","print-quality=5","-o","OutputMode=highest"]
        elif self.quality.get()=="Medium": sf=["-o","print-quality=4"]
        cmd=(["lpr","-P",printer,"-o","ColorModel=Gray",
              "-o",f"Resolution={self.dpi.get()}dpi"]+sf+[tiff_path])
        r=subprocess.run(cmd,capture_output=True,text=True)
        if r.returncode!=0: raise RuntimeError(r.stderr)
    def _print(self):
        if not self.tiff or not os.path.exists(self.tiff):
            messagebox.showinfo("Not Rendered","Render first."); return
        p=self.prn.get()
        if not p: messagebox.showinfo("No Printer","Select a printer."); return
        if not messagebox.askyesno("Confirm",f"Print to:\n{p}\n\nContinue?"): return
        self.prog.start(12); self.stv.set("Sending to printer...")
        def work():
            try:
                self._do_print(self.tiff,p)
                self.after(0,self.stv.set,"Sent!"); self.after(0,messagebox.showinfo,"Done","Print job sent!")
            except Exception as e: self.after(0,messagebox.showerror,"Print Error",str(e))
            finally: self.after(0,self.prog.stop)
        threading.Thread(target=work,daemon=True).start()
    def _print_all(self):
        if not self.queue: messagebox.showinfo("Empty","Add PDFs first."); return
        p=self.prn.get()
        if not p: messagebox.showinfo("No Printer","Select a printer."); return
        n=len(self.queue)
        if not messagebox.askyesno("Print All",f"Print {n} file(s) to:\n{p}\n\nContinue?"): return
        self.prog.start(12)
        def work():
            for i,pdf in enumerate(self.queue):
                self.after(0,self.stv.set,f"Processing {i+1}/{n}: {Path(pdf).name}")
                try:
                    out=tempfile.mktemp(suffix=".tiff")
                    img=render_pdf(self.gs,pdf,out,dpi=self.dpi.get(),lpi=self.lpi.get(),
                                   angle=self.ang.get(),shape=self.shp.get(),
                                   saturation=self.sat.get(),mode=self.mode.get(),
                                   wcor=self.wcor.get(),lcor=self.lcor.get(),
                                   quality=self.quality.get())
                    self._auto_save(img,pdf); self._do_print(out,p); time.sleep(1)
                except Exception as e:
                    self.after(0,messagebox.showerror,"Error",f"{Path(pdf).name}:\n{e}")
            self.after(0,self.prog.stop)
            self.after(0,self.stv.set,f"Done - {n} file(s) sent!")
            self.after(0,messagebox.showinfo,"All Done",f"All {n} file(s) printed!")
        threading.Thread(target=work,daemon=True).start()

App().mainloop()
