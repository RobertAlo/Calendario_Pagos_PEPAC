# -*- coding: utf-8 -*-
"""
Calendario FEAGA/FEADER – v3.4.0
- Calendario: tk.Button + <ButtonRelease-1> + after(0, ...) -> cada clic refresca sí o sí.
- Botón "Limpiar panel" que además resetea la selección del calendario (por si quieres empezar de cero).
- Terminología FEAGA en toda la UI.
- Persistencia SQLite (thread-safe).
"""

import calendar
import re
import sqlite3
import threading
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox

DB_FILE = "pagos_pepac.sqlite3"

def iso(d: date) -> str: return d.strftime("%Y-%m-%d")
def fmt_dmy(d: date) -> str: return d.strftime("%d/%m/%Y")

def parse_ddmmyyyy(s: str) -> date | None:
    s = s.strip()
    if not s: return None
    try: return datetime.strptime(s, "%d/%m/%Y").date()
    except Exception: return None

def daterange(d1: date, d2: date):
    cur = d1
    while cur <= d2:
        yield cur
        cur += timedelta(days=1)

# -------------------- BD (hilo-seguro) --------------------
class PaymentsDB:
    def __init__(self, path: str = DB_FILE):
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self._lock = threading.Lock()
        self._ensure_schema()

    def _ensure_schema(self):
        with self._lock:
            c=self.conn.cursor()
            c.execute("""
                CREATE TABLE IF NOT EXISTS pagos(
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    fecha TEXT NOT NULL,          -- YYYY-MM-DD
                    tipo TEXT NOT NULL,
                    fondo TEXT NOT NULL,
                    detalle TEXT NOT NULL,
                    fuente TEXT DEFAULT '',
                    origen TEXT NOT NULL CHECK (origen IN ('manual','web','heuristica')),
                    created_at TEXT DEFAULT (datetime('now'))
                )
            """)
            c.execute("""CREATE UNIQUE INDEX IF NOT EXISTS ux_pagos
                         ON pagos(fecha,tipo,fondo,detalle,origen)""")
            self.conn.commit()

    def add(self, d: date, tipo: str, fondo: str, detalle: str, fuente: str="", origen: str="manual"):
        with self._lock:
            self.conn.execute(
                "INSERT OR IGNORE INTO pagos(fecha,tipo,fondo,detalle,fuente,origen) VALUES (?,?,?,?,?,?)",
                (iso(d), tipo.strip(), fondo.strip(), detalle.strip(), fuente.strip(), origen))
            self.conn.commit()

    def add_range(self, d1: date, d2: date, **kwargs):
        with self._lock:
            cur=self.conn.cursor()
            for d in daterange(d1,d2):
                cur.execute("INSERT OR IGNORE INTO pagos(fecha,tipo,fondo,detalle,fuente,origen) VALUES (?,?,?,?,?,?)",
                            (iso(d), kwargs.get("tipo",""), kwargs.get("fondo",""),
                             kwargs.get("detalle",""), kwargs.get("fuente",""), kwargs.get("origen","manual")))
            self.conn.commit()

    def delete_day(self, d: date, origen: str | None = None):
        with self._lock:
            if origen:
                self.conn.execute("DELETE FROM pagos WHERE fecha=? AND origen=?", (iso(d), origen))
            else:
                self.conn.execute("DELETE FROM pagos WHERE fecha=?", (iso(d),))
            self.conn.commit()

    def delete_all(self, include_heuristic=True):
        with self._lock:
            if include_heuristic: self.conn.execute("DELETE FROM pagos")
            else: self.conn.execute("DELETE FROM pagos WHERE origen IN ('manual','web')")
            self.conn.commit()

    def get_day(self, d: date, origins: set[str] | None = None) -> list[dict]:
        with self._lock:
            q="SELECT * FROM pagos WHERE fecha=?"; args=[iso(d)]
            if origins:
                q+=f" AND origen IN ({','.join('?'*len(origins))})"; args+=list(origins)
            q+=" ORDER BY origen DESC, tipo"
            rows=self.conn.execute(q,args).fetchall()
            return [dict(r) for r in rows]

    def get_month(self, y:int, m:int, origins: set[str] | None = None) -> list[dict]:
        d1=date(y,m,1); d2=date(y,m,calendar.monthrange(y,m)[1])
        return self.get_range(d1,d2,origins)

    def get_range(self, d1: date, d2: date, origins: set[str] | None = None) -> list[dict]:
        with self._lock:
            q="SELECT * FROM pagos WHERE date(fecha) BETWEEN date(?) AND date(?)"; args=[iso(d1),iso(d2)]
            if origins:
                q+=f" AND origen IN ({','.join('?'*len(origins))})"; args+=list(origins)
            q+=" ORDER BY fecha, origen DESC, tipo"
            rows=self.conn.execute(q,args).fetchall()
            return [dict(r) for r in rows]

    def has_day(self, d: date) -> bool:
        with self._lock:
            r=self.conn.execute("SELECT 1 FROM pagos WHERE fecha=? LIMIT 1",(iso(d),)).fetchone()
        return bool(r)

# -------------------- Referencia FEAGA --------------------
class FeagaRef:
    @staticmethod
    def campaign_year_for(d: date) -> int:
        return d.year if d.month>=10 else d.year-1

    @staticmethod
    def windows_for_campaign(cy:int):
        ant1=date(cy,10,16); ant2=date(cy,11,30)
        sal1=date(cy,12,1);  sal2=date(cy+1,6,30)
        return (("Anticipo ayudas directas", ant1, ant2),
                ("Saldo ayudas directas",     sal1, sal2))

    @staticmethod
    def seed(db: PaymentsDB, cy:int):
        for tipo, a, b in FeagaRef.windows_for_campaign(cy):
            db.add_range(a,b, tipo=tipo, fondo="FEAGA",
                         detalle=f"Ventana general de {('anticipos' if 'Anticipo' in tipo else 'saldos')}. Campaña {cy}.",
                         fuente="Referencia FEAGA", origen="heuristica")

    @staticmethod
    def day_in_any_window(d: date) -> list[dict]:
        cy=FeagaRef.campaign_year_for(d)
        out=[]
        for tipo,a,b in FeagaRef.windows_for_campaign(cy):
            if a<=d<=b:
                out.append({"fecha": iso(d), "tipo": tipo, "fondo":"FEAGA",
                            "detalle": f"Ventana general ({a.strftime('%d/%m')}–{b.strftime('%d/%m')}). Campaña {cy}.",
                            "fuente":"Referencia FEAGA", "origen":"info"})
        # Recordatorio FEADER
        out.append({"fecha": iso(d), "tipo": "Referencia: FEADER (desarrollo rural)",
                    "fondo":"FEADER", "detalle":"Pagos según resoluciones/convocatorias autonómicas.",
                    "fuente":"Referencia", "origen":"info"})
        return out

    @staticmethod
    def month_generic_for_day(d: date) -> list[dict]:
        cy=FeagaRef.campaign_year_for(d)
        y, m = d.year, d.month
        m1=date(y,m,1); m2=date(y,m,calendar.monthrange(y,m)[1])
        rows=[]
        for tipo,a,b in FeagaRef.windows_for_campaign(cy):
            start=max(a,m1); end=min(b,m2)
            if start<=end and not (a<=d<=b):
                detalle=(f"Este día ({fmt_dmy(d)}) está fuera; en {m:02d}/{y} la ventana es "
                         f"{start.strftime('%d/%m')}–{end.strftime('%d/%m')} (campaña {cy}).")
                rows.append({"fecha": iso(d), "tipo": f"Referencia mes: {tipo}",
                             "fondo":"FEAGA", "detalle": detalle,
                             "fuente":"Referencia FEAGA", "origen":"info"})
        if not rows:
            rows.append({"fecha": iso(d), "tipo":"Referencia mes: Sin pagos FEAGA generales",
                         "fondo":"—", "detalle":"Fuera de ventanas de anticipo/saldo en este mes.",
                         "fuente":"Referencia FEAGA", "origen":"info"})
        # Añadimos recordatorio FEADER
        rows.append({"fecha": iso(d), "tipo": "Referencia: FEADER (desarrollo rural)",
                     "fondo":"FEADER", "detalle":"Pagos según resoluciones/convocatorias autonómicas.",
                     "fuente":"Referencia", "origen":"info"})
        return rows

# -------------------- Scraper web (opcional) --------------------
class FegaWebScraper:
    NOTE_URLS = [
        "https://www.fega.gob.es/sites/default/files/files/document/Nota_web_Ecorregimenes_Ca_2024_ANTICIPO.pdf",
        "https://www.fega.gob.es/sites/default/files/files/document/Nota_Web_AAS_Ca_2024_ANTICIPO.pdf",
        "https://www.fga.gob.es/sites/default/files/files/document/241115_NOTA_WEB_EERR_PRIMER_SALDO_Ca_2024_def.pdf".replace("fga","fega")
    ]
    def __init__(self):
        try:
            import requests  # noqa
        except Exception:
            self._ok=False
        else:
            self._ok=True
    def available(self): return self._ok

    def fetch_into_db(self, db: PaymentsDB, year_hint:int|None=None):
        if not self._ok: raise RuntimeError("Falta 'requests' (pip install requests).")
        import requests
        for url in self.NOTE_URLS:
            try:
                r=requests.get(url,timeout=15)
                content=r.content.decode("latin-1",errors="ignore") if r.status_code==200 else url
            except Exception:
                content=url
            m=re.search(r"(20\d{2})", content)
            y=int(m.group(1)) if m else (year_hint or date.today().year)
            ant1=date(y,10,16); ant2=date(y,11,30)
            sal1=date(y,12,1);  sal2=date(y+1,6,30)
            db.add_range(ant1, ant2, tipo="Anticipo ayudas directas", fondo="FEAGA",
                         detalle="Ventana general de anticipos (nota FEAGA).", fuente=url, origen="web")
            db.add_range(sal1, sal2, tipo="Saldo ayudas directas", fondo="FEAGA",
                         detalle="Ventana general de saldos (nota FEAGA).", fuente=url, origen="web")

# -------------------- StatusBar pastel --------------------
class StatusBar(tk.Frame):
    COLORS = {"info":("#e6f0ff","#093a76"), "ok":("#eaf7ee","#1b5e20"),
              "warn":("#fff7e6","#8a4b00"), "error":("#fdecea","#8a1c13")}
    def __init__(self, master):
        super().__init__(master, bg="#e6f0ff", highlightbackground="#cddcfb", highlightthickness=1)
        self._label = tk.Label(self, text="", bg="#e6f0ff", fg="#093a76", anchor="w")
        self._label.pack(fill="x", padx=10, pady=4)
        self.hide()
    def show(self, kind: str, text: str):
        bg, fg = self.COLORS.get(kind, self.COLORS["info"])
        self.configure(bg=bg, highlightbackground=bg)
        self._label.configure(text=text, bg=bg, fg=fg)
        self.pack(fill="x", padx=10, pady=(6,0)); self.update_idletasks()
    def hide(self): self.pack_forget()

# -------------------- Scroll contenedor --------------------
class VerticalScrolledFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.canvas=tk.Canvas(self,highlightthickness=0)
        self.vsb=ttk.Scrollbar(self,orient="vertical",command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.inner=ttk.Frame(self.canvas)
        self._win=self.canvas.create_window((0,0),window=self.inner,anchor="nw")
        self.canvas.grid(row=0,column=0,sticky="nsew"); self.vsb.grid(row=0,column=1,sticky="ns")
        self.rowconfigure(0,weight=1); self.columnconfigure(0,weight=1)
        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)
        self.canvas.bind_all("<Button-4>", self._on_wheel)
        self.canvas.bind_all("<Button-5>", self._on_wheel)
    def _on_inner_configure(self,_):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.itemconfigure(self._win, width=self.canvas.winfo_width())
    def _on_canvas_configure(self,e):
        self.canvas.itemconfigure(self._win, width=e.width)
    def _on_wheel(self, ev):
        try:
            if ev.num==4: self.canvas.yview_scroll(-3,"units")
            elif ev.num==5: self.canvas.yview_scroll(3,"units")
            else: self.canvas.yview_scroll(int(-ev.delta/40),"units")
        except Exception: pass

# -------------------- Calendario con botones (triple seguridad) --------------------
class YearCalendarFrame(ttk.Frame):
    MESES_ES=["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
    DIAS=["L","M","X","J","V","S","D"]

    def __init__(self, master, year=None, on_day_click=None, on_day_context=None, has_events_predicate=None):
        super().__init__(master)
        calendar.setfirstweekday(calendar.MONDAY)
        self.current_year=year or date.today().year
        self.on_day_click=on_day_click
        self.on_day_context=on_day_context
        self.has_events_predicate=has_events_predicate or (lambda y,m,d: False)
        self._btns={}; self._selected=None
        self._styles(); self._controls()
        self.scroll=VerticalScrolledFrame(self); self.scroll.grid(row=1,column=0,sticky="nsew")
        self.rowconfigure(1,weight=1); self.columnconfigure(0,weight=1)
        self._grid_year()

    def _styles(self):
        st=ttk.Style(self)
        st.configure("CalHead.TLabel", padding=(0,0), font=("Segoe UI",9))
        st.configure("CalHeadWE.TLabel", padding=(0,0), font=("Segoe UI",9), foreground="#b00000")

    def _controls(self):
        top=ttk.Frame(self); top.grid(row=0,column=0,sticky="ew",pady=(0,4)); top.columnconfigure(6,weight=1)
        ttk.Button(top,text="<<",width=3,command=self._prev).grid(row=0,column=0,padx=(0,4))
        self.year_var=tk.IntVar(value=self.current_year)
        spin=ttk.Spinbox(top,from_=1900,to=2100,width=6,textvariable=self.year_var,justify="center",
                         command=lambda:self._set(self.year_var.get()))
        spin.grid(row=0,column=1)
        spin.bind("<Return>", lambda e:self._set(self.year_var.get()))
        spin.bind("<FocusOut>", lambda e:self._set(self.year_var.get()))
        ttk.Button(top,text=">>",width=3,command=self._next).grid(row=0,column=2,padx=(4,8))
        ttk.Button(top,text="Hoy",command=self._go_today).grid(row=0,column=3)
        self.info=ttk.Label(top,text="(clic=día · doble clic=mes · clic derecho=menú)")
        self.info.grid(row=0,column=5,sticky="e")

    def refresh(self): self._grid_year()

    def clear_selection(self):
        """Limpia selección visual y lógica (para el botón 'Limpiar panel')."""
        if self._selected and self._selected in self._btns:
            self._restore_style(self._selected)
        self._selected=None

    def _grid_year(self):
        for ch in self.scroll.inner.winfo_children(): ch.destroy()
        self._btns.clear()
        grid=ttk.Frame(self.scroll.inner); grid.grid(row=0,column=0,sticky="nsew")
        for r in range(4): grid.rowconfigure(r,weight=1,uniform="m")
        for c in range(3): grid.columnconfigure(c,weight=1,uniform="m")
        for month in range(1,13):
            r,c=(month-1)//3,(month-1)%3
            self._month(grid,self.current_year,month).grid(row=r,column=c,padx=4,pady=4,sticky="nsew")
        self.scroll.inner.rowconfigure(0,weight=1); self.scroll.inner.columnconfigure(0,weight=1)
        if self._selected and self._selected.year==self.current_year:
            self._highlight(self._selected)
        else:
            self._highlight(date(self.current_year,1,1))

    def _month(self,parent,year,month):
        f=ttk.Frame(parent,borderwidth=1,relief="solid",padding=(4,3,4,4))
        ttk.Label(f,text=self.MESES_ES[month-1],font=("Segoe UI",9,"bold")).grid(row=0,column=0,columnspan=7,sticky="ew",pady=(0,2))
        for i,d in enumerate(self.DIAS):
            ttk.Label(f,text=d,style=("CalHeadWE.TLabel" if i in (5,6) else "CalHead.TLabel")).grid(row=1,column=i,padx=1,sticky="nsew")

        weeks=calendar.monthcalendar(year,month)
        font_norm=("Segoe UI",9); font_bold=("Segoe UI",9,"bold")
        color_weekend="#b00000"; color_event="#0a66c2"

        for r,week in enumerate(weeks,start=2):
            for c in range(7):
                dd=week[c]
                if dd==0:
                    ttk.Label(f,text="").grid(row=r,column=c,padx=1,pady=1,sticky="nsew"); continue
                has=self.has_events_predicate(year,month,dd)
                is_weekend=(c in (5,6))
                fg = color_weekend if is_weekend else (color_event if has else "#000")
                bg = "#eef5ff" if has else "#f6f6f6"
                b=tk.Button(f,text=str(dd),width=3,bd=1,relief="raised",
                            bg=bg,activebackground="#dde8ff",fg=fg,highlightthickness=0,
                            font=(font_bold if has else font_norm))
                b.grid(row=r,column=c,padx=1,pady=1,sticky="nsew")
                dt=date(year,month,dd)
                # 1) command, 2) ButtonRelease-1, 3) forzamos via after(0,...)
                b.configure(command=lambda dti=dt: self._click_day(dti))
                b.bind("<ButtonRelease-1>", lambda e,dti=dt: self._click_day(dti))
                b.bind("<Double-Button-1>", lambda e,y=year,m=month: self._dbl_month(y,m))
                b.bind("<Button-3>", lambda e,dti=dt: self._ctx(e,dti))
                b._meta=(bg,fg,has)
                self._btns[dt]=b

        for c in range(7): f.columnconfigure(c,weight=1,uniform="d")
        for rr in range(2,2+len(weeks)): f.grid_rowconfigure(rr,weight=1,minsize=26)
        return f

    def _restore_style(self, dt: date):
        b=self._btns.get(dt)
        if not b: return
        bg,fg,has=b._meta
        b.configure(relief="raised", bg=bg, fg=fg, font=("Segoe UI",9,"bold" if has else "normal"))

    def _highlight(self, dt: date):
        if self._selected and self._selected in self._btns:
            self._restore_style(self._selected)
        self._selected=dt
        b=self._btns.get(dt)
        if b: b.configure(relief="sunken", bg="#dbeafe")

    def _click_day(self, dt: date):
        self._highlight(dt)
        if callable(self.on_day_click):
            # tercera defensa: ejecutar en la cola de eventos, imposible perderse
            self.after(0, lambda d=dt: self.on_day_click(d))

    def _ctx(self, e, dt: date):
        if callable(self.on_day_context): self.on_day_context(e, dt)

    def _dbl_month(self, y,m):
        if callable(self.on_day_click): self.after(0, lambda: self.on_day_click(date(y,m,1), force_month=True))

    def go_to_date(self, dt: date):
        self._set(dt.year); self._highlight(dt)
        if callable(self.on_day_click): self.after(0, lambda d=dt: self.on_day_click(d))

    def get_selected_date(self)->date|None: return self._selected

    def _set(self,year):
        try:year=int(year)
        except:year=self.current_year
        year=max(1900,min(2100,year)); self.current_year=year
        try:self.year_var.set(year)
        except Exception: pass
        self._grid_year()
    def _prev(self): self._set(self.current_year-1)
    def _next(self): self._set(self.current_year+1)
    def _go_today(self): self.go_to_date(date.today())

# -------------------- Panel de pagos --------------------
class PaymentsInfoFrame(ttk.Frame):
    ORIGIN_LABEL = {"manual":"Manual","web":"Web","heuristica":"Referencia FEAGA","info":"Información"}
    def __init__(self, master):
        super().__init__(master, padding=(8,6,8,6), style="Pane.TFrame")
        head=tk.Frame(self,bg="#e6f0ff",highlightbackground="#cddcfb",highlightthickness=1)
        tk.Label(head,text="Pagos",bg="#e6f0ff",fg="#093a76",
                 font=("Segoe UI",10,"bold")).pack(side="left",padx=6,pady=3)
        self.title_lbl=tk.Label(head,text="—",bg="#e6f0ff",fg="#093a76")
        self.title_lbl.pack(side="right",padx=6)
        head.grid(row=0,column=0,columnspan=2,sticky="ew",pady=(0,6))

        legend=ttk.Frame(self); legend.grid(row=1,column=0,columnspan=2,sticky="w")
        self._legend(legend,"#e8fff3","Manual")
        self._legend(legend,"#fff7e6","Web")
        self._legend(legend,"#eef5ff","Referencia FEAGA")
        self._legend(legend,"#e7f3fe","Información")

        cols=("fecha","tipo","fondo","detalle","origen","fuente")
        self.tree=ttk.Treeview(self,columns=cols,show="headings",selectmode="none",style="Colored.Treeview")
        headers={"fecha":(92,"w"),"tipo":(200,"w"),"fondo":(80,"w"),
                 "detalle":(560,"w"),"origen":(130,"center"),"fuente":(220,"w")}
        for c,(w,anc) in headers.items():
            self.tree.heading(c,text=c.capitalize()); self.tree.column(c,width=w,anchor=anc,stretch=True)
        self.tree.grid(row=2,column=0,sticky="nsew")
        ysb=ttk.Scrollbar(self,orient="vertical",command=self.tree.yview)
        xsb=ttk.Scrollbar(self,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set,xscrollcommand=xsb.set)
        ysb.grid(row=2,column=1,sticky="ns"); xsb.grid(row=3,column=0,sticky="ew")
        self.grid_rowconfigure(2,weight=1); self.grid_columnconfigure(0,weight=1)
        self.tree.tag_configure("manual", background="#e8fff3")
        self.tree.tag_configure("web", background="#fff7e6")
        self.tree.tag_configure("heuristica", background="#eef5ff")
        self.tree.tag_configure("info", background="#e7f3fe")
        ttk.Label(self,text="Nota: FEAGA 16/10–30/11 (anticipos), 01/12–30/06 (saldos).",
                  foreground="#666").grid(row=4,column=0,columnspan=2,sticky="w",pady=(6,0))

    @staticmethod
    def _legend(parent, color, text):
        box=tk.Frame(parent,bg=color,width=16,height=12,highlightthickness=1,highlightbackground="#aaa")
        box.pack(side="left",padx=(6,4),pady=(0,4))
        ttk.Label(parent,text=text).pack(side="left",padx=(0,10))

    def clear(self,title="—"):
        self.title_lbl.config(text=title)
        self.tree.delete(*self.tree.get_children())
        self.update_idletasks()

    def show_rows(self,title: str, rows: list[dict]):
        self.title_lbl.config(text=f"{title} · {len(rows)} elemento{'s' if len(rows)!=1 else ''}")
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            label = self.ORIGIN_LABEL.get(r.get("origen",""), r.get("origen",""))
            vals=(r.get("fecha",""), r["tipo"], r["fondo"], r["detalle"], label, r.get("fuente",""))
            self.tree.insert("", "end", values=vals, tags=(r.get("origen",""),))
        self.update_idletasks()

# -------------------- App principal --------------------
class App(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.db=PaymentsDB(DB_FILE)
        self._setup_styles()
        self.show_manual=tk.BooleanVar(value=True)
        self.show_web=tk.BooleanVar(value=True)
        self.show_heur=tk.BooleanVar(value=True)
        FeagaRef.seed(self.db, FeagaRef.campaign_year_for(date.today()))
        self._build_ui()
        self._current_dt=None
        self._show_day(date.today())

    def _setup_styles(self):
        st=ttk.Style()
        try:
            import platform
            st.theme_use("winnative" if platform.system()=="Windows" else "clam")
        except Exception: pass
        st.configure("Pane.TFrame", background="#fafafa")
        st.configure("Colored.Treeview", rowheight=22)

    @staticmethod
    def _color_btn(parent, text, bg, command):
        btn=tk.Button(parent,text=text,bg=bg,fg="white",activebackground=bg,
                      relief="raised",bd=1,highlightthickness=0,command=command)
        btn.configure(font=("Segoe UI",9,"bold"), padx=10, pady=4, cursor="hand2")
        return btn

    def _build_ui(self):
        self.pack(fill="both",expand=True)

        title=tk.Frame(self,bg="#f0f7ff")
        tk.Label(title,text="Calendario FEAGA / FEADER – v3.4.0",
                 bg="#f0f7ff",fg="#053e7b",font=("Segoe UI",13,"bold")).pack(side="left",padx=10,pady=8)
        title.pack(fill="x")

        self.console = StatusBar(self)
        self.console.show("info", "Usa el calendario o los botones de arriba para consultar pagos.")

        top=tk.Frame(self,bg="#eef5ff",highlightbackground="#cddcfb",highlightthickness=1)
        self._color_btn(top,"Pagos de hoy","#2e7d32", lambda: self._show_day(date.today())).pack(side="left", padx=(8,4), pady=6)
        self._color_btn(top,"Mostrar pagos del mes","#0d6efd", self._show_month_of_selected).pack(side="left", padx=4, pady=6)
        ttk.Button(top,text="Buscar rango…",command=self._show_range_dialog).pack(side="left", padx=6, pady=6)
        ttk.Button(top,text="Añadir pago…",command=self._add_payment_dialog).pack(side="left", padx=(12,4), pady=6)
        ttk.Button(top,text="Borrar pagos del día…",command=self._delete_selected_day_dialog).pack(side="left", padx=4, pady=6)
        ttk.Button(top,text="Limpiar panel",command=self._clear_view).pack(side="left", padx=(12,4), pady=6)

        ttk.Label(top,text=" | Mostrar: ").pack(side="left", padx=(12,2))
        ttk.Checkbutton(top,text="Manual",variable=self.show_manual,command=self._refresh_current_view).pack(side="left")
        ttk.Checkbutton(top,text="Web",variable=self.show_web,command=self._refresh_current_view).pack(side="left")
        ttk.Checkbutton(top,text="Referencia FEAGA",variable=self.show_heur,command=self._refresh_current_view).pack(side="left")

        ttk.Button(top,text="Regenerar referencia FEAGA",command=self._regen_heuristics).pack(side="right", padx=6, pady=6)
        ttk.Button(top,text="Vaciar TODO",command=lambda:self._clear_db(include_heur=True)).pack(side="right", padx=(0,8), pady=6)
        ttk.Button(top,text="Vaciar manual+web",command=lambda:self._clear_db(include_heur=False)).pack(side="right", padx=6, pady=6)
        tk.Button(top,text="¿?", width=3, command=self._help_clear).pack(side="right", padx=(0,6), pady=6)
        top.pack(fill="x")

        self.tabs=ttk.Notebook(self); self.tabs.pack(fill="both",expand=True,padx=10,pady=10)

        tab_cal=ttk.Frame(self.tabs); self.tabs.add(tab_cal,text="Calendario")
        v_split=ttk.Panedwindow(tab_cal,orient="vertical"); v_split.pack(fill="both",expand=True)

        self.pay_frame=PaymentsInfoFrame(v_split)
        cal_holder=ttk.Frame(v_split)
        v_split.add(cal_holder,weight=3); v_split.add(self.pay_frame,weight=2)

        tools=ttk.Frame(cal_holder); tools.pack(fill="x",pady=(0,4))
        self._color_btn(tools,"Actualizar pagos (web)","#ff8f00", self._update_from_web).pack(side="left", padx=2, pady=2)
        ttk.Label(tools,text="(opcional; los registros quedan como origen 'web')",foreground="#666").pack(side="left", padx=8)

        def has_ev(y,m,d):
            dt=date(y,m,d)
            if self.db.has_day(dt): return True
            return bool(FeagaRef.day_in_any_window(dt))

        self.yearcal=YearCalendarFrame(
            cal_holder, year=date.today().year,
            on_day_click=lambda dt,**kw: (self._show_month(dt.year,dt.month) if kw.get("force_month") else self._show_day(dt)),
            on_day_context=self._on_day_context,
            has_events_predicate=has_ev
        )
        self.yearcal.pack(fill="both",expand=True)

        tab_idx=ttk.Frame(self.tabs); self.tabs.add(tab_idx,text="Índice")
        self._build_index_tab(tab_idx)

    # ---------- Utilidades UI ----------
    def _clear_view(self):
        """Borra el panel y limpia la selección del calendario (para empezar de cero)."""
        self._current_dt=None
        self.pay_frame.clear("—")
        self.yearcal.clear_selection()
        self.console.show("info","Panel limpio. Selecciona cualquier día del calendario.")

    # Índice
    def _build_index_tab(self,parent):
        top=ttk.Frame(parent); top.pack(fill="x",pady=(6,4))
        ttk.Label(top,text="Desde (dd/mm/aaaa):").pack(side="left",padx=(6,4))
        self.idx_from=ttk.Entry(top,width=12); self.idx_from.pack(side="left",padx=(0,10))
        ttk.Label(top,text="Hasta:").pack(side="left")
        self.idx_to=ttk.Entry(top,width=12); self.idx_to.pack(side="left",padx=(4,10))
        ttk.Button(top,text="Refrescar",command=self._refresh_index_tab).pack(side="left")
        ttk.Button(top,text="Ir a día",command=self._goto_from_index_tab).pack(side="left",padx=(10,0))

        cols=("fecha","tipo","fondo","detalle","origen","fuente")
        self.idx_tree=ttk.Treeview(parent,columns=cols,show="headings",selectmode="browse",style="Colored.Treeview")
        headers={"fecha":(92,"w"),"tipo":(180,"w"),"fondo":(80,"w"),"detalle":(560,"w"),"origen":(130,"center"),"fuente":(220,"w")}
        for c,(w,anc) in headers.items():
            self.idx_tree.heading(c,text=c.capitalize()); self.idx_tree.column(c,width=w,anchor=anc,stretch=True)
        self.idx_tree.pack(fill="both",expand=True,padx=6,pady=(0,6))
        ysb=ttk.Scrollbar(parent,orient="vertical",command=self.idx_tree.yview)
        self.idx_tree.configure(yscrollcommand=ysb.set)
        ysb.place(in_=self.idx_tree,relx=1.0,rely=0,relheight=1.0,x=-16)
        self.idx_tree.tag_configure("manual",background="#e8fff3")
        self.idx_tree.tag_configure("web",background="#fff7e6")
        self.idx_tree.tag_configure("heuristica",background="#eef5ff")
        self.idx_tree.tag_configure("info",background="#e7f3fe")
        self._refresh_index_tab()

    def _refresh_index_tab(self):
        d1=parse_ddmmyyyy(self.idx_from.get()) if self.idx_from.get().strip() else date(1900,1,1)
        d2=parse_ddmmyyyy(self.idx_to.get()) if self.idx_to.get().strip() else date(2100,12,31)
        rows=self.db.get_range(d1,d2,origins=None)
        self.idx_tree.delete(*self.idx_tree.get_children())
        for r in rows:
            label = PaymentsInfoFrame.ORIGIN_LABEL.get(r["origen"], r["origen"])
            self.idx_tree.insert("", "end",
                                 values=(r["fecha"],r["tipo"],r["fondo"],r["detalle"],label,r.get("fuente","")),
                                 tags=(r["origen"],))

    def _goto_from_index_tab(self):
        sel=self.idx_tree.selection()
        if not sel: messagebox.showinfo("Índice","Selecciona una fila."); return
        vals=self.idx_tree.item(sel[0],"values")
        try: dt=datetime.strptime(vals[0],"%Y-%m-%d").date()
        except Exception: messagebox.showerror("Índice","Fecha inválida."); return
        self.tabs.select(0); self.yearcal.go_to_date(dt); self._show_day(dt)

    # Contextual calendario
    def _on_day_context(self, ev, dt: date):
        menu=tk.Menu(self,tearoff=0)
        menu.add_command(label="Añadir pago…",command=lambda:self._add_payment_dialog(dt))
        menu.add_command(label="Borrar pagos del día (manual+web)…",command=lambda:self._delete_day(dt, include_heur=False))
        menu.add_command(label="Borrar pagos del día (TODO)…",command=lambda:self._delete_day(dt, include_heur=True))
        menu.add_separator()
        menu.add_command(label="Ver mes",command=lambda:self._show_month(dt.year,dt.month))
        try: menu.tk_popup(ev.x_root,ev.y_root)
        finally: menu.grab_release()

    # Filtros activos
    def _active_origins(self)->set[str]:
        s=set()
        if self.show_manual.get(): s.add("manual")
        if self.show_web.get(): s.add("web")
        if self.show_heur.get(): s.add("heuristica")
        return s or {"manual","web","heuristica"}

    # Mostrar día
    def _show_day(self, dt: date):
        self._current_dt=dt
        rows=self.db.get_day(dt, self._active_origins())
        for r in rows: r["fecha"]=fmt_dmy(datetime.strptime(r["fecha"],"%Y-%m-%d").date())
        if rows:
            self.pay_frame.show_rows(f"Día {fmt_dmy(dt)}", rows)
            self.console.show("ok", f"Mostrando {len(rows)} pago(s) para {fmt_dmy(dt)}.")
        else:
            gen = FeagaRef.day_in_any_window(dt) or FeagaRef.month_generic_for_day(dt)
            for r in gen: r["fecha"]=fmt_dmy(datetime.strptime(r["fecha"],"%Y-%m-%d").date())
            self.pay_frame.show_rows(f"Día {fmt_dmy(dt)} (referencia)", gen)
            self.console.show("info", f"No hay pagos guardados para {fmt_dmy(dt)}. Mostrando referencias FEAGA/FEADER.")

    def _show_month(self, y:int, m:int):
        rows=self.db.get_month(y,m, self._active_origins())
        for r in rows: r["fecha"]=fmt_dmy(datetime.strptime(r["fecha"],"%Y-%m-%d").date())
        if not rows:
            self.pay_frame.clear("—")
            self.console.show("warn", f"No hay pagos en {m:02d}/{y} con los filtros actuales.")
            return
        self.pay_frame.show_rows(f"Mes {m:02d}/{y}", rows)
        self.console.show("ok", f"Mostrando {len(rows)} elemento(s) del mes {m:02d}/{y}.")
        self.tabs.select(0)

    def _show_month_of_selected(self):
        dt=self.yearcal.get_selected_date() or date.today()
        self._show_month(dt.year, dt.month)

    # Rango
    def _show_range_dialog(self):
        dlg=tk.Toplevel(self); dlg.title("Buscar por rango"); dlg.transient(self.winfo_toplevel()); dlg.grab_set()
        ttk.Label(dlg,text="Desde (dd/mm/aaaa):").grid(row=0,column=0,sticky="e",padx=6,pady=6)
        ttk.Label(dlg,text="Hasta (dd/mm/aaaa):").grid(row=1,column=0,sticky="e",padx=6,pady=6)
        e1=ttk.Entry(dlg,width=14); e2=ttk.Entry(dlg,width=14)
        e1.grid(row=0,column=1,padx=6,pady=6); e2.grid(row=1,column=1,padx=6,pady=6)
        e1.insert(0, date.today().replace(day=1).strftime("%d/%m/%Y"))
        e2.insert(0, date.today().strftime("%d/%m/%Y"))
        btns=ttk.Frame(dlg); btns.grid(row=2,column=0,columnspan=2,pady=8)
        def ok():
            d1=parse_ddmmyyyy(e1.get()); d2=parse_ddmmyyyy(e2.get())
            if not d1 or not d2 or d1>d2: messagebox.showerror("Rango","Fechas inválidas."); return
            rows=self.db.get_range(d1,d2, self._active_origins())
            for r in rows: r["fecha"]=fmt_dmy(datetime.strptime(r["fecha"],"%Y-%m-%d").date())
            if not rows:
                self.pay_frame.clear("—"); self.console.show("warn","Sin resultados con los filtros actuales.")
            else:
                self.pay_frame.show_rows(f"Rango {fmt_dmy(d1)} – {fmt_dmy(d2)}", rows)
                self.console.show("ok", f"Mostrando {len(rows)} elemento(s) en el rango.")
            dlg.destroy()
        ttk.Button(btns,text="Buscar",command=ok).pack(side="left",padx=6)
        ttk.Button(btns,text="Cancelar",command=dlg.destroy).pack(side="left",padx=6)
        dlg.wait_window(dlg)

    # Altas/bajas
    def _add_payment_dialog(self, dt: date | None = None):
        dt = dt or (self.yearcal.get_selected_date() or date.today())
        win=tk.Toplevel(self); win.title("Añadir pago"); win.transient(self.winfo_toplevel()); win.grab_set()
        ttk.Label(win,text="Fecha (dd/mm/aaaa):").grid(row=0,column=0,sticky="e",padx=6,pady=4)
        ttk.Label(win,text="Tipo:").grid(row=1,column=0,sticky="e",padx=6,pady=4)
        ttk.Label(win,text="Fondo:").grid(row=2,column=0,sticky="e",padx=6,pady=4)
        ttk.Label(win,text="Detalle:").grid(row=3,column=0,sticky="ne",padx=6,pady=4)
        ttk.Label(win,text="Fuente (opcional):").grid(row=4,column=0,sticky="e",padx=6,pady=4)
        efecha=ttk.Entry(win,width=14); efecha.insert(0, fmt_dmy(dt)); efecha.grid(row=0,column=1,sticky="w",padx=6,pady=4)
        etipo=ttk.Entry(win,width=40); etipo.grid(row=1,column=1,sticky="we",padx=6,pady=4)
        efondo=ttk.Combobox(win,values=["FEAGA","FEADER","—"],width=12); efondo.set("FEAGA"); efondo.grid(row=2,column=1,sticky="w",padx=6,pady=4)
        tdetalle=tk.Text(win,width=60,height=5,wrap="word"); tdetalle.grid(row=3,column=1,sticky="we",padx=6,pady=4)
        efuente=ttk.Entry(win,width=60); efuente.grid(row=4,column=1,sticky="we",padx=6,pady=4)
        btns=ttk.Frame(win); btns.grid(row=5,column=0,columnspan=2,pady=8)
        def ok():
            d=parse_ddmmyyyy(efecha.get())
            if not d: messagebox.showerror("Añadir pago","Fecha inválida."); return
            tipo=etipo.get().strip(); fondo=efondo.get().strip() or "—"; det=tdetalle.get("1.0","end").strip()
            if not tipo or not det: messagebox.showerror("Añadir pago","Rellena Tipo y Detalle."); return
            self.db.add(d,tipo,fondo,det,efuente.get().strip(),origen="manual")
            self.yearcal.refresh(); self._refresh_current_view(); self.console.show("ok","Pago añadido.")
            win.destroy()
        ttk.Button(btns,text="Guardar",command=ok).pack(side="left",padx=6)
        ttk.Button(btns,text="Cancelar",command=win.destroy).pack(side="left",padx=6)
        win.wait_window(win)

    def _delete_day(self, dt: date, include_heur: bool):
        if include_heur:
            if not messagebox.askyesno("Borrar día","¿Borrar TODOS los pagos (incluye Referencia FEAGA) de ese día?"): return
            self.db.delete_day(dt, origen=None)
        else:
            if not messagebox.askyesno("Borrar día","¿Borrar SOLO pagos manual+web de ese día? (se conserva Referencia FEAGA)"): return
            self.db.delete_day(dt, origen="manual"); self.db.delete_day(dt, origen="web")
        self.yearcal.refresh(); self._refresh_current_view(); self.console.show("ok","Día actualizado.")

    def _delete_selected_day_dialog(self):
        dt=self.yearcal.get_selected_date() or date.today()
        dlg=tk.Toplevel(self); dlg.title("Borrar pagos del día"); dlg.transient(self.winfo_toplevel()); dlg.grab_set()
        ttk.Label(dlg,text=f"Día seleccionado: {fmt_dmy(dt)}").pack(padx=10,pady=(10,6))
        ttk.Button(dlg,text="Borrar manual+web",command=lambda:(dlg.destroy(), self._delete_day(dt, include_heur=False))).pack(padx=10,pady=4)
        ttk.Button(dlg,text="Borrar TODO (incluye Referencia FEAGA)",command=lambda:(dlg.destroy(), self._delete_day(dt, include_heur=True))).pack(padx=10,pady=(0,10))
        dlg.wait_window(dlg)

    def _clear_db(self, include_heur: bool):
        if include_heur:
            ok=messagebox.askyesno("Vaciar BD","¿Seguro que quieres borrar TODOS los pagos (incluye Referencia FEAGA)?")
        else:
            ok=messagebox.askyesno("Vaciar BD","¿Seguro que quieres borrar pagos MANUAL+WEB? (mantiene Referencia FEAGA)")
        if not ok: return
        self.db.delete_all(include_heuristic=include_heur)
        if not include_heur:
            FeagaRef.seed(self.db, FeagaRef.campaign_year_for(date.today()))
        self.yearcal.refresh(); self.pay_frame.clear("—"); self._refresh_index_tab()
        self.console.show("ok","Base de datos actualizada.")

    def _regen_heuristics(self):
        FeagaRef.seed(self.db, FeagaRef.campaign_year_for(date.today()))
        self.yearcal.refresh(); self._refresh_current_view()
        self.console.show("ok","Referencia FEAGA regenerada para la campaña actual.")

    def _help_clear(self):
        messagebox.showinfo(
            "¿Qué hace 'Vaciar manual+web'?",
            "Elimina SOLO los registros añadidos manualmente o descargados de la web.\n"
            "Mantiene las fechas de la 'Referencia FEAGA' (anticipos/saldos)."
        )

    def _refresh_current_view(self):
        if self._current_dt: self._show_day(self._current_dt)

    # Web (hilo seguro + messagebox)
    def _update_from_web(self):
        def run():
            try:
                sc=FegaWebScraper()
                if not sc.available(): raise RuntimeError("Falta 'requests'.")
                sc.fetch_into_db(self.db, year_hint=date.today().year)
                self.after(0, lambda:(self.yearcal.refresh(),
                                      self._refresh_current_view(),
                                      self._refresh_index_tab(),
                                      self.console.show("ok","Pagos actualizados desde FEAGA (web)."),
                                      messagebox.showinfo("Web","Pagos actualizados desde FEAGA (web).")))
            except Exception as ex:
                self.after(0, lambda: (self.console.show("error", f"No se pudo completar la descarga: {ex}"),
                                       messagebox.showerror("Web", f"No se pudo completar la descarga:\n{ex}")))
        threading.Thread(target=run,daemon=True).start()
        self.console.show("warn","Buscando ventanas en notas FEAGA… (se guardarán como origen 'web').")

# -------------------- main --------------------
def main():
    root=tk.Tk()
    root.title("Calendario FEAGA/FEADER – v3.4.0")
    root.geometry("1280x820")
    try:
        from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)
    except Exception: pass
    app=App(root); app.pack(fill="both",expand=True)
    root.mainloop()

if __name__=="__main__":
    main()
