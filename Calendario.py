# -*- coding: utf-8 -*-
"""
Calendario FEAGA/FEADER – v2.1.0
- Elimino el panel izquierdo vacío: los botones “Pagos de hoy” y “Mostrar pagos del mes” van a una barra superior.
- Calendario y panel de pagos ocupan TODO el ancho (sin huecos).
- Se mantiene: toggle por día, “Mostrar pagos del mes”, scraping/ingestor opcionales, pestaña Índice, popups modales propios.

Opcionales:
    pip install requests PyPDF2
"""

import calendar
import json
import re
import threading
from datetime import date, timedelta, datetime
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog

# -------------------- Utilidades --------------------
SPANISH_MONTHS = {
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"setiembre":9,"octubre":10,
    "noviembre":11,"diciembre":12
}
def parse_spanish_date(text, default_year=None):
    text = text.strip().lower()
    m = re.search(r"\b(\d{1,2})\s*(?:de\s*)?(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)(?:\s*de\s*(\d{4}))?\b", text)
    if m:
        d, mes, y = int(m.group(1)), SPANISH_MONTHS[m.group(2)], (int(m.group(3)) if m.group(3) else default_year or date.today().year)
        return (y, mes, d)
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b", text)
    if m:
        d, mes = int(m.group(1)), int(m.group(2))
        y = int(m.group(3)) if m.group(3) else (default_year or date.today().year)
        if y < 100: y += 2000
        return (y, mes, d)
    return None

def daterange(d1: date, d2: date):
    cur = d1
    while cur <= d2:
        yield cur
        cur += timedelta(days=1)

def parse_ddmmyyyy(s: str):
    s = s.strip()
    if not s: return None
    try:
        return datetime.strptime(s, "%d/%m/%Y").date()
    except Exception:
        return None

# -------------------- Índice de pagos (reales) --------------------
class PaymentsIndex:
    def __init__(self): self._by_date = {}
    def clear(self): self._by_date.clear()
    def add(self, dt: date, tipo: str, fondo: str, detalle: str, fuente: str = ""):
        self._by_date.setdefault(dt, []).append({"tipo":tipo,"fondo":fondo,"detalle":detalle,"fuente":fuente})
    def add_range(self, d1: date, d2: date, **kwargs):
        for dt in daterange(d1,d2): self.add(dt, **kwargs)
    def get_day(self, dt: date): return list(self._by_date.get(dt, []))
    def has_day(self, dt: date) -> bool: return bool(self._by_date.get(dt))
    def iter_all(self):
        for dt in sorted(self._by_date.keys()):
            for it in self._by_date[dt]:
                yield dt, it
    # Persistencia JSON
    def to_dict(self):
        out={}
        for dt, items in self._by_date.items():
            out[dt.isoformat()] = items
        return out
    def from_dict(self, data: dict):
        self._by_date.clear()
        for k, items in data.items():
            try:
                dt = datetime.fromisoformat(k).date()
            except Exception:
                continue
            self._by_date[dt] = list(items)
    def save_json(self, path: Path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)
    def load_json(self, path: Path):
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.from_dict(data)

# -------------------- Heurística FEAGA --------------------
class HeuristicPaymentsProvider:
    def get_for_day(self, dt: date):
        y, m = dt.year, dt.month
        campaign_year = y if m >= 10 else y - 1
        start_ant = date(campaign_year, 10, 16); end_ant = date(campaign_year, 11, 30)
        start_sal = date(campaign_year, 12, 1);  end_sal = date(campaign_year + 1, 6, 30)
        out = []
        if start_ant <= dt <= end_ant:
            out.append({"tipo":"Anticipo ayudas directas","fondo":"FEAGA",
                        "detalle":f"Hasta el 70% campaña {campaign_year}. Ventana 16/10–30/11.",
                        "fuente":"Heurística FEGA"})
        if (start_sal <= dt <= date(campaign_year,12,31)) or (date(campaign_year+1,1,1) <= dt <= end_sal):
            out.append({"tipo":"Saldo ayudas directas","fondo":"FEAGA",
                        "detalle":f"Hasta el 30% restante campaña {campaign_year}. Ventana 01/12–30/06 (año+1).",
                        "fuente":"Heurística FEGA"})
        if not out:
            out.append({"tipo":"Sin pagos FEAGA generales","fondo":"—",
                        "detalle":"Fuera de ventanas de anticipo/saldo. Revise resoluciones específicas.",
                        "fuente":"Heurística FEGA"})
        return out

# -------------------- Scraper / Ingestor (opcionales) --------------------
class FegaWebScraper:
    NOTE_URLS = [
        "https://www.fega.gob.es/sites/default/files/files/document/Nota_web_Ecorregimenes_Ca_2024_ANTICIPO.pdf",
        "https://www.fega.gob.es/sites/default/files/files/document/Nota_Web_AAS_Ca_2024_ANTICIPO.pdf",
        "https://www.fega.gob.es/sites/default/files/files/document/241115_NOTA_WEB_EERR_PRIMER_SALDO_Ca_2024_def.pdf",
    ]
    def __init__(self):
        try: import requests  # noqa
        except Exception: self._ok=False
        else: self._ok=True
    def available(self): return self._ok
    def fetch_into_index(self, index: PaymentsIndex, year_hint:int|None=None):
        if not self._ok: raise RuntimeError("Falta 'requests'.")
        import requests
        for url in self.NOTE_URLS:
            try:
                r = requests.get(url, timeout=15)
                txt = r.content.decode("latin-1", errors="ignore") if r.status_code==200 else url
            except Exception:
                txt = url
            m = re.search(r"(20\d{2})", txt)
            y = int(m.group(1)) if m else (year_hint or date.today().year)
            ant1 = parse_spanish_date("16 de octubre", y); ant2 = parse_spanish_date("30 de noviembre", y)
            sal1 = parse_spanish_date("1 de diciembre", y); sal2 = parse_spanish_date("30 de junio", y+1)
            if ant1 and ant2:
                index.add_range(date(*ant1), date(*ant2),
                                tipo="Anticipo ayudas directas", fondo="FEAGA",
                                detalle="Ventana general de anticipos (nota FEGA).", fuente=url)
            if sal1 and sal2:
                index.add_range(date(*sal1), date(*sal2),
                                tipo="Saldo ayudas directas", fondo="FEAGA",
                                detalle="Ventana general de saldos (nota FEGA).", fuente=url)

class FegaPDFIngestor:
    def __init__(self):
        try: import PyPDF2  # noqa
        except Exception: self._ok=False
        else: self._ok=True
    def available(self): return self._ok
    def _extract_text(self, path: Path) -> str:
        import PyPDF2
        txt=[]
        try:
            with open(path,"rb") as fh:
                reader=PyPDF2.PdfReader(fh)
                for p in reader.pages:
                    try: txt.append(p.extract_text() or "")
                    except Exception: pass
        except Exception: return ""
        return "\n".join(txt)
    def ingest_folder(self, folder: Path, index: PaymentsIndex, default_year:int|None=None):
        if not self._ok: raise RuntimeError("Falta 'PyPDF2'.")
        folder = Path(folder)
        if not folder.exists(): raise FileNotFoundError(str(folder))
        for pdf in folder.glob("*.pdf"):
            text = self._extract_text(pdf).lower()
            if not text: continue
            m = re.search(r"(20\d{2})", text)
            y = int(m.group(1)) if m else (default_year or date.today().year)
            for mo in re.finditer(r"del\s+(\d{1,2})\s+al\s+(\d{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)", text):
                d1,d2,mes = int(mo.group(1)), int(mo.group(2)), SPANISH_MONTHS[mo.group(3)]
                ctx = text[max(0,mo.start()-120):mo.end()+120]
                etiqueta,fondo="Pago/ventana","—"
                if "anticipo" in ctx: etiqueta,fondo="Anticipo ayudas directas","FEAGA"
                elif "saldo" in ctx: etiqueta,fondo="Saldo ayudas directas","FEAGA"
                elif "feader" in ctx: etiqueta,fondo="Pago medidas desarrollo rural","FEADER"
                index.add_range(date(y,mes,d1), date(y,mes,d2),
                                tipo=etiqueta, fondo=fondo,
                                detalle=f"Ventana detectada en {pdf.name}", fuente=str(pdf))
            for mo in re.finditer(r"\b(\d{1,2})\s*(?:de\s*)?(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\b", text):
                d, mes = int(mo.group(1)), SPANISH_MONTHS[mo.group(2)]
                ctx = text[max(0,mo.start()-140):mo.end()+140]
                etiqueta,fondo=None,"—"
                if "anticipo" in ctx: etiqueta,fondo="Anticipo ayudas directas","FEAGA"
                elif "saldo" in ctx: etiqueta,fondo="Saldo ayudas directas","FEAGA"
                elif "pago" in ctx and "feader" in ctx: etiqueta,fondo="Pago medidas desarrollo rural","FEADER"
                if etiqueta:
                    index.add(date(y,mes,d), tipo=etiqueta, fondo=fondo,
                              detalle=f"Fecha mencionada en {pdf.name}", fuente=str(pdf))

# -------------------- Scroll contenedor --------------------
class VerticalScrolledFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)
        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0,0), window=self.inner, anchor="nw")
        self.canvas.grid(row=0,column=0,sticky="nsew"); self.vsb.grid(row=0,column=1,sticky="ns")
        self.rowconfigure(0,weight=1); self.columnconfigure(0,weight=1)
        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        # rueda
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

# -------------------- Calendario anual --------------------
class YearCalendarFrame(ttk.Frame):
    MESES_ES=["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
    DIAS=["L","M","X","J","V","S","D"]

    def __init__(self, master, year=None, on_date_change=None, on_date_double_click=None,
                 on_day_click=None, has_events_predicate=None):
        super().__init__(master)
        calendar.setfirstweekday(calendar.MONDAY)
        self.current_year = year or date.today().year
        self.on_date_change = on_date_change
        self.on_date_double_click = on_date_double_click
        self.on_day_click = on_day_click
        self.has_events_predicate = has_events_predicate or (lambda y,m,d: False)

        if self.current_year == date.today().year:
            init = f"{self.current_year:04d}-{date.today().month:02d}-{date.today().day:02d}"
        else:
            init = f"{self.current_year:04d}-01-01"
        self.sel_var = tk.StringVar(value=init)
        self.sel_var.trace_add('write', lambda *args: self._notify_change())

        self._styles(); self._controls()

        self.scroll = VerticalScrolledFrame(self); self.scroll.grid(row=1,column=0,sticky="nsew")
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
        self.info=ttk.Label(top,text="(clic=toggle pagos · doble clic=insertar)")
        self.info.grid(row=0,column=5,sticky="e")

    def refresh(self):
        self._grid_year()
        self.ensure_selection_in_current_year()

    def _grid_year(self):
        for ch in self.scroll.inner.winfo_children(): ch.destroy()
        grid=ttk.Frame(self.scroll.inner); grid.grid(row=0,column=0,sticky="nsew")
        for r in range(4): grid.rowconfigure(r,weight=1,uniform="m")
        for c in range(3): grid.columnconfigure(c,weight=1,uniform="m")
        for month in range(1,13):
            r,c=(month-1)//3,(month-1)%3
            self._month(grid,self.current_year,month).grid(row=r,column=c,padx=4,pady=4,sticky="nsew")
        self.scroll.inner.rowconfigure(0,weight=1); self.scroll.inner.columnconfigure(0,weight=1)
        self._notify_change()

    def _month(self,parent,year,month):
        f=ttk.Frame(parent,borderwidth=1,relief="solid",padding=(4,3,4,4))
        ttk.Label(f,text=self.MESES_ES[month-1],font=("Segoe UI",9,"bold")).grid(row=0,column=0,columnspan=7,sticky="ew",pady=(0,2))
        for i,d in enumerate(self.DIAS):
            ttk.Label(f,text=d,style=("CalHeadWE.TLabel" if i in (5,6) else "CalHead.TLabel")).grid(row=1,column=i,padx=1,sticky="nsew")

        weeks=calendar.monthcalendar(year,month)
        font_norm=("Segoe UI",9); font_bold=("Segoe UI",9,"bold")
        color_weekend="#b00000"; color_event="#084f8a"

        for r,week in enumerate(weeks,start=2):
            for c in range(7):
                day=week[c]
                if day==0:
                    ttk.Label(f,text="").grid(row=r,column=c,padx=1,pady=1,sticky="nsew"); continue
                has=self.has_events_predicate(year,month,day)
                is_weekend = (c in (5,6))
                cfg = {"font": (font_bold if has else font_norm), "bd":1, "relief":"raised",
                       "highlightthickness":0, "indicatoron":0, "takefocus":0}
                if is_weekend: cfg["fg"]=color_weekend
                elif has: cfg["fg"]=color_event
                val=f"{year:04d}-{month:02d}-{day:02d}"
                rb = tk.Radiobutton(f, text=str(day), value=val, variable=self.sel_var, **cfg)
                rb.grid(row=r,column=c,padx=1,pady=1,sticky="nsew")

                def _on_release(e=None, v=val, y=year, m=month, d=day):
                    was_same = (self.sel_var.get() == v)
                    self.sel_var.set(v)  # si cambia -> on_date_change; si era el mismo -> toggle
                    if was_same and callable(self.on_day_click):
                        self.on_day_click(y,m,d)
                rb.bind("<ButtonRelease-1>", _on_release)

                rb.bind("<Double-Button-1>", lambda e,y=year,m=month,d=day:self._dbl(y,m,d))

        for c in range(7): f.columnconfigure(c,weight=1,uniform="d")
        f.grid_rowconfigure(0,minsize=18); f.grid_rowconfigure(1,minsize=18)
        for r in range(2,2+len(weeks)): f.grid_rowconfigure(r,weight=1,minsize=26)
        return f

    def go_to_date(self, dt: date):
        self._set(dt.year)
        self.sel_var.set(f"{dt.year:04d}-{dt.month:02d}-{dt.day:02d}")
    def get_selected_date(self):
        val=self.sel_var.get()
        if not val: return None
        try:
            y,m,d = map(int, val.split("-"))
            return date(y,m,d)
        except Exception:
            return None
    def ensure_selection_in_current_year(self):
        val = self.sel_var.get()
        try:
            y, m, d = map(int, val.split("-"))
        except Exception:
            if self.current_year == date.today().year:
                y, m, d = self.current_year, date.today().month, date.today().day
            else:
                y, m, d = self.current_year, 1, 1
            self.sel_var.set(f"{y:04d}-{m:02d}-{d:02d}")
            return
        if y != self.current_year:
            last_day = calendar.monthrange(self.current_year, m)[1]
            d = min(d, last_day)
            self.sel_var.set(f"{self.current_year:04d}-{m:02d}-{d:02d}")

    def _notify_change(self):
        val=self.sel_var.get()
        if not val: return
        try:
            y,m,d = map(int, val.split("-"))
        except Exception:
            return
        self.info.config(text=f"Seleccionado: {d:02d}/{m:02d}/{y}  (clic=toggle · doble clic=insertar)")
        if callable(self.on_date_change): self.on_date_change(y,m,d)
    def _dbl(self,y,m,d):
        self.sel_var.set(f"{y:04d}-{m:02d}-{d:02d}")
        if callable(self.on_date_double_click): self.on_date_double_click(y,m,d)
    def _set(self,year):
        try: year=int(year)
        except: year=self.current_year
        year=max(1900,min(2100,year))
        self.current_year=year; self.year_var.set(year); self._grid_year()
        self.ensure_selection_in_current_year()
    def _prev(self): self._set(self.current_year-1)
    def _next(self): self._set(self.current_year+1)
    def _go_today(self): self.go_to_date(date.today())

# -------------------- Mini-frame de pagos --------------------
class PaymentsInfoFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master,padding=(6,4,6,4))
        top=ttk.Frame(self); top.grid(row=0,column=0,sticky="ew")
        ttk.Label(top,text="Pagos en la fecha seleccionada",font=("Segoe UI",10,"bold")).pack(side="left")
        self.date_lbl=ttk.Label(top,text="—"); self.date_lbl.pack(side="right")

        cols=("fecha","tipo","fondo","detalle","fuente")
        self.tree=ttk.Treeview(self,columns=cols,show="headings")
        headers={"fecha":(92,"w"),"tipo":(180,"w"),"fondo":(80,"w"),"detalle":(640,"w"),"fuente":(240,"w")}
        for c,(w,anc) in headers.items():
            self.tree.heading(c,text=c.capitalize()); self.tree.column(c,width=w,anchor=anc,stretch=True)
        self.tree.grid(row=1,column=0,sticky="nsew",pady=(6,0))
        ysb=ttk.Scrollbar(self,orient="vertical",command=self.tree.yview)
        xsb=ttk.Scrollbar(self,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set,xscrollcommand=xsb.set)
        ysb.grid(row=1,column=1,sticky="ns"); xsb.grid(row=2,column=0,sticky="ew")
        self.grid_rowconfigure(1,weight=1); self.grid_columnconfigure(0,weight=1)

        ttk.Label(self,text="Nota: FEAGA 16/10–30/11 (anticipos), 01/12–30/06 (saldos). FEADER según resoluciones.",
                  foreground="#444").grid(row=3,column=0,columnspan=2,sticky="w",pady=(6,0))

    def clear(self, title="—"):
        self.date_lbl.config(text=title)
        self.tree.delete(*self.tree.get_children())
        self.update_idletasks()

    def show_list(self, title, rows):
        self.date_lbl.config(text=title)
        self.tree.delete(*self.tree.get_children())
        for it in rows:
            self.tree.insert("", "end", values=(it.get("fecha",""), it["tipo"], it["fondo"], it["detalle"], it.get("fuente","")))
        self.update_idletasks()

    def show_day(self, dt: date, items):
        rows=[]
        for it in items:
            row=dict(it); row["fecha"]=dt.strftime("%d/%m/%Y"); rows.append(row)
        self.show_list(dt.strftime("Día %d/%m/%Y"), rows)

    def show_month(self, y:int, m:int, dated_items):
        rows=[]
        for d,it in dated_items:
            row=dict(it); row["fecha"]=d.strftime("%d/%m/%Y"); rows.append(row)
        title=f"Mes {m:02d}/{y}"
        self.show_list(title, rows)

# -------------------- App principal --------------------
class CalendarioFEAGA_FEADERFrame(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.heur = HeuristicPaymentsProvider()
        self.index = PaymentsIndex()

        self.display_state = {"mode":"none", "day":None, "month":None}  # none|day|month

        self._build_ui()

        # Selección/pintado inicial: HOY (sin popup)
        self.yearcal.go_to_date(date.today())
        self.yearcal.ensure_selection_in_current_year()
        self._show_day(date.today(), allow_popup=False)
        self._update_today_banner()

    # -------- Popups modales propios --------
    def _popup(self, title, message, error=False):
        top = tk.Toplevel(self.winfo_toplevel()); top.title(title)
        top.transient(self.winfo_toplevel())
        try: top.attributes("-topmost", True)
        except Exception: pass
        top.grab_set()
        frm = ttk.Frame(top, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=message, justify="left",
                  foreground=("#b00000" if error else "#222")).pack(fill="x", expand=True)
        ttk.Button(frm, text="OK", command=top.destroy).pack(pady=(10,0))
        top.update_idletasks()
        try:
            px = self.winfo_rootx() + (self.winfo_width()//2 - top.winfo_width()//2)
            py = self.winfo_rooty() + (self.winfo_height()//2 - top.winfo_height()//2)
            top.geometry(f"+{max(0,px)}+{max(0,py)}")
        except Exception: pass
        top.wait_window(top)
    def _popup_info(self, title, message): self._popup(title, message, error=False)
    def _popup_error(self, title, message): self._popup(title, message, error=True)

    # ---- UI
    def _build_ui(self):
        self.pack(fill="both", expand=True)

        # Título
        ttk.Label(self, text="Calendario FEAGA / FEADER – v2.1.0",
                  font=("Segoe UI",13,"bold")).pack(padx=10, pady=(10,0), anchor="w")

        # Barra superior (dos botones) — SIN panel izquierdo
        topbar = ttk.Frame(self); topbar.pack(fill="x", padx=10, pady=(6,8))
        ttk.Button(topbar,text="Pagos de hoy",command=self._open_today).pack(side="left")
        ttk.Button(topbar,text="Mostrar pagos del mes",command=self._show_month_of_selected).pack(side="left", padx=(6,0))
        ttk.Label(topbar, text="Usa el calendario (clic=toggle) o 'Mostrar pagos del mes'.",
                  foreground="#555").pack(side="left", padx=12)

        # Pestañas
        self.tabs=ttk.Notebook(self); self.tabs.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Calendario
        tab_cal=ttk.Frame(self.tabs); self.tabs.add(tab_cal, text="Calendario")
        v_split=ttk.Panedwindow(tab_cal,orient="vertical"); v_split.pack(fill="both", expand=True)

        # Panel de pagos (ANTES del calendario para evitar callbacks prematuros)
        self.pay_frame=PaymentsInfoFrame(v_split)
        cal_holder=ttk.Frame(v_split)
        v_split.add(cal_holder,weight=3); v_split.add(self.pay_frame,weight=2)
        try:
            v_split.paneconfigure(cal_holder,minsize=140); v_split.paneconfigure(self.pay_frame,minsize=100)
        except Exception: pass

        # Herramientas del calendario
        tools=ttk.Frame(cal_holder); tools.pack(fill="x",pady=(0,4))
        ttk.Button(tools,text="Actualizar pagos (web)",command=self._update_from_web).pack(side="left")
        ttk.Button(tools,text="Importar circulares (PDF)…",command=self._import_pdfs).pack(side="left",padx=(6,0))
        ttk.Button(tools,text="Guardar índice…",command=self._save_index).pack(side="right")
        ttk.Button(tools,text="Cargar índice…",command=self._load_index).pack(side="right",padx=(6,0))
        ttk.Button(tools,text="Vaciar pagos (mantener heurística)",command=self._reset_index_only).pack(side="right",padx=(6,0))

        # Calendario
        self.yearcal=YearCalendarFrame(
            cal_holder, year=date.today().year,
            on_date_change=self._on_calendar_change,
            on_day_click=self._on_day_clicked_same,
            on_date_double_click=None,
            has_events_predicate=lambda y,m,d: self.index.has_day(date(y,m,d)) or any(
                i["tipo"]!="Sin pagos FEAGA generales" for i in self.heur.get_for_day(date(y,m,d))
            )
        )
        self.yearcal.pack(fill="both", expand=True)

        # --- Índice
        tab_idx=ttk.Frame(self.tabs); self.tabs.add(tab_idx,text="Índice")
        self._build_index_tab(tab_idx)

    # ---- Pestaña "Índice"
    def _build_index_tab(self, parent):
        top=ttk.Frame(parent); top.pack(fill="x", pady=(6,4))
        ttk.Label(top, text="Filtro texto:").pack(side="left", padx=(6,4))
        self.idx_filter_txt = ttk.Entry(top, width=30); self.idx_filter_txt.pack(side="left", padx=(0,10))
        ttk.Label(top, text="Desde (dd/mm/aaaa):").pack(side="left")
        self.idx_from = ttk.Entry(top, width=12); self.idx_from.pack(side="left", padx=(4,10))
        ttk.Label(top, text="Hasta:").pack(side="left")
        self.idx_to = ttk.Entry(top, width=12); self.idx_to.pack(side="left", padx=(4,10))
        ttk.Button(top, text="Refrescar", command=self._refresh_index_tab).pack(side="left")
        ttk.Button(top, text="Ir a día", command=self._goto_from_index_tab).pack(side="left", padx=(10,0))

        cols=("fecha","tipo","fondo","detalle","fuente")
        self.idx_tree=ttk.Treeview(parent, columns=cols, show="headings", selectmode="browse")
        headers={"fecha":(90,"w"),"tipo":(180,"w"),"fondo":(80,"w"),"detalle":(620,"w"),"fuente":(240,"w")}
        for c,(w,anc) in headers.items():
            self.idx_tree.heading(c, text=c.capitalize()); self.idx_tree.column(c, width=w, anchor=anc, stretch=True)
        self.idx_tree.pack(fill="both", expand=True, padx=6, pady=(0,6))
        ysb=ttk.Scrollbar(parent, orient="vertical", command=self.idx_tree.yview)
        self.idx_tree.configure(yscrollcommand=ysb.set)
        ysb.place(in_=self.idx_tree, relx=1.0, rely=0, relheight=1.0, x=-16)

        self._refresh_index_tab()

    def _refresh_index_tab(self):
        txt = self.idx_filter_txt.get().strip().lower()
        d1 = parse_ddmmyyyy(self.idx_from.get()) if self.idx_from.get().strip() else None
        d2 = parse_ddmmyyyy(self.idx_to.get()) if self.idx_to.get().strip() else None
        self.idx_tree.delete(*self.idx_tree.get_children())
        for dt, it in self.index.iter_all():
            if d1 and dt < d1: continue
            if d2 and dt > d2: continue
            blob = f"{it.get('tipo','')} {it.get('fondo','')} {it.get('detalle','')} {it.get('fuente','')}".lower()
            if txt and txt not in blob: continue
            self.idx_tree.insert("", "end", values=(dt.strftime("%d/%m/%Y"), it["tipo"], it["fondo"], it["detalle"], it.get("fuente","")))

    def _goto_from_index_tab(self):
        sel = self.idx_tree.selection()
        if not sel:
            self._popup_info("Índice", "Selecciona primero una fila del índice."); return
        vals = self.idx_tree.item(sel[0], "values")
        try:
            dt = datetime.strptime(vals[0], "%d/%m/%Y").date()
        except Exception:
            self._popup_error("Índice", "No se pudo interpretar la fecha seleccionada."); return
        for i in range(self.tabs.index("end")):
            if self.tabs.tab(i,"text")=="Calendario": self.tabs.select(i); break
        self.yearcal.go_to_date(dt)
        self._show_day(dt, allow_popup=False)

    # ---- Avisos “hoy”
    def _effective_items_for_day(self, dt: date):
        real = self.index.get_day(dt)
        heur = [i for i in self.heur.get_for_day(dt) if i["tipo"]!="Sin pagos FEAGA generales"]
        seen=set(); out=[]
        for it in real+heur:
            key=(it["tipo"],it["fondo"],it["detalle"])
            if key in seen: continue
            seen.add(key); out.append(it)
        return out, real

    def _effective_today_items(self):
        t = date.today()
        items,_ = self._effective_items_for_day(t)
        return items

    def _update_today_banner(self):
        items=self._effective_today_items()
        if items:
            tipos=", ".join(sorted({i["tipo"] for i in items}))
            self._banner_show(f"Hoy ({date.today().strftime('%d/%m/%Y')}) hay pagos FEAGA: {tipos}.")
        else:
            self._banner_hide()

    def _banner_show(self, text):
        if not hasattr(self, "_alert_bar"):
            self._alert_bar = tk.Frame(self, bg="#fff3cd", highlightbackground="#ffeeba", highlightthickness=1)
            self._alert_msg = tk.Label(self._alert_bar, text="", bg="#fff3cd", fg="#856404", font=("Segoe UI",9,"bold"))
            self._alert_msg.pack(side="left", padx=8, pady=4)
            tk.Button(self._alert_bar, text="Ver hoy", command=self._open_today, relief="groove").pack(side="right", padx=6, pady=4)
            tk.Button(self._alert_bar, text="X", command=lambda:self._alert_bar.pack_forget(), relief="flat", bg="#fff3cd").pack(side="right", padx=(0,6))
        self._alert_msg.config(text=text); self._alert_bar.pack(fill="x", padx=10, pady=(0,6))
    def _banner_hide(self):
        try: self._alert_bar.pack_forget()
        except Exception: pass

    def _select_calendar_tab(self):
        for i in range(self.tabs.index("end")):
            if self.tabs.tab(i,"text")=="Calendario":
                self.tabs.select(i); break

    def _open_today(self):
        self._select_calendar_tab()
        self.yearcal.go_to_date(date.today())
        self._show_day(date.today(), allow_popup=False)

    # ---- Carga desde FEGA (web/PDF)
    def _update_from_web(self):
        def run():
            try:
                sc=FegaWebScraper()
                if not sc.available(): raise RuntimeError("Falta 'requests'.")
                sc.fetch_into_index(self.index, year_hint=date.today().year)
                self.after(0, lambda:(self.yearcal.refresh(), self._update_today_banner(), self._refresh_index_tab(),
                                      self._popup_info("Listo","Pagos actualizados desde FEGA (web).")))
            except Exception as ex:
                self.after(0, lambda: self._popup_error("Error", f"No se pudo completar el scraping web:\n{ex}"))
        threading.Thread(target=run, daemon=True).start()
        self._popup_info("Actualizando","Buscando ventanas en notas FEGA…")

    def _import_pdfs(self):
        folder=filedialog.askdirectory(title="Selecciona la carpeta con las circulares FEGA (PDF)")
        if not folder: return
        def run():
            try:
                ing=FegaPDFIngestor()
                if not ing.available(): raise RuntimeError("Falta 'PyPDF2'.")
                ing.ingest_folder(Path(folder), self.index, default_year=date.today().year)
                self.after(0, lambda:(self.yearcal.refresh(), self._update_today_banner(), self._refresh_index_tab(),
                                      self._popup_info("Listo",f"Se importaron PDFs de {folder}")))
            except Exception as ex:
                self.after(0, lambda: self._popup_error("Error", f"No se pudieron importar los PDFs:\n{ex}"))
        threading.Thread(target=run, daemon=True).start()
        self._popup_info("Importando","Leyendo circulares PDF…")

    def _save_index(self):
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON","*.json")],
                                            title="Guardar índice de pagos", initialfile="indice_pagos.json")
        if not path: return
        try:
            self.index.save_json(Path(path))
            self._popup_info("Índice", f"Índice guardado en:\n{path}")
        except Exception as ex:
            self._popup_error("Índice", f"No se pudo guardar el índice:\n{ex}")

    def _load_index(self):
        path = filedialog.askopenfilename(filetypes=[("JSON","*.json")],
                                          title="Cargar índice de pagos")
        if not path: return
        try:
            self.index.load_json(Path(path))
            self.yearcal.refresh(); self._update_today_banner(); self._refresh_index_tab()
            self._popup_info("Índice", f"Índice cargado desde:\n{path}")
        except Exception as ex:
            self._popup_error("Índice", f"No se pudo cargar el índice:\n{ex}")

    def _reset_index_only(self):
        self.index.clear()
        self.yearcal.refresh(); self._update_today_banner(); self._refresh_index_tab()
        self._popup_info("Pagos","Índice vaciado. Se mantiene la heurística (no persistente).")

    # ---- Integración calendario / toggle día
    def _on_calendar_change(self,y,m,d):
        dt = date(y,m,d)
        if not hasattr(self, "pay_frame"): return
        self._show_day(dt, allow_popup=True)

    def _on_day_clicked_same(self,y,m,d):
        dt = date(y,m,d)
        if self.display_state["mode"]=="day" and self.display_state["day"]==dt:
            self.pay_frame.clear("—")
            self.display_state = {"mode":"none","day":None,"month":None}
        else:
            self._show_day(dt, allow_popup=True)

    def _show_day(self, dt: date, allow_popup: bool):
        items, real = self._effective_items_for_day(dt)
        if not items:
            if allow_popup:
                self._popup_info("Pagos", f"No hay pagos para la fecha seleccionada ({dt.strftime('%d/%m/%Y')}).")
            self.pay_frame.clear("—")
            self.display_state = {"mode":"none","day":None,"month":None}
            return
        self.pay_frame.show_day(dt, items)
        self.display_state = {"mode":"day","day":dt,"month":None}

    # ---- Resumen mensual
    def _show_month_of_selected(self):
        dt = self.yearcal.get_selected_date() or date.today()
        y, m = dt.year, dt.month
        first = date(y,m,1)
        last_day = calendar.monthrange(y,m)[1]
        last = date(y,m,last_day)
        rows=[]; seen=set()
        for d in daterange(first,last):
            items,_ = self._effective_items_for_day(d)
            for it in items:
                key=(d, it["tipo"], it["fondo"], it["detalle"])
                if key in seen: continue
                seen.add(key); rows.append((d,it))
        if not rows:
            self._popup_info("Pagos del mes", f"No hay pagos en {m:02d}/{y}.")
            self.pay_frame.clear("—")
            self.display_state = {"mode":"none","day":None,"month":None}
            return
        self.pay_frame.show_month(y,m, rows)
        self.display_state = {"mode":"month","day":None,"month":(y,m)}
        self._select_calendar_tab()

# -------------------- main --------------------
def main():
    root=tk.Tk()
    root.title("Calendario FEAGA/FEADER – v2.1.0")
    root.geometry("1280x820")
    try:
        from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)
    except Exception: pass
    try:
        import platform; ttk.Style().theme_use("winnative" if platform.system()=="Windows" else "clam")
    except Exception: pass
    app=CalendarioFEAGA_FEADERFrame(root); app.pack(fill="both", expand=True)
    root.mainloop()

if __name__=="__main__":
    main()
