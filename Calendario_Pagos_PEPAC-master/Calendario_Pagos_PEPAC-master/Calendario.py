# -*- coding: utf-8 -*-
"""
Calendario FEAGA/FEADER – [v4.2.0 estable VERSION]

Novedades en esta versión:
- Hitos del mes (visual): ventana normal con botones de minimizar/maximizar/cerrar, redimensionable.
  Incluye panel de DETALLE a la derecha: clic en un día => lista completa del día (sin recortes).
  Doble clic en un día => salta al listado principal (como antes).
- Actualizar pagos (web): muestra ventana emergente con las FUENTES consultadas y resumen final.
- Resto de funcionalidad se mantiene. Auto-carga de Excel de Aragón incluida (1 vez por año).
- NUEVO: Centro de Ayuda (¿? y tecla F1) + tooltips en botones.

Dependencias opcionales:
  pip install pandas openpyxl requests beautifulsoup4 lxml
"""

import calendar
import re
import sqlite3
import threading
from datetime import date, datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from collections import defaultdict, Counter
from contextlib import suppress
import unicodedata
from pathlib import Path

# -------------------- Config --------------------
DB_FILE = "pagos_pepac.sqlite3"

# Excel de Aragón para autoload (se intenta al arrancar; si no existe o faltan deps, se ignora)
AUTOLOAD_ARAGON_EXCEL = Path("/mnt/data/tabla_FEAGA_FEADER_Aragon.xlsx")
AUTOLOAD_FALLBACKS = [Path("tabla_FEAGA_FEADER_Aragon.xlsx"), Path("data/tabla_FEAGA_FEADER_Aragon.xlsx")]

def iso(d: date) -> str: return d.strftime("%Y-%m-%d")
def fmt_dmy(d: date) -> str: return d.strftime("%d/%m/%Y")

def parse_ddmmyyyy(s: str) -> date | None:
    s = (s or "").strip()
    if not s: return None
    for fmt in ("%d/%m/%Y","%d/%m/%y","%Y-%m-%d"):
        try: return datetime.strptime(s, fmt).date()
        except Exception: pass
    try:
        s2 = s.replace("-", "/")
        return datetime.strptime(s2, "%d/%m/%Y").date()
    except Exception:
        return None

def daterange(d1: date, d2: date):
    cur = d1
    while cur <= d2:
        yield cur
        cur += timedelta(days=1)

def strip_accents_lower(s: str) -> str:
    s = str(s or "")
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn").lower()

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
                    origen TEXT NOT NULL CHECK (origen IN ('manual','web','heuristica','info')),
                    created_at TEXT DEFAULT (datetime('now'))
                )
            """)
            c.execute("""CREATE UNIQUE INDEX IF NOT EXISTS ux_pagos
                         ON pagos(fecha,tipo,fondo,detalle,origen)""")
            c.execute("""
                CREATE TABLE IF NOT EXISTS app_meta(
                    k TEXT PRIMARY KEY,
                    v TEXT
                )
            """)
            self.conn.commit()

    # ---- metadatos ----
    def get_meta(self, key: str) -> str | None:
        with self._lock:
            r = self.conn.execute("SELECT v FROM app_meta WHERE k=?", (key,)).fetchone()
            return r["v"] if r else None

    def set_meta(self, key: str, value: str):
        with self._lock:
            self.conn.execute("INSERT INTO app_meta(k,v) VALUES(?,?) ON CONFLICT(k) DO UPDATE SET v=excluded.v", (key, value))
            self.conn.commit()

    # ---- métricas ----
    def count_rows(self) -> int:
        with self._lock:
            r = self.conn.execute("SELECT COUNT(*) AS n FROM pagos").fetchone()
            return int(r["n"] or 0)

    # ---- I/O pagos ----
    def add(self, d: date, tipo: str, fondo: str, detalle: str, fuente: str="", origen: str="manual"):
        with self._lock:
            self.conn.execute(
                "INSERT OR IGNORE INTO pagos(fecha,tipo,fondo,detalle,fuente,origen) VALUES (?,?,?,?,?,?)",
                (iso(d), (tipo or "").strip(), (fondo or "—").strip(), (detalle or "").strip(), (fuente or "").strip(), origen))
            self.conn.commit()

    def add_range(self, d1: date, d2: date, **kwargs):
        with self._lock:
            cur=self.conn.cursor()
            for d in daterange(d1,d2):
                cur.execute("INSERT OR IGNORE INTO pagos(fecha,tipo,fondo,detalle,fuente,origen) VALUES (?,?,?,?,?,?)",
                            (iso(d), kwargs.get("tipo",""), kwargs.get("fondo","—"),
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
        rows.append({"fecha": iso(d), "tipo": "Referencia: FEADER (desarrollo rural)",
                     "fondo":"FEADER", "detalle":"Pagos según resoluciones/convocatorias autonómicas.",
                     "fuente":"Referencia", "origen":"info"})
        return rows

# -------------------- Helpers visual --------------------
def group_by_fondo(rows: list[dict]) -> dict[str, list[dict]]:
    buckets = defaultdict(list)
    for r in rows:
        buckets[(r.get("fondo") or "—").upper()].append(r)
    return buckets

def top_k_types(rows: list[dict], k=3) -> list[str]:
    cnt = Counter((r.get("tipo","").strip() or "—") for r in rows)
    return [t for t,_ in cnt.most_common(k)]

def short_date_esp(iso_str: str) -> str:
    try:
        return datetime.strptime(iso_str, "%Y-%m-%d").strftime("%d/%m/%Y")
    except Exception:
        return iso_str

def recast_as_month_item(day_dt: date, original_row: dict) -> dict:
    row = dict(original_row)
    iso_day = iso(day_dt)
    original_iso = row.get("fecha", iso_day)
    row["fecha"] = iso_day
    row["origen"] = "info"
    pref = "Del mes · "
    row["detalle"] = f"{pref}{row.get('detalle','')}".strip()
    if original_iso and original_iso != iso_day:
        row["detalle"] += f" (original: {short_date_esp(original_iso)})"
    return row

# -------------------- Scraper web (FEGA) --------------------
class FegaWebScraper:
    NOTE_URLS = [
        ("Anticipo ecorregímenes", "https://www.fega.gob.es/sites/default/files/files/document/Nota_web_Ecorregimenes_Ca_2024_ANTICIPO.pdf"),
        ("Anticipo ayudas asociadas", "https://www.fega.gob.es/sites/default/files/files/document/Nota_Web_AAS_Ca_2024_ANTICIPO.pdf"),
        ("Saldo ecorregímenes (EERR)", "https://www.fega.gob.es/sites/default/files/files/document/241115_NOTA_WEB_EERR_PRIMER_SALDO_Ca_2024_def.pdf"),
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
        for etiqueta, url in self.NOTE_URLS:
            try:
                r=requests.get(url,timeout=15)
                content=r.content.decode("latin-1",errors="ignore") if r.status_code==200 else url
            except Exception:
                content=url
            m=re.search(r"(20\d{2})", content)
            y=int(m.group(1)) if m else (year_hint or date.today().year)
            ant1=date(y,10,16); ant2=date(y,11,30)
            sal1=date(y,12,1);  sal2=date(y+1,6,30)
            if "Anticipo" in etiqueta:
                db.add_range(ant1, ant2, tipo=etiqueta, fondo="FEAGA",
                             detalle="Ventana general de anticipos (nota FEGA).",
                             fuente=url, origen="web")
            else:
                db.add_range(sal1, sal2, tipo=etiqueta, fondo="FEAGA",
                             detalle="Ventana general de saldos (nota FEGA).",
                             fuente=url, origen="web")

# -------------------- Scraper multi-fuente --------------------
class MultiSourceScraper:
    EXTRA_SOURCES = [
        # Ejemplo: ("Aragón – Noticias PAC", "https://www.aragon.es/en/-/noticias-pac"),
    ]
    def __init__(self):
        try:
            import requests  # noqa
            self._ok = True
        except Exception:
            self._ok = False
    def available(self): return self._ok

    def fetch_into_db(self, db: PaymentsDB, year_hint:int|None=None):
        if not self._ok: raise RuntimeError("Falta 'requests'.")
        # FEGA (PDFs)
        FegaWebScraper().fetch_into_db(db, year_hint=year_hint)
        # Noticias FEGA
        with suppress(Exception):
            self._fetch_fega_news(db, year_hint)
        # Fuentes extra declaradas
        for name, url in self.EXTRA_SOURCES:
            with suppress(Exception):
                self._fetch_generic_html(db, url, label=name)

    def _fetch_fega_news(self, db: PaymentsDB, year_hint:int|None):
        import requests
        base = "https://www.fega.gob.es/es/noticias"
        r = requests.get(base, timeout=15)
        if r.status_code != 200: return
        html = r.text
        links = re.findall(r'href="([^"]+)"[^>]*>(.*?)</a>', html, flags=re.I|re.S)
        for href, title in links:
            t = re.sub(r"<.*?>", "", title).strip()
            if not t: continue
            tl = strip_accents_lower(t)
            if not any(k in tl for k in ("anticipo","saldo","ecorreg","asociad","pago")):
                continue
            m = re.search(r"(\d{1,2}[/-]\d{1,2}[/-](20\d{2}))", href + " " + t)
            y = None
            if m:
                d = parse_ddmmyyyy(m.group(1))
                y = d.year if isinstance(d, date) else None
            y = y or year_hint or date.today().year
            etiqueta = t[:120]
            if any(k in tl for k in ("anticipo","ecorreg","asociad")):
                ant1=date(y,10,16); ant2=date(y,11,30)
                db.add_range(ant1, ant2, tipo=f"[Noticias FEGA] {etiqueta}", fondo="FEAGA",
                             detalle="Ventana general (anticipo) detectada en noticias FEGA.",
                             fuente=(href if href.startswith("http") else base), origen="web")
            elif "saldo" in tl:
                sal1=date(y,12,1);  sal2=date(y+1,6,30)
                db.add_range(sal1, sal2, tipo=f"[Noticias FEGA] {etiqueta}", fondo="FEAGA",
                             detalle="Ventana general (saldo) detectada en noticias FEGA.",
                             fuente=(href if href.startswith("http") else base), origen="web")

    def _fetch_generic_html(self, db: PaymentsDB, url: str, label: str):
        import requests
        r = requests.get(url, timeout=15)
        if r.status_code != 200: return
        text = r.text
        links = re.findall(r'href="([^"]+)"[^>]*>(.*?)</a>', text, flags=re.I|re.S)
        for href, title in links:
            t = re.sub(r"<.*?>", "", title).strip()
            if not t: continue
            fondo = "FEAGA" if "feaga" in strip_accents_lower(t) else ("FEADER" if "feader" in strip_accents_lower(t) else "—")
            if fondo == "—": continue
            m = re.search(r"(20\d{2})", href + " " + t)
            y = int(m.group(1)) if m else date.today().year
            if "anticipo" in strip_accents_lower(t):
                ant1=date(y,10,16); ant2=date(y,11,30)
                db.add_range(ant1, ant2, tipo=f"[{label}] {t[:120]}", fondo=fondo,
                             detalle="Ventana (anticipo) detectada.",
                             fuente=(href if href.startswith("http") else url), origen="web")
            elif "saldo" in strip_accents_lower(t):
                sal1=date(y,12,1);  sal2=date(y+1,6,30)
                db.add_range(sal1, sal2, tipo=f"[{label}] {t[:120]}", fondo=fondo,
                             detalle="Ventana (saldo) detectada.",
                             fuente=(href if href.startswith("http") else url), origen="web")

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

# -------------------- Calendario con botones --------------------
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

# -------------------- Ventana de “Fuentes consultadas” --------------------
class WebSourcesDialog(tk.Toplevel):
    def __init__(self, parent, sources:list[tuple[str,str]]):
        super().__init__(parent)
        self.title("Fuentes consultadas en la actualización web")
        # Ventana normal, con controles estándar:
        self.resizable(True, True)
        self.geometry("700x420+120+120")
        self.configure(bg="#f7f9fc")
        tk.Label(self, text="Se consultarán las siguientes fuentes:", bg="#f7f9fc",
                 fg="#093a76", font=("Segoe UI",10,"bold")).pack(anchor="w", padx=10, pady=(10,6))
        self.tree=ttk.Treeview(self, columns=("url","estado"), show="headings")
        self.tree.heading("url", text="Fuente")
        self.tree.heading("estado", text="Estado")
        self.tree.column("url", width=520, anchor="w")
        self.tree.column("estado", width=120, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=10, pady=(0,8))
        self._rows=[]
        for name,url in sources:
            iid=self.tree.insert("", "end", values=(f"{name}  —  {url}", "pendiente"))
            self._rows.append(iid)
        self._footer=tk.Label(self, text="Listo para iniciar.", bg="#f7f9fc")
        self._footer.pack(fill="x", padx=10, pady=(0,10))
        btns=tk.Frame(self, bg="#f7f9fc"); btns.pack(fill="x", padx=10, pady=(0,10))
        self._close=tk.Button(btns, text="Cerrar", state="disabled", command=self.destroy)
        self._close.pack(side="right")

    def mark_running(self):
        for iid in self._rows:
            self.tree.set(iid, "estado", "consultando…")
        self._footer.configure(text="Consultando fuentes…")

    def mark_done(self, added:int):
        for iid in self._rows:
            if self.tree.set(iid, "estado")=="consultando…":
                self.tree.set(iid, "estado", "OK")
        self._footer.configure(text=f"Completado. Registros añadidos: {added}.")
        self._close.configure(state="normal")

# -------------------- Hitos del mes (visual) con panel de detalle --------------------
class MonthHitosDialog(tk.Toplevel):
    """Ventana normal redimensionable con grid mensual a la izquierda y detalle a la derecha."""
    def __init__(self, app: "App", year: int, month: int):
        super().__init__(app)
        self.app = app
        self.title(f"Hitos {month:02d}/{year}")
        # Ventana normal (sin transient ni grab) => botones minimizar/maximizar/cerrar disponibles
        self.resizable(True, True)
        self.configure(bg="#f7f9fc")
        self.minsize(960, 620)
        self._build(year, month)

    def _build(self, y:int, m:int):
        # Barra superior
        top = tk.Frame(self, bg="#e7f0ff", highlightbackground="#cddcfb", highlightthickness=1)
        tk.Label(top, text=f"Hitos del mes {m:02d}/{y}", bg="#e7f0ff", fg="#093a76",
                 font=("Segoe UI",11,"bold")).pack(side="left", padx=10, pady=6)
        tk.Button(top, text="Maximizar", command=lambda:self.state("zoomed")).pack(side="left", padx=(8,0), pady=6)
        tk.Button(top, text="Exportar mes (CSV)", command=lambda:self._export_csv(y,m)).pack(side="right", padx=8, pady=6)
        top.grid(row=0, column=0, sticky="ew")
        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure(0, weight=1)

        # PanedWindow izquierda (calendario) / derecha (detalle)
        paned = ttk.Panedwindow(self, orient="horizontal")
        paned.grid(row=1, column=0, sticky="nsew")
        self.grid_rowconfigure(1, weight=1)

        left = tk.Frame(paned, bg="#f7f9fc")
        right = tk.Frame(paned, bg="#ffffff")
        paned.add(left, weight=3)   # (ttk no usa weight real, pero permite arrastrar la divisoria)
        paned.add(right, weight=2)

        # Leyenda abajo
        leg = tk.Frame(self, bg="#f7f9fc"); leg.grid(row=2, column=0, sticky="ew", padx=10, pady=(6,6))
        def badge(color, txt):
            box=tk.Frame(leg,bg=color,width=16,height=12,highlightthickness=1,highlightbackground="#aaa")
            box.pack(side="left", padx=(6,4), pady=2)
            tk.Label(leg, text=txt, bg="#f7f9fc").pack(side="left", padx=(0,10))
        badge("#e8fff3","Manual / Web"); badge("#eef5ff","Referencia FEAGA"); badge("#e7f3fe","Información (del mes)")

        # ======= Lado izquierdo: grid mensual responsive =======
        grid = tk.Frame(left, bg="#f7f9fc")
        grid.pack(fill="both", expand=True, padx=8, pady=8)

        month_rows = self.app.db.get_month(y, m, origins=None)
        by_day = defaultdict(list)
        for r in month_rows:
            by_day[r["fecha"]].append(r)

        weeks = calendar.monthcalendar(y, m)
        rows_cnt = len(weeks)

        for rr in range(rows_cnt): grid.grid_rowconfigure(rr, weight=1, uniform="rows")
        for cc in range(7): grid.grid_columnconfigure(cc, weight=1, uniform="cols")

        for rr, week in enumerate(weeks):
            for cc, dd in enumerate(week):
                cell = tk.Frame(grid, bd=1, relief="solid", bg="#ffffff")
                cell.grid(row=rr, column=cc, padx=3, pady=3, sticky="nsew")
                if dd == 0:
                    continue
                dt = date(y,m,dd); iso_dt = iso(dt)
                rows = list(by_day.get(iso_dt, []))
                for ref in FeagaRef.day_in_any_window(dt): rows.append(ref)

                # Encabezado: día
                hdr = tk.Frame(cell, bg="#f1f6ff")
                hdr.pack(fill="x")
                tk.Label(hdr, text=str(dd), bg="#f1f6ff", fg="#093a76",
                         font=("Segoe UI",10,"bold")).pack(side="left", padx=6)

                # Totales
                buckets = group_by_fondo(rows)
                line = tk.Frame(cell, bg="#ffffff"); line.pack(fill="x", padx=6, pady=(2,0))
                feaga_n = len(buckets.get("FEAGA", []))
                feader_n = len(buckets.get("FEADER", []))
                tk.Label(line, text=f"FEAGA: {feaga_n}", bg="#ffffff").pack(side="left")
                tk.Label(line, text=f"  FEADER: {feader_n}", bg="#ffffff").pack(side="left", padx=(8,0))

                # Chips (muestra tipos principales; el detalle va a la derecha)
                chips = top_k_types(rows, k=3)
                chipsf = tk.Frame(cell, bg="#ffffff"); chipsf.pack(fill="both", expand=True, padx=6, pady=(4,4))
                chip_labels=[]
                for t in chips:
                    lab = tk.Label(chipsf, text=f"• {t}", bg="#eef5ff", fg="#093a76", justify="left", anchor="w")
                    lab.pack(anchor="w", fill="x", pady=1)
                    chip_labels.append(lab)
                def _on_conf(e, labs=chip_labels):
                    wl = max(60, e.width - 16)
                    for lb in labs: lb.configure(wraplength=wl)
                cell.bind("<Configure>", _on_conf)

                # Clic = previsualiza detalle en el panel derecho. Doble clic = ir al día en la app.
                cell.bind("<Button-1>", lambda _e, dti=dt: self._show_detail(right, dti))
                for w in cell.winfo_children():
                    w.bind("<Button-1>", lambda _e, dti=dt: self._show_detail(right, dti))
                    w.bind("<Double-Button-1>", lambda _e, dti=dt: self._goto_day(dti))

        # ======= Lado derecho: panel de detalle =======
        self._build_detail_panel(right)
        # Selección inicial (hoy si cae en el mes, si no, día 1)
        init_day = date.today() if (date.today().year==y and date.today().month==m) else date(y,m,1)
        self._show_detail(right, init_day)

    def _build_detail_panel(self, holder: tk.Frame):
        head = tk.Frame(holder, bg="#eaf2ff", highlightbackground="#cddcfb", highlightthickness=1)
        self._detail_title = tk.Label(head, text="Detalle —", bg="#eaf2ff", fg="#093a76", font=("Segoe UI",10,"bold"))
        self._detail_title.pack(side="left", padx=8, pady=6)
        tk.Button(head, text="Abrir en listado principal", command=self._open_in_main).pack(side="right", padx=8, pady=6)
        head.pack(fill="x")

        cols=("fecha","tipo","fondo","detalle","origen","fuente")
        self.detail_tree=ttk.Treeview(holder, columns=cols, show="headings")
        for c,w,anc in (("fecha",92,"w"),("tipo",220,"w"),("fondo",80,"center"),
                        ("detalle",560,"w"),("origen",120,"center"),("fuente",220,"w")):
            self.detail_tree.heading(c, text=c.capitalize()); self.detail_tree.column(c, width=w, anchor=anc, stretch=True)
        ysb=ttk.Scrollbar(holder,orient="vertical",command=self.detail_tree.yview)
        xsb=ttk.Scrollbar(holder,orient="horizontal",command=self.detail_tree.xview)
        self.detail_tree.configure(yscrollcommand=ysb.set,xscrollcommand=xsb.set)
        self.detail_tree.pack(fill="both", expand=True, side="left", padx=6, pady=6)
        ysb.pack(fill="y", side="right"); xsb.pack(fill="x", side="bottom")
        self._detail_dt = None

    def _show_detail(self, holder: tk.Frame, d: date):
        self._detail_dt = d
        self._detail_title.configure(text=f"Detalle — {fmt_dmy(d)}")
        rows = self.app.db.get_day(d, origins=None)
        # Añadimos referencia FEAGA/FEADER informativa si no hay
        if not rows:
            rows = FeagaRef.day_in_any_window(d)
        else:
            for ref in FeagaRef.day_in_any_window(d):
                rows.append(ref)
        # Mostrar
        self.detail_tree.delete(*self.detail_tree.get_children())
        for r in rows:
            vals=(r.get("fecha",""), r.get("tipo",""), r.get("fondo",""), r.get("detalle",""),
                  r.get("origen",""), r.get("fuente",""))
            self.detail_tree.insert("", "end", values=vals)

    def _open_in_main(self):
        if self._detail_dt:
            self.app.tabs.select(0)
            self.app.yearcal.go_to_date(self._detail_dt)
            self.app._show_day(self._detail_dt)

    def _goto_day(self, d: date):
        self._open_in_main()

    def _export_csv(self, y:int, m:int):
        import csv, os
        rows = self.app.db.get_month(y, m, origins=None)
        if not rows:
            messagebox.showinfo("Exportar", "No hay filas para ese mes."); return
        path = f"pagos_{y}_{m:02d}.csv"
        with open(path, "w", newline="", encoding="utf-8") as f:
            w = csv.writer(f, delimiter=";")
            w.writerow(["fecha","tipo","fondo","detalle","origen","fuente"])
            for r in rows:
                w.writerow([r.get("fecha",""), r.get("tipo",""), r.get("fondo",""), r.get("detalle",""), r.get("origen",""), r.get("fuente","")])
        messagebox.showinfo("Exportar", f"Exportado: {os.path.abspath(path)}")

# -------------------- ToolTip (globos de ayuda) --------------------
class ToolTip:
    def __init__(self, widget, text, delay=650):
        self.widget = widget
        self.text = text
        self.delay = delay
        self._id = None
        self._tip = None
        widget.bind("<Enter>", self._schedule)
        widget.bind("<Leave>", self._hide)
        widget.bind("<Button-1>", self._hide)
        widget.bind("<Motion>", self._move)

    def _schedule(self, _e=None):
        self._cancel()
        self._id = self.widget.after(self.delay, self._show)

    def _cancel(self):
        if self._id:
            try: self.widget.after_cancel(self._id)
            except Exception: pass
            self._id = None

    def _show(self):
        if self._tip or not self.widget.winfo_viewable():
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        self._tip = tk.Toplevel(self.widget)
        self._tip.wm_overrideredirect(True)
        self._tip.attributes("-topmost", True)
        frm = tk.Frame(self._tip, bg="#333", padx=8, pady=6)
        frm.pack()
        lab = tk.Label(frm, text=self.text, bg="#333", fg="#fff",
                       justify="left", wraplength=360)
        lab.pack()
        self._tip.geometry(f"+{x}+{y}")

    def _hide(self, _e=None):
        self._cancel()
        if self._tip:
            try: self._tip.destroy()
            except Exception: pass
            self._tip = None

    def _move(self, e):
        if self._tip:
            self._tip.geometry(f"+{e.x_root+14}+{e.y_root+14}")

# -------------------- Centro de Ayuda --------------------
class HelpCenterDialog(tk.Toplevel):
    def __init__(self, app: "App"):
        super().__init__(app)
        self.app = app
        self.title("Ayuda · Calendario FEAGA/FEADER")
        self.geometry("980x620+120+90")
        self.minsize(820, 520)
        self.configure(bg="#f7f9fc")

        # Top: búsqueda
        top = tk.Frame(self, bg="#e7f0ff", highlightbackground="#cddcfb", highlightthickness=1)
        tk.Label(top, text="Centro de ayuda", bg="#e7f0ff", fg="#093a76",
                 font=("Segoe UI",11,"bold")).pack(side="left", padx=10, pady=6)
        tk.Label(top, text="Buscar:", bg="#e7f0ff").pack(side="left", padx=(12,0))
        self.q = tk.Entry(top, width=32)
        self.q.pack(side="left", padx=6)
        self.q.bind("<KeyRelease>", self._filter)
        tk.Button(top, text="Cerrar", command=self.destroy).pack(side="right", padx=8)
        top.grid(row=0, column=0, sticky="ew")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Paned: índice / contenido
        pan = ttk.Panedwindow(self, orient="horizontal")
        pan.grid(row=1, column=0, sticky="nsew")

        left = tk.Frame(pan, bg="#f7f9fc")
        right = tk.Frame(pan, bg="#ffffff")
        pan.add(left, weight=1)
        pan.add(right, weight=3)

        # Índice
        self.list = tk.Listbox(left, activestyle="dotbox")
        self.list.pack(fill="both", expand=True, padx=8, pady=8)
        self.list.bind("<<ListboxSelect>>", self._on_select)
        self.list.bind("<Return>", self._on_open)

        # Contenido
        self.text = tk.Text(right, wrap="word", bg="#ffffff")
        self.text.pack(fill="both", expand=True, padx=8, pady=8)
        self.text.configure(state="disabled")
        ysb = ttk.Scrollbar(right, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=ysb.set)
        ysb.place(in_=self.text, relx=1.0, rely=0, relheight=1.0, x=-16)

        # Datos
        self._topics = self._build_topics()
        self._all_keys = list(self._topics.keys())
        for k in self._all_keys:
            self.list.insert("end", k)

        # Accesos rápidos
        self.bind("<Control-f>", lambda _e: (self.q.focus_set(), self.q.select_range(0, "end")))

        # Selección inicial
        self._open("Barra superior (botones principales)")

    def _build_topics(self) -> dict[str,str]:
        v = "v4.1.0"
        return {
            "Barra superior (botones principales)":
            f"""• Pagos de hoy → muestra el día actual en el panel de pagos.
• Hitos del mes (visual) → nueva ventana redimensionable con calendario a la izquierda y panel de detalle a la derecha. Clic en día = previsualiza. Doble clic = abre el listado principal.
• Listado del mes → muestra todas las filas del mes seleccionado (según filtros).
• Buscar rango… → ventana para consultar entre dos fechas (dd/mm/aaaa).
• Añadir pago… → alta manual. Campos obligatorios: Fecha, Tipo, Detalle. Fondo: FEAGA/FEADER/—.
• Borrar pagos del día… → dos opciones:
    – Borrar manual+web: conserva la Referencia FEAGA (heurística).
    – Borrar TODO: elimina también la Referencia FEAGA.
• Limpiar panel → limpia la tabla y deselecciona el calendario. El botón parpadea tras mostrar resultados para guiar al usuario.
• Mostrar: Manual / Web / Referencia FEAGA → filtros visibles. Si desmarcas todos, el sistema cae a los tres por defecto.
• Regenerar referencia FEAGA → resembrado de ventanas Anticipo (16/10–30/11) y Saldo (01/12–30/06) de la campaña actual (no duplica por índice único).
• Vaciar manual+web → elimina orígenes MANUAL y WEB, preserva Referencia FEAGA. Después resembramos por si acaso; los duplicados se ignoran.
• Vaciar TODO → elimina todo; usa luego “Regenerar referencia FEAGA” si quieres recuperar las ventanas.
• ¿? (Ayuda, {v}) → abre este Centro de Ayuda. Atajo: F1.
""",

            "Calendario anual (interacción)":
            """• Clic en un día → muestra ese día.
• Doble clic en un día → muestra el mes completo.
• Clic derecho → menú contextual (añadir, borrar, ver mes).
• Días con eventos se dibujan con fondo suave; fines de semana en granate.
• Marcadores: si no hay datos del día, se muestran referencias FEAGA/FEADER del día o, en su defecto, del mes.
""",

            "Panel de Pagos (tabla inferior)":
            """• Colores por origen:
    – Verde pálido: Manual
    – Amarillo pálido: Web
    – Azul muy suave: Referencia FEAGA (heurística)
    – Celeste: Información (mensajes informativos)
• La columna “Fuente” incluye URL o nota de origen cuando existe.
• Algunas filas “Del mes · … (original: dd/mm/aaaa)” indican que el sistema trajo información mensual para contextualizar un día sin datos propios.
""",

            "Hitos del mes (visual)":
            """• Ventana normal (minimizar/maximizar/cerrar) y redimensionable.
• Izquierda: rejilla mensual con chips (hasta 3 tipos principales) y contadores FEAGA/FEADER por día.
• Derecha: panel de detalle completo del día (sin recortes).
• Clic = previsualiza a la derecha; doble clic = abre el listado principal en la pestaña Calendario.
• Botón “Exportar mes (CSV)” para extraer todas las filas del mes.
• Leyenda inferior de colores.
""",

            "Actualizar pagos (web)":
            """• Abre una ventana de “Fuentes consultadas”.
• Descarga desde:
    – Notas FEGA (PDF) para anticipo/saldo.
    – Noticias FEGA (titulares que contengan 'anticipo', 'saldo', 'pago', etc.).
    – Fuentes extra que declares en MultiSourceScraper.EXTRA_SOURCES.
• Requiere 'requests' (opcional). Los nuevos registros se insertan con origen WEB, evitando duplicados.
• Al finalizar, muestra cuántas filas se añadieron.
""",

            "Importar Excel / CSV":
            """• Importador genérico (Fecha/Tipo/Fondo/Detalle/Fuente).
• Importador especializado “Aragón” (Mes/Actividad/FEAGA/FEADER): interpreta expresiones tipo
  “del 3 al 15 de mayo”, “a partir del 10”, “12 de junio”, etc., insertando rangos o días.
• Dependencias opcionales: pandas + openpyxl.
• Autocarga silenciosa (si encuentra el Excel de Aragón en rutas predefinidas) una vez por año.
""",

            "Pestaña Índice":
            """• Permite listar por rango de fechas y saltar a un día concreto con “Ir a día”.
• Respeta los colores por origen. Botón “Refrescar” para actualizar tras cambios.
""",

            "Menú contextual (clic derecho en día)":
            """• Añadir pago…
• Borrar manual+web del día
• Borrar TODO del día
• Ver mes
""",

            "Atajos de teclado":
            """• F1 → abrir Centro de Ayuda
• Ctrl+F dentro del Centro de Ayuda → foco al buscador
"""
        }

    def _filter(self, _e=None):
        q = (self.q.get() or "").strip().lower()
        self.list.delete(0, "end")
        for k in self._all_keys:
            if not q or q in k.lower() or q in self._topics[k].lower():
                self.list.insert("end", k)
        if self.list.size() and not self.list.curselection():
            self.list.selection_set(0)

    def _on_select(self, _e=None):
        sel = self.list.curselection()
        if sel:
            self._open(self.list.get(sel[0]))

    def _on_open(self, _e=None):
        self._on_select()

    def _open(self, key: str):
        txt = self._topics.get(key, "")
        self.text.configure(state="normal")
        self.text.delete("1.0", "end")
        self.text.insert("1.0", f"{key}\n", ("h1",))
        self.text.insert("end", "\n")
        self.text.insert("end", txt)
        self.text.tag_configure("h1", font=("Segoe UI", 12, "bold"), foreground="#093a76")
        self.text.configure(state="disabled")

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
        headers={"fecha":(92,"w"),"tipo":(220,"w"),"fondo":(80,"w"),
                 "detalle":(560,"w"),"origen":(140,"center"),"fuente":(220,"w")}
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
            vals=(r.get("fecha",""), r.get("tipo",""), r.get("fondo",""), r.get("detalle",""), label, r.get("fuente",""))
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
        self._blink_job=None; self._blink_cycles=0

        # Auto-cargar Excel Aragón (solo 1ª vez por año)
        self._maybe_autoload_aragon_excel()

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
        tk.Label(title,text="Calendario FEAGA / FEADER – v4.1.0",
                 bg="#f0f7ff",fg="#053e7b",font=("Segoe UI",13,"bold")).pack(side="left",padx=10,pady=8)
        title.pack(fill="x")

        self.console = StatusBar(self)
        self.console.show("info", "Usa el calendario o los botones de arriba para consultar pagos.")

        top=tk.Frame(self,bg="#eef5ff",highlightbackground="#cddcfb",highlightthickness=1)

        # Botones principales (guardamos referencia para tooltips)
        self.btn_hoy = self._color_btn(top,"Pagos de hoy","#2e7d32", lambda: self._show_day(date.today()))
        self.btn_hoy.pack(side="left", padx=(8,4), pady=6)

        self.btn_hitos = self._color_btn(top,"Hitos del mes (visual)","#0d6efd", self._show_month_visual_of_selected)
        self.btn_hitos.pack(side="left", padx=4, pady=6)

        self.btn_listado = ttk.Button(top,text="Listado del mes",command=self._show_month_of_selected)
        self.btn_listado.pack(side="left", padx=6, pady=6)

        self.btn_rango = ttk.Button(top,text="Buscar rango…",command=self._show_range_dialog)
        self.btn_rango.pack(side="left", padx=6, pady=6)

        self.btn_add = ttk.Button(top,text="Añadir pago…",command=self._add_payment_dialog)
        self.btn_add.pack(side="left", padx=(12,4), pady=6)

        self.btn_borrar = ttk.Button(top,text="Borrar pagos del día…",command=self._delete_selected_day_dialog)
        self.btn_borrar.pack(side="left", padx=4, pady=6)

        self.clear_btn = tk.Button(top,text="Limpiar panel",bg="#9c27b0",fg="white",
                                   activebackground="#9c27b0",relief="raised",bd=1,highlightthickness=0,
                                   command=self._clear_view)
        self.clear_btn.configure(font=("Segoe UI",9,"bold"), padx=10, pady=4, cursor="hand2")
        self.clear_btn.pack(side="left", padx=(12,4), pady=6)
        self._clear_btn_bg = self.clear_btn.cget("bg")

        ttk.Label(top,text=" | Mostrar: ").pack(side="left", padx=(12,2))
        self.chk_manual = ttk.Checkbutton(top,text="Manual",variable=self.show_manual,command=self._refresh_current_view)
        self.chk_manual.pack(side="left")
        self.chk_web = ttk.Checkbutton(top,text="Web",variable=self.show_web,command=self._refresh_current_view)
        self.chk_web.pack(side="left")
        self.chk_heur = ttk.Checkbutton(top,text="Referencia FEAGA",variable=self.show_heur,command=self._refresh_current_view)
        self.chk_heur.pack(side="left")

        # Botones de la derecha
        self.btn_regen = ttk.Button(top,text="Regenerar referencia FEAGA",command=self._regen_heuristics)
        self.btn_regen.pack(side="right", padx=6, pady=6)
        self.btn_vaciar_todo = ttk.Button(top,text="Vaciar TODO",command=lambda:self._clear_db(include_heur=True))
        self.btn_vaciar_todo.pack(side="right", padx=(0,8), pady=6)
        self.btn_vaciar_mw = ttk.Button(top,text="Vaciar manual+web",command=lambda:self._clear_db(include_heur=False))
        self.btn_vaciar_mw.pack(side="right", padx=6, pady=6)
        self.btn_help = tk.Button(top,text="¿?", width=3, command=self._open_help)
        self.btn_help.pack(side="right", padx=(0,6), pady=6)

        top.pack(fill="x")

        # Tooltips
        ToolTip(self.btn_hoy, "Ir al día de hoy.")
        ToolTip(self.btn_hitos, "Vista mensual visual con panel de detalle y exportación CSV.")
        ToolTip(self.btn_listado, "Listar todas las filas del mes seleccionado.")
        ToolTip(self.btn_rango, "Buscar por rango de fechas (dd/mm/aaaa).")
        ToolTip(self.btn_add, "Añadir un pago manual.")
        ToolTip(self.btn_borrar, "Borrar pagos del día seleccionado (manual+web o todo).")
        ToolTip(self.clear_btn, "Limpiar la tabla y la selección del calendario.")
        ToolTip(self.chk_manual, "Mostrar/ocultar filas añadidas manualmente.")
        ToolTip(self.chk_web, "Mostrar/ocultar filas descargadas de la web.")
        ToolTip(self.chk_heur, "Mostrar/ocultar la referencia FEAGA (ventanas anticipo/saldo).")
        ToolTip(self.btn_vaciar_mw, "Borra MANUAL+WEB. Mantiene Referencia FEAGA.")
        ToolTip(self.btn_vaciar_todo, "Borra TODO (incluida la Referencia FEAGA).")
        ToolTip(self.btn_regen, "Resembrar la Referencia FEAGA de la campaña actual.")
        ToolTip(self.btn_help, "Abrir Centro de Ayuda (F1).")

        self.tabs=ttk.Notebook(self); self.tabs.pack(fill="both",expand=True,padx=10,pady=10)

        tab_cal=ttk.Frame(self.tabs); self.tabs.add(tab_cal,text="Calendario")
        v_split=ttk.Panedwindow(tab_cal,orient="vertical"); v_split.pack(fill="both",expand=True)

        self.pay_frame=PaymentsInfoFrame(v_split)
        cal_holder=ttk.Frame(v_split)
        v_split.add(cal_holder,weight=3); v_split.add(self.pay_frame,weight=2)

        tools=ttk.Frame(cal_holder); tools.pack(fill="x",pady=(0,4))
        self.btn_actualizar = self._color_btn(tools,"Actualizar pagos (web)","#ff8f00", self._update_from_web)
        self.btn_actualizar.pack(side="left", padx=2, pady=2)
        self.btn_importar = ttk.Button(tools,text="Importar Excel…",command=self._import_excel)
        self.btn_importar.pack(side="left", padx=8, pady=2)
        ttk.Label(tools,text="(orígenes: Web/Manual; FEAGA=genérico, FEADER=según resoluciones)",foreground="#666").pack(side="left", padx=8)
        ToolTip(self.btn_actualizar, "Consultar FEGA y otras fuentes (requiere 'requests').")
        ToolTip(self.btn_importar, "Importar Excel/CSV (genérico o plantilla Aragón).")

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

    # ---------- Parpadeo del botón Limpiar panel ----------
    def _blink_clear_button(self, cycles:int=6, interval:int=180):
        if self._blink_job:
            try: self.after_cancel(self._blink_job)
            except Exception: pass
            self._blink_job=None
        self._blink_cycles = cycles*2
        def _step():
            if self._blink_cycles<=0:
                self.clear_btn.configure(bg=self._clear_btn_bg, activebackground=self._clear_btn_bg)
                self._blink_job=None
                return
            cur=self.clear_btn.cget("bg")
            nxt = "#ffc107" if cur==self._clear_btn_bg else self._clear_btn_bg
            self.clear_btn.configure(bg=nxt, activebackground=nxt)
            self._blink_cycles -= 1
            self._blink_job = self.after(interval, _step)
        _step()

    # ---------- Utilidades UI ----------
    def _clear_view(self):
        if self._blink_job:
            try: self.after_cancel(self._blink_job)
            except Exception: pass
            self._blink_job=None
        self.clear_btn.configure(bg=self._clear_btn_bg, activebackground=self._clear_btn_bg)
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
        self.idx_btn_refresh = ttk.Button(top,text="Refrescar",command=self._refresh_index_tab)
        self.idx_btn_refresh.pack(side="left")
        self.idx_btn_goto = ttk.Button(top,text="Ir a día",command=self._goto_from_index_tab)
        self.idx_btn_goto.pack(side="left",padx=(10,0))
        ToolTip(self.idx_btn_refresh, "Volver a cargar el listado con el rango indicado.")
        ToolTip(self.idx_btn_goto, "Abrir la vista del día seleccionado en la pestaña Calendario.")

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
        rows = self.db.get_day(dt, self._active_origins())

        def has_day_fondo(fondo:str)->bool:
            return any((r.get("fondo","").upper()==fondo and r.get("origen") in ("manual","web")) for r in rows)

        day_has_feaga = has_day_fondo("FEAGA")
        day_has_feader = has_day_fondo("FEADER")

        if not rows:
            gen = FeagaRef.day_in_any_window(dt) or []
            month_rows = self.db.get_month(dt.year, dt.month, self._active_origins())
            feader_month = [r for r in month_rows if (r.get("fondo","").upper()=="FEADER")]
            for r in feader_month[:12]: gen.append(recast_as_month_item(dt, r))
            if not gen: gen = FeagaRef.month_generic_for_day(dt)
            for r in gen:
                if re.match(r"^\d{4}-\d{2}-\d{2}$", r.get("fecha","")):
                    r["fecha"] = fmt_dmy(datetime.strptime(r["fecha"],"%Y-%m-%d").date())
            self.pay_frame.show_rows(f"Día {fmt_dmy(dt)} (referencia)", gen)
            self.console.show("info", f"No hay pagos guardados para {fmt_dmy(dt)}. Se muestran referencias FEAGA y, si hay, FEADER del mes.")
            self._blink_clear_button(); return

        if not day_has_feaga:
            for ref in FeagaRef.day_in_any_window(dt): rows.append(ref)

        month_rows = self.db.get_month(dt.year, dt.month, self._active_origins())

        if not day_has_feaga:
            feaga_month = [r for r in month_rows if r.get("fondo","").upper()=="FEAGA"]
            for r in feaga_month[:20]: rows.append(recast_as_month_item(dt, r))

        if not day_has_feader:
            feader_month = [r for r in month_rows if r.get("fondo","").upper()=="FEADER"]
            for r in feader_month[:20]: rows.append(recast_as_month_item(dt, r))

        for r in rows:
            if re.match(r"^\d{4}-\d{2}-\d{2}$", r.get("fecha","")):
                r["fecha"]=fmt_dmy(datetime.strptime(r["fecha"],"%Y-%m-%d").date())

        self.pay_frame.show_rows(f"Día {fmt_dmy(dt)}", rows)
        self.console.show("ok", f"Mostrando {len(rows)} elemento(s) para {fmt_dmy(dt)} (con caídas a mes si aplican).")
        self._blink_clear_button()

    def _show_month(self, y:int, m:int):
        rows=self.db.get_month(y,m, self._active_origins())
        for r in rows:
            if re.match(r"^\d{4}-\d{2}-\d{2}$", r.get("fecha","")):
                r["fecha"]=fmt_dmy(datetime.strptime(r["fecha"],"%Y-%m-%d").date())
        if not rows:
            self.pay_frame.clear("—")
            self.console.show("warn", f"No hay pagos en {m:02d}/{y} con los filtros actuales.")
        else:
            self.pay_frame.show_rows(f"Mes {m:02d}/{y}", rows)
            self.console.show("ok", f"Mostrando {len(rows)} elemento(s) del mes {m:02d}/{y}.")
        self.tabs.select(0)
        self._blink_clear_button()

    def _show_month_of_selected(self):
        dt=self.yearcal.get_selected_date() or date.today()
        self._show_month(dt.year, dt.month)

    def _show_month_visual_of_selected(self):
        dt = self.yearcal.get_selected_date() or date.today()
        MonthHitosDialog(self, dt.year, dt.month)

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
            for r in rows:
                if re.match(r"^\d{4}-\d{2}-\d{2}$", r.get("fecha","")):
                    r["fecha"]=fmt_dmy(datetime.strptime(r["fecha"],"%Y-%m-%d").date())
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

    # Conservamos el diálogo antiguo por si quieres usarlo en algún sitio
    def _help_clear(self):
        messagebox.showinfo(
            "¿Qué hace 'Vaciar manual+web'?",
            "Elimina SOLO los registros añadidos manualmente o descargados de la web.\n"
            "Mantiene las fechas de la 'Referencia FEAGA' (anticipos/saldos)."
        )

    # Nuevo Centro de Ayuda
    def _open_help(self):
        HelpCenterDialog(self)

    def _refresh_current_view(self):
        if self._current_dt: self._show_day(self._current_dt)

    # -------------------- Importar Excel/CSV --------------------
    @staticmethod
    def _norm(s:str)->str: return re.sub(r"[^a-z]", "", strip_accents_lower(s))

    def _find_col(self, columns, candidates)->str|None:
        cn=[self._norm(x) for x in candidates]
        for c in columns:
            if self._norm(c) in cn: return c
        for c in columns:
            nc=self._norm(c)
            if any(x in nc for x in cn): return c
        return None

    def _ask_year(self, default:int) -> int | None:
        dlg=tk.Toplevel(self); dlg.title("Año destino"); dlg.transient(self.winfo_toplevel()); dlg.grab_set()
        ttk.Label(dlg,text="Año destino para los hitos del Excel:").grid(row=0,column=0,sticky="e",padx=8,pady=8)
        var=tk.IntVar(value=default)
        sp=ttk.Spinbox(dlg,from_=2000,to=2100,width=6,textvariable=var,justify="center")
        sp.grid(row=0,column=1,sticky="w",padx=8,pady=8)
        ans={"ok":False}
        def ok(): ans["ok"]=True; dlg.destroy()
        def cancel(): dlg.destroy()
        btns=ttk.Frame(dlg); btns.grid(row=1,column=0,columnspan=2,pady=(0,8))
        ttk.Button(btns,text="Aceptar",command=ok).pack(side="left",padx=6)
        ttk.Button(btns,text="Cancelar",command=cancel).pack(side="left",padx=6)
        dlg.wait_window(dlg)
        return var.get() if ans["ok"] else None

    def _import_excel(self):
        path=filedialog.askopenfilename(
            title="Selecciona Excel/CSV de pagos",
            filetypes=[("Excel","*.xlsx;*.xls"),("CSV","*.csv"),("Todos","*.*")]
        )
        if not path: return
        try:
            import pandas as pd
        except Exception:
            messagebox.showerror("Excel","Necesitas instalar dependencias:\n\npip install pandas openpyxl")
            return
        try:
            if path.lower().endswith(".csv"):
                df=pd.read_csv(path, sep=None, engine="python")
            else:
                df=pd.read_excel(path)  # requiere openpyxl
        except Exception as ex:
            messagebox.showerror("Excel", f"No se pudo leer el archivo:\n{ex}")
            return

        if df.empty:
            messagebox.showinfo("Excel","El fichero no tiene filas."); return

        # Mapeo especial Aragón
        c_mes   = self._find_col(df.columns, ["mes"])
        c_act   = self._find_col(df.columns, ["actividad","observaciones","descripcion","descripción"])
        c_fea   = self._find_col(df.columns, ["ayuda feaga","feaga"])
        c_fed   = self._find_col(df.columns, ["ayuda feader","feader"])

        if c_mes and c_act and (c_fea or c_fed):
            y = self._ask_year(self.yearcal.current_year)
            if y is None: return
            inserted, skipped, errs = self._import_aragon_calendar_df(df, c_mes, c_act, c_fea, c_fed, y)
            self.yearcal.refresh(); self._refresh_current_view(); self._refresh_index_tab()
            msg=f"Importación (Aragón) completada.\nInsertados: {inserted}\nOmitidos: {skipped}"
            if errs: msg += f"\n\nPrimeros errores:\n" + "\n".join(errs[:10])
            messagebox.showinfo("Excel", msg)
            return

        # Importador genérico (Fecha/Tipo/Fondo/Detalle/Fuente)
        c_fecha  = self._find_col(df.columns, ["fecha","día","dia","date"])
        c_tipo   = self._find_col(df.columns, ["tipo","concepto","pago","descripcion","descripción"])
        c_fondo  = self._find_col(df.columns, ["fondo","feaga","feader","linea","línea"])
        c_det    = self._find_col(df.columns, ["detalle","descripcion","descripción","observaciones","nota","obs"])
        c_fuente = self._find_col(df.columns, ["fuente","source","url","enlace"])
        if not (c_fecha and c_tipo and c_det):
            cols="\n".join(map(str,df.columns))
            messagebox.showerror("Excel",
                "No encuentro columnas mínimas (Fecha, Tipo y Detalle).\n"
                f"Columnas detectadas:\n{cols}")
            return

        inserted=0; skipped=0; errs=[]
        for i,row in df.iterrows():
            try:
                raw=row[c_fecha]
                if isinstance(raw,(datetime,date)): d = raw.date() if isinstance(raw,datetime) else raw
                else:
                    d = parse_ddmmyyyy(str(raw)) or (datetime.fromisoformat(str(raw)).date() if str(raw) else None)
                if not d: skipped+=1; errs.append(f"Fila {i+2}: fecha inválida '{raw}'"); continue
                tipo = str(row[c_tipo]).strip()
                detalle = str(row[c_det]).strip()
                if not tipo or not detalle: skipped+=1; errs.append(f"Fila {i+2}: falta Tipo/Detalle"); continue
                fondo = (str(row[c_fondo]).strip().upper() if c_fondo else "—")
                if "FEADER" in fondo: fondo="FEADER"
                elif "FEAGA" in fondo or "FEGA" in fondo: fondo="FEAGA"
                elif fondo in ("", "NAN"): fondo="—"
                fuente = str(row[c_fuente]).strip() if c_fuente else ""
                self.db.add(d, tipo, fondo, detalle, fuente, origen="manual")
                inserted+=1
            except Exception as ex:
                skipped+=1; errs.append(f"Fila {i+2}: {ex}")

        self.yearcal.refresh(); self._refresh_current_view()
        msg=f"Importación completada.\nInsertados: {inserted}\nOmitidos: {skipped}"
        if errs: msg += f"\n\nPrimeros errores:\n" + "\n".join(errs[:10])
        messagebox.showinfo("Excel", msg)

    # --- Importador especializado: Excel Aragón (Mes/Actividad/FEAGA/FEADER) ---
    def _import_aragon_calendar_df(self, df, c_mes, c_act, c_fea, c_fed, year:int):
        meses_map = {
            "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
            "julio":7,"agosto":8,"septiembre":9,"setiembre":9,"octubre":10,"noviembre":11,"diciembre":12
        }
        def parse_month_name(s: str) -> int|None:
            k = strip_accents_lower(s).strip()
            return meses_map.get(k)

        def end_of_month(y,m):
            return date(y,m,calendar.monthrange(y,m)[1])

        # Patrones en español
        meses_regex = "(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)"
        rx_del_al_de_mes = re.compile(rf"\bdel\s+(\d{{1,2}})\s+al\s+(\d{{1,2}})\s+de\s+{meses_regex}\b", re.I)
        rx_del_al         = re.compile(r"\bdel\s+(\d{1,2})\s+al\s+(\d{1,2})\b", re.I)
        rx_a_partir_de_mes= re.compile(rf"\ba\s+partir\s+del\s+(\d{{1,2}})\s+de\s+{meses_regex}\b", re.I)
        rx_a_partir       = re.compile(r"\ba\s+partir\s+del\s+(\d{1,2})\b", re.I)
        rx_dia_de_mes     = re.compile(rf"\b(\d{{1,2}})\s+de\s+{meses_regex}\b", re.I)
        rx_dia_suelto     = re.compile(r"\b(\d{1,2})\b")

        def clean_tipo_det(line: str) -> tuple[str,str]:
            txt = line.strip().lstrip("-").strip()
            if ":" in txt:
                left, right = txt.split(":",1)
                tipo = right.strip() or txt[:80]
                return (tipo, txt)
            return (txt[:80], txt)

        def yes_col(val) -> bool:
            t = strip_accents_lower(val).strip()
            return t.startswith("si") or t.startswith("sí")

        inserted=0; skipped=0; errs=[]
        for i,row in df.iterrows():
            try:
                mes_txt = row[c_mes]
                m = parse_month_name(mes_txt)
                if not m:
                    skipped+=1; errs.append(f"Fila {i+2}: mes inválido '{mes_txt}'"); continue
                act_text = str(row[c_act] or "").strip()
                if not act_text:
                    skipped+=1; errs.append(f"Fila {i+2}: actividad vacía"); continue
                is_feaga = yes_col(row[c_fea]) if c_fea else False
                is_feader= yes_col(row[c_fed]) if c_fed else False

                if not (is_feaga or is_feader):
                    d = date(year, m, 1)
                    tipo, detalle = clean_tipo_det(act_text.splitlines()[0])
                    self.db.add(d, tipo, "—", f"[Mes] {detalle}", "Excel Aragón (Calendario)", origen="manual")
                    inserted += 1
                    continue

                lines = [ln.strip() for ln in act_text.splitlines() if ln.strip()] or [act_text]

                for ln in lines:
                    ln_clean = strip_accents_lower(ln)

                    m1 = rx_del_al_de_mes.search(ln_clean)
                    if m1:
                        d1, d2, mesname = int(m1.group(1)), int(m1.group(2)), m1.group(3)
                        mm = parse_month_name(mesname) or m
                        start = date(year, mm, d1); end = date(year, mm, d2)
                        tipo, detalle = clean_tipo_det(ln)
                        if is_feaga: self.db.add_range(start, end, tipo=tipo, fondo="FEAGA", detalle=detalle, fuente="Excel Aragón (Calendario)", origen="manual")
                        if is_feader: self.db.add_range(start, end, tipo=tipo, fondo="FEADER", detalle=detalle, fuente="Excel Aragón (Calendario)", origen="manual")
                        inserted += (d2 - d1 + 1); continue

                    m2 = rx_del_al.search(ln_clean)
                    if m2:
                        d1, d2 = int(m2.group(1)), int(m2.group(2))
                        start = date(year, m, d1); end = date(year, m, d2)
                        tipo, detalle = clean_tipo_det(ln)
                        if is_feaga: self.db.add_range(start, end, tipo=tipo, fondo="FEAGA", detalle=detalle, fuente="Excel Aragón (Calendario)", origen="manual")
                        if is_feader: self.db.add_range(start, end, tipo=tipo, fondo="FEADER", detalle=detalle, fuente="Excel Aragón (Calendario)", origen="manual")
                        inserted += (d2 - d1 + 1); continue

                    m3 = rx_a_partir_de_mes.search(ln_clean)
                    if m3:
                        d1, mesname = int(m3.group(1)), m3.group(2)
                        mm = parse_month_name(mesname) or m
                        start = date(year, mm, d1); end = end_of_month(year, mm)
                        tipo, detalle = clean_tipo_det(ln)
                        if is_feaga: self.db.add_range(start, end, tipo=tipo, fondo="FEAGA", detalle=detalle, fuente="Excel Aragón (Calendario)", origen="manual")
                        if is_feader: self.db.add_range(start, end, tipo=tipo, fondo="FEADER", detalle=detalle, fuente="Excel Aragón (Calendario)", origen="manual")
                        inserted += (end - start).days + 1; continue

                    m4 = rx_a_partir.search(ln_clean)
                    if m4:
                        d1 = int(m4.group(1))
                        start = date(year, m, d1); end = end_of_month(year, m)
                        tipo, detalle = clean_tipo_det(ln)
                        if is_feaga: self.db.add_range(start, end, tipo=tipo, fondo="FEAGA", detalle=detalle, fuente="Excel Aragón (Calendario)", origen="manual")
                        if is_feader: self.db.add_range(start, end, tipo=tipo, fondo="FEADER", detalle=detalle, fuente="Excel Aragón (Calendario)", origen="manual")
                        inserted += (end - start).days + 1; continue

                    m5 = rx_dia_de_mes.search(ln_clean)
                    if m5:
                        d1, mesname = int(m5.group(1)), m5.group(2)
                        mm = parse_month_name(mesname) or m
                        theday = date(year, mm, d1)
                        tipo, detalle = clean_tipo_det(ln)
                        if is_feaga: self.db.add(theday, tipo, "FEAGA", detalle, "Excel Aragón (Calendario)", origen="manual")
                        if is_feader: self.db.add(theday, tipo, "FEADER", detalle, "Excel Aragón (Calendario)", origen="manual")
                        inserted += 1; continue

                    m6 = rx_dia_suelto.search(ln_clean)
                    if m6 and ":" in ln:
                        try:
                            d1 = int(m6.group(1))
                            theday = date(year, m, d1)
                            tipo, detalle = clean_tipo_det(ln)
                            if is_feaga: self.db.add(theday, tipo, "FEAGA", detalle, "Excel Aragón (Calendario)", origen="manual")
                            if is_feader: self.db.add(theday, tipo, "FEADER", detalle, "Excel Aragón (Calendario)", origen="manual")
                            inserted += 1; continue
                        except Exception:
                            pass

                    theday = date(year, m, 1)
                    tipo, detalle = clean_tipo_det(ln)
                    if is_feaga: self.db.add(theday, tipo, "FEAGA", f"[Mes] {detalle}", "Excel Aragón (Calendario)", origen="manual")
                    if is_feader: self.db.add(theday, tipo, "FEADER", f"[Mes] {detalle}", "Excel Aragón (Calendario)", origen="manual")
                    inserted += 1

            except Exception as ex:
                skipped+=1; errs.append(f"Fila {i+2}: {ex}")

        return inserted, skipped, errs

    # ---- Auto-carga al arrancar (silenciosa, 1 vez por año) ----
    def _maybe_autoload_aragon_excel(self):
        year = self.yearcal.current_year
        key = f"autoload_aragon_{year}"
        if self.db.get_meta(key): return  # ya importado para este año

        xls_path = None
        for p in [AUTOLOAD_ARAGON_EXCEL, *AUTOLOAD_FALLBACKS]:
            if p and Path(p).exists(): xls_path = Path(p); break
        if not xls_path:
            self.console.show("info", "Autocarga Aragón: fichero no encontrado (se ignora)."); return

        try:
            import pandas as pd
        except Exception:
            self.console.show("warn", "Autocarga Aragón: faltan dependencias (pandas/openpyxl). Usa 'Importar Excel…'."); return

        try:
            df = pd.read_excel(xls_path)
        except Exception as ex:
            self.console.show("warn", f"Autocarga Aragón: no se pudo leer el Excel ({ex})."); return

        c_mes   = self._find_col(df.columns, ["mes"])
        c_act   = self._find_col(df.columns, ["actividad","observaciones","descripcion","descripción"])
        c_fea   = self._find_col(df.columns, ["ayuda feaga","feaga"])
        c_fed   = self._find_col(df.columns, ["ayuda feader","feader"])
        if not (c_mes and c_act and (c_fea or c_fed)):
            self.console.show("warn", "Autocarga Aragón: formato no reconocido (Mes/Actividad/FEAGA/FEADER)."); return

        inserted, skipped, errs = self._import_aragon_calendar_df(df, c_mes, c_act, c_fea, c_fed, year)
        self.db.set_meta(key, "ok")
        self.yearcal.refresh(); self._refresh_current_view(); self._refresh_index_tab()
        msg = f"Autocarga Aragón ({year}): insertados {inserted}, omitidos {skipped}."
        if errs: msg += " (Se omitieron algunas filas por formato.)"
        self.console.show("ok", msg)

    # Web (fuentes + hilo)
    def _update_from_web(self):
        # Fuentes a mostrar
        sources = []
        for etiqueta, url in FegaWebScraper.NOTE_URLS:
            sources.append((f"FEGA – {etiqueta}", url))
        sources.append(("FEGA – Noticias", "https://www.fega.gob.es/es/noticias"))
        for name, url in MultiSourceScraper.EXTRA_SOURCES:
            sources.append((name, url))

        win = WebSourcesDialog(self, sources)
        win.mark_running()

        def run():
            try:
                n0 = self.db.count_rows()
                sc=MultiSourceScraper()
                if not sc.available(): raise RuntimeError("Falta 'requests'.")
                sc.fetch_into_db(self.db, year_hint=date.today().year)
                n1 = self.db.count_rows()
                added = max(0, n1-n0)
                self.after(0, lambda:(self.yearcal.refresh(),
                                      self._refresh_current_view(),
                                      self._refresh_index_tab(),
                                      win.mark_done(added),
                                      self.console.show("ok","Pagos actualizados desde múltiples fuentes web.")))
            except Exception as ex:
                self.after(0, lambda: (win.mark_done(0),
                                       self.console.show("error", f"No se pudo completar la descarga: {ex}"),
                                       messagebox.showerror("Web", f"No se pudo completar la descarga:\n{ex}")))
        threading.Thread(target=run,daemon=True).start()
        self.console.show("warn","Buscando pagos/ventanas en FEGA (PDF, noticias) y fuentes extra…")

# -------------------- main --------------------
def main():
    root=tk.Tk()
    root.title("Calendario FEAGA/FEADER – v4.1.0")
    root.geometry("1280x820")
    try:
        from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)
    except Exception: pass
    app=App(root); app.pack(fill="both",expand=True)
    # Atajo universal de ayuda
    root.bind_all("<F1>", lambda e: app._open_help())
    root.mainloop()

if __name__=="__main__":
    main()
