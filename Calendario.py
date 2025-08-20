# -*- coding: utf-8 -*-
"""
Editor Calendario FEAGA/FEADER (v0.8)
- Calendario 12 meses desplazable, fines de semana marcados y días con pagos resaltados.
- Panel de "Pagos en la fecha seleccionada" redimensionable.
- Botón "Actualizar pagos (web)" también en el panel de Meses + "Pagos de hoy".
- Aviso automático si HOY hay pagos vigentes (barra amarilla con acceso directo).
- Scraping FEGA (web) y lectura de PDFs de circulares FEGA para poblar pagos (opcionales).
- Heurística FEAGA integrada (anticipos y saldos) para marcar ventanas.

Dependencias opcionales para enriquecer:
  pip install requests beautifulsoup4 PyPDF2

Requisitos base:
  pip install pandas xlsxwriter
"""

import calendar
import re
from datetime import date, datetime, timedelta
import threading
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

# =========================
#  Utilidades
# =========================
SPANISH_MONTHS = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10,
    "noviembre": 11, "diciembre": 12
}

def parse_spanish_date(text, default_year=None):
    text = text.strip().lower()
    m = re.search(r"\b(\d{1,2})\s*(?:de\s*)?(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)(?:\s*de\s*(\d{4}))?\b", text)
    if m:
        d, mes, y = int(m.group(1)), SPANISH_MONTHS[m.group(2)], (int(m.group(3)) if m.group(3) else default_year)
        if y is None:
            y = date.today().year
        return (y, mes, d)
    m = re.search(r"\b(\d{1,2})[/-](\d{1,2})(?:[/-](\d{2,4}))?\b", text)
    if m:
        d, mes = int(m.group(1)), int(m.group(2))
        y = int(m.group(3)) if m.group(3) else (default_year or date.today().year)
        if y < 100:
            y += 2000
        return (y, mes, d)
    return None

def daterange(d1: date, d2: date):
    if d2 < d1:
        return
    cur = d1
    while cur <= d2:
        yield cur
        cur = cur + timedelta(days=1)

# ===========================================================
#  Modelo de pagos
# ===========================================================
class PaymentsIndex:
    """Índice de pagos por fecha -> lista de dicts con: tipo, fondo, detalle, fuente"""
    def __init__(self):
        self._by_date = {}

    def clear(self):
        self._by_date.clear()

    def add(self, dt: date, tipo: str, fondo: str, detalle: str, fuente: str = ""):
        self._by_date.setdefault(dt, []).append({
            "tipo": tipo, "fondo": fondo, "detalle": detalle, "fuente": fuente
        })

    def add_range(self, d1: date, d2: date, **kwargs):
        for dt in daterange(d1, d2):
            self.add(dt, **kwargs)

    def get_day(self, dt: date):
        return list(self._by_date.get(dt, []))

    def has_day(self, dt: date) -> bool:
        return dt in self._by_date

# ===========================================================
#  Heurística FEGA (anticipos/saldos)
# ===========================================================
class HeuristicPaymentsProvider:
    def get_for_day(self, dt: date):
        y, m, _ = dt.year, dt.month, dt.day
        campaign_year = y if m >= 10 else y - 1

        start_ant = date(campaign_year, 10, 16)
        end_ant   = date(campaign_year, 11, 30)
        start_sal = date(campaign_year, 12, 1)
        end_sal   = date(campaign_year + 1, 6, 30)

        out = []
        if start_ant <= dt <= end_ant:
            out.append({"tipo":"Anticipo ayudas directas","fondo":"FEAGA",
                        "detalle":f"Hasta el 70% campaña {campaign_year}. Ventana general: 16/10–30/11.",
                        "fuente":"Heurística FEGA"})
        if (start_sal <= dt <= date(campaign_year,12,31)) or (date(campaign_year+1,1,1) <= dt <= end_sal):
            out.append({"tipo":"Saldo ayudas directas","fondo":"FEAGA",
                        "detalle":f"Hasta el 30% restante campaña {campaign_year}. Ventana general: 01/12–30/06(año+1).",
                        "fuente":"Heurística FEGA"})
        # No añadimos FEADER genérico aquí para no “ruido” en avisos de hoy.
        if not any(i["fondo"] == "FEAGA" for i in out):
            out.append({"tipo":"Sin pagos FEAGA generales","fondo":"—",
                        "detalle":"Fuera de ventanas generales de anticipos/saldos. Revise resoluciones específicas.",
                        "fuente":"Heurística FEGA"})
        return out

    def get_ranges_for_campaign(self, campaign_year: int):
        start_ant = date(campaign_year, 10, 16)
        end_ant   = date(campaign_year, 11, 30)
        start_sal = date(campaign_year, 12, 1)
        end_sal   = date(campaign_year + 1, 6, 30)
        return (start_ant, end_ant, start_sal, end_sal)

# ===========================================================
#  Scraper FEGA (web) y lector de PDFs (opcionales)
# ===========================================================
class FegaWebScraper:
    NOTE_URLS = [
        "https://www.fega.gob.es/sites/default/files/files/document/Nota_web_Ecorregimenes_Ca_2024_ANTICIPO.pdf",
        "https://www.fega.gob.es/sites/default/files/files/document/Nota_Web_AAS_Ca_2024_ANTICIPO.pdf",
        "https://www.fega.gob.es/sites/default/files/files/document/241115_NOTA_WEB_EERR_PRIMER_SALDO_Ca_2024_def.pdf",
    ]
    def __init__(self):
        try:
            import requests  # noqa
        except Exception:
            self._available = False
        else:
            self._available = True
    def available(self): return self._available

    def fetch_into_index(self, index: PaymentsIndex, year_hint: int | None = None):
        if not self._available:
            raise RuntimeError("requests no está instalado. Instala 'requests' para scraping.")
        import requests
        for url in self.NOTE_URLS:
            try:
                r = requests.get(url, timeout=15)
                if r.status_code != 200:
                    continue
                text = r.content.decode("latin-1", errors="ignore")
            except Exception:
                text = url
            y = year_hint
            m = re.search(r"(20\d{2})", text)
            if m:
                y = int(m.group(1))
            if not y:
                y = date.today().year

            ant1 = parse_spanish_date("16 de octubre", y)
            ant2 = parse_spanish_date("30 de noviembre", y)
            sal1 = parse_spanish_date("1 de diciembre", y)
            sal2 = parse_spanish_date("30 de junio", y + 1)

            if ant1 and ant2:
                index.add_range(date(*ant1), date(*ant2),
                                tipo="Anticipo ayudas directas", fondo="FEAGA",
                                detalle=f"Ventana general de anticipos (nota FEGA).",
                                fuente=url)
            if sal1 and sal2:
                index.add_range(date(*sal1), date(*sal2),
                                tipo="Saldo ayudas directas", fondo="FEAGA",
                                detalle=f"Ventana general de saldos (nota FEGA).",
                                fuente=url)
        return index

class FegaPDFIngestor:
    def __init__(self):
        try:
            import PyPDF2  # noqa
        except Exception:
            self._available = False
        else:
            self._available = True
    def available(self): return self._available

    def _extract_text(self, path: Path) -> str:
        import PyPDF2
        txt = []
        try:
            with open(path, "rb") as fh:
                reader = PyPDF2.PdfReader(fh)
                for page in reader.pages:
                    try:
                        txt.append(page.extract_text() or "")
                    except Exception:
                        pass
        except Exception:
            return ""
        return "\n".join(txt)

    def ingest_folder(self, folder: Path, index: PaymentsIndex, default_year: int | None = None):
        if not self._available:
            raise RuntimeError("PyPDF2 no está instalado. Instala 'PyPDF2'.")
        if not folder or not Path(folder).exists():
            raise FileNotFoundError(str(folder))

        for pdf in Path(folder).glob("*.pdf"):
            text = self._extract_text(pdf).lower()
            if not text:
                continue

            y = default_year
            m = re.search(r"(20\d{2})", text)
            if m:
                y = int(m.group(1))
            if not y:
                y = date.today().year

            for mo in re.finditer(r"del\s+(\d{1,2})\s+al\s+(\d{1,2})\s+de\s+(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)", text):
                d1, d2, mes = int(mo.group(1)), int(mo.group(2)), SPANISH_MONTHS[mo.group(3)]
                contexto = text[max(0, mo.start()-120):mo.end()+120]
                etiqueta, fondo = "Pago/ventana", "—"
                if "anticipo" in contexto:
                    etiqueta, fondo = "Anticipo ayudas directas", "FEAGA"
                elif "saldo" in contexto:
                    etiqueta, fondo = "Saldo ayudas directas", "FEAGA"
                elif "feader" in contexto:
                    etiqueta, fondo = "Pago medidas desarrollo rural", "FEADER"
                index.add_range(date(y, mes, d1), date(y, mes, d2),
                                tipo=etiqueta, fondo=fondo,
                                detalle=f"Ventana detectada en circular {pdf.name}", fuente=str(pdf))

            for mo in re.finditer(r"\b(\d{1,2})\s*(?:de\s*)?(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|octubre|noviembre|diciembre)\b", text):
                d, mes = int(mo.group(1)), SPANISH_MONTHS[mo.group(2)]
                contexto = text[max(0, mo.start()-140):mo.end()+140]
                etiqueta, fondo = None, "—"
                if "anticipo" in contexto:
                    etiqueta, fondo = "Anticipo ayudas directas", "FEAGA"
                elif "saldo" in contexto:
                    etiqueta, fondo = "Saldo ayudas directas", "FEAGA"
                elif "pago" in contexto and "feader" in contexto:
                    etiqueta, fondo = "Pago medidas desarrollo rural", "FEADER"
                if etiqueta:
                    index.add(date(y, mes, d), tipo=etiqueta, fondo=fondo,
                              detalle=f"Fecha mencionada en {pdf.name}", fuente=str(pdf))

# ===========================================================
#  Contenedor desplazable vertical (Canvas)
# ===========================================================
class VerticalScrolledFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.inner = ttk.Frame(self.canvas)
        self._win = self.canvas.create_window((0, 0), window=self.inner, anchor="nw")

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.inner.bind("<Configure>", self._on_inner_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        self.canvas.bind_all("<Button-4>", self._on_mousewheel)
        self.canvas.bind_all("<Button-5>", self._on_mousewheel)

    def _on_inner_configure(self, _):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        self.canvas.itemconfigure(self._win, width=self.canvas.winfo_width())

    def _on_canvas_configure(self, e):
        self.canvas.itemconfigure(self._win, width=e.width)

    def _on_mousewheel(self, event):
        try:
            if event.num == 4:
                self.canvas.yview_scroll(-3, "units")
            elif event.num == 5:
                self.canvas.yview_scroll(3, "units")
            else:
                self.canvas.yview_scroll(int(-event.delta/40), "units")
        except Exception:
            pass

# ===========================================================
#  Calendario anual
# ===========================================================
class YearCalendarFrame(ttk.Frame):
    MESES_ES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]
    DIAS_CORTOS_ES = ["L","M","X","J","V","S","D"]

    def __init__(self, master, year=None, on_date_click=None, on_date_double_click=None, has_events_predicate=None):
        super().__init__(master)
        calendar.setfirstweekday(calendar.MONDAY)
        self.current_year = year or date.today().year
        self.on_date_click = on_date_click
        self.on_date_double_click = on_date_double_click
        self.has_events_predicate = has_events_predicate or (lambda y,m,d: False)

        self._init_styles()
        self._build_controls()

        self.scroll = VerticalScrolledFrame(self)
        self.scroll.grid(row=1, column=0, sticky="nsew")
        self.rowconfigure(1, weight=1)
        self.columnconfigure(0, weight=1)

        self._build_year_grid()

    def _init_styles(self):
        st = ttk.Style(self)
        st.configure("CalDay.TButton", padding=(1,1), font=("Segoe UI", 9))
        st.configure("CalWeekend.TButton", padding=(1,1), font=("Segoe UI", 9), foreground="#b00000")
        st.configure("CalEvent.TButton", padding=(1,1), font=("Segoe UI", 9, "bold"), foreground="#084f8a")
        st.configure("CalEventWE.TButton", padding=(1,1), font=("Segoe UI", 9, "bold"), foreground="#b00000")
        st.configure("CalHead.TLabel", padding=(0,0), font=("Segoe UI", 9))
        st.configure("CalHeadWE.TLabel", padding=(0,0), font=("Segoe UI", 9), foreground="#b00000")

    def _build_controls(self):
        top = ttk.Frame(self)
        top.grid(row=0, column=0, sticky="ew", pady=(0,4))
        top.columnconfigure(5, weight=1)
        ttk.Button(top, text="<<", width=3, command=self._prev_year).grid(row=0, column=0, padx=(0,4))
        self.year_var = tk.IntVar(value=self.current_year)
        spin = ttk.Spinbox(top, from_=1900, to=2100, width=6, textvariable=self.year_var, justify="center", command=self._spin_changed)
        spin.grid(row=0, column=1)
        spin.bind("<Return>", lambda e: self._set_year(self.year_var.get()))
        spin.bind("<FocusOut>", lambda e: self._set_year(self.year_var.get()))
        ttk.Button(top, text=">>", width=3, command=self._next_year).grid(row=0, column=2, padx=(4,8))
        ttk.Button(top, text="Hoy", command=self._go_today).grid(row=0, column=3)
        self.info_label = ttk.Label(top, text="(clic=pagos · doble clic=insertar en Actividad)")
        self.info_label.grid(row=0, column=5, sticky="e")

    def refresh(self):
        self._build_year_grid()

    def _build_year_grid(self):
        for ch in self.scroll.inner.winfo_children():
            ch.destroy()
        grid = ttk.Frame(self.scroll.inner)
        grid.grid(row=0, column=0, sticky="nsew")
        for r in range(4): grid.rowconfigure(r, weight=1, uniform="months")
        for c in range(3): grid.columnconfigure(c, weight=1, uniform="months")
        for month in range(1, 12+1):
            r, c = (month-1)//3, (month-1)%3
            mf = self._build_month(grid, self.current_year, month)
            mf.grid(row=r, column=c, padx=4, pady=4, sticky="nsew")
        self.scroll.inner.rowconfigure(0, weight=1); self.scroll.inner.columnconfigure(0, weight=1)

    def _build_month(self, parent, year, month):
        f = ttk.Frame(parent, borderwidth=1, relief="solid", padding=(4,3,4,4))
        ttk.Label(f, text=self.MESES_ES[month-1], anchor="center", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, columnspan=7, sticky="ew", pady=(0,2))
        for i, d in enumerate(self.DIAS_CORTOS_ES):
            style = "CalHeadWE.TLabel" if i in (5,6) else "CalHead.TLabel"
            ttk.Label(f, text=d, style=style, anchor="center").grid(row=1, column=i, padx=1, pady=0, sticky="nsew")
        weeks = calendar.monthcalendar(year, month)
        for r, week in enumerate(weeks, start=2):
            for c in range(7):
                day = week[c]
                if day == 0:
                    ttk.Label(f, text="", anchor="center").grid(row=r, column=c, padx=1, pady=1, sticky="nsew"); continue
                has = self.has_events_predicate(year, month, day)
                if c in (5,6):
                    style = "CalEventWE.TButton" if has else "CalWeekend.TButton"
                else:
                    style = "CalEvent.TButton" if has else "CalDay.TButton"
                btn = ttk.Button(f, text=str(day), style=style)
                btn.grid(row=r, column=c, padx=1, pady=1, sticky="nsew")
                btn.bind("<Button-1>", lambda e, y=year, m=month, d=day: self._on_day_click(y, m, d))
                btn.bind("<Double-Button-1>", lambda e, y=year, m=month, d=day: self._on_day_dblclick(y, m, d))
        for c in range(7): f.columnconfigure(c, weight=1, uniform="days")
        f.grid_rowconfigure(0, minsize=18); f.grid_rowconfigure(1, minsize=18)
        for r in range(2, 2+len(weeks)): f.grid_rowconfigure(r, weight=1, minsize=26)
        return f

    # API pública
    def go_to_date(self, dt: date):
        self._set_year(dt.year)
        self._on_day_click(dt.year, dt.month, dt.day)

    # eventos
    def _on_day_click(self, y,m,d):
        self.info_label.config(text=f"Seleccionado: {d:02d}/{m:02d}/{y}  (clic=pagos · doble clic=insertar)")
        if callable(self.on_date_click): self.on_date_click(y,m,d)
    def _on_day_dblclick(self, y,m,d):
        self._on_day_click(y,m,d)
        if callable(self.on_date_double_click): self.on_date_double_click(y,m,d)

    # navegación
    def _set_year(self, year):
        try: year = int(year)
        except: year = self.current_year
        year = max(1900, min(2100, year))
        self.current_year = year; self.year_var.set(year); self._build_year_grid()
    def _prev_year(self): self._set_year(self.current_year - 1)
    def _next_year(self): self._set_year(self.current_year + 1)
    def _go_today(self):
        t = date.today(); self.go_to_date(t)
    def _spin_changed(self): self._set_year(self.year_var.get())

# ===========================================================
#  Panel pagos (Treeview)
# ===========================================================
class PaymentsInfoFrame(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=(6,4,6,4))
        top = ttk.Frame(self); top.grid(row=0, column=0, sticky="ew")
        ttk.Label(top, text="Pagos en la fecha seleccionada", font=("Segoe UI", 10, "bold")).pack(side="left")
        self.date_lbl = ttk.Label(top, text="—"); self.date_lbl.pack(side="right")

        cols=("tipo","fondo","detalle","fuente")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for c, (w, st) in {"tipo":(180,"w"), "fondo":(80,"w"), "detalle":(600,"w"), "fuente":(220,"w")}.items():
            self.tree.heading(c, text=c.capitalize()); self.tree.column(c, width=w, anchor=st, stretch=True)
        self.tree.grid(row=1, column=0, sticky="nsew", pady=(6,0))
        ysb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        xsb = ttk.Scrollbar(self, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        ysb.grid(row=1, column=1, sticky="ns"); xsb.grid(row=2, column=0, sticky="ew")

        self.grid_rowconfigure(1, weight=1); self.grid_columnconfigure(0, weight=1)
        ttk.Label(self, text="Nota: FEAGA 16/10–30/11 (anticipos), 01/12–30/06 (saldos). FEADER según resoluciones.",
                  foreground="#444").grid(row=3, column=0, columnspan=2, sticky="w", pady=(6,0))

    def show(self, dt: date, items: list[dict]):
        self.date_lbl.config(text=dt.strftime("%d/%m/%Y"))
        self.tree.delete(*self.tree.get_children())
        for it in items:
            self.tree.insert("", "end", values=(it["tipo"], it["fondo"], it["detalle"], it.get("fuente","")))

# ===========================================================
#  App principal
# ===========================================================
class CalendarioFEAGA_FEADERFrame(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.heur = HeuristicPaymentsProvider()
        self.index = PaymentsIndex()
        self._build_dataset()
        self._build_ui()
        self._load_into_tree(self.df)

        # Pre-carga heurística para el año visible (marcas en calendario)
        self._populate_index_from_heuristics_for_year(date.today().year)
        # Aviso inicial de pagos vigentes hoy
        self._update_today_banner()

    # -------- Datos base tabla ----------
    def _build_dataset(self):
        data = {
            "Mes": [
                "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre",
                "Diciembre", "Enero - Junio (año siguiente)", "Todo el año"
            ],
            "Actividad": [
                "- Preparación de la campaña PAC.\n- Difusión de novedades y requisitos del PEPAC 2023-2027.\n- Asesoramiento a agricultores.",
                "- 1 de febrero: Apertura del período de presentación de la Solicitud Única de la PAC en Aragón.\n- Inicio de la recepción de solicitudes.\n- Asistencia técnica y asesoramiento personalizado.",
                "- Continuación de la presentación de la Solicitud Única.\n- Actualización de registros y documentación.\n- Resolución de dudas y consultas de los solicitantes.",
                "- Continuación de la presentación de la Solicitud Única.\n- Revisión y verificación de datos declarados.\n- Preparación para el cierre del plazo de solicitudes.",
                "- 15 de mayo: Fecha límite para la presentación de la Solicitud Única sin penalización.\n- Del 16 al 31 de mayo: Presentación de solicitudes con penalización del 1% por día hábil de retraso.\n- 31 de mayo: Último día para realizar modificaciones de la Solicitud Única sin penalización.\n- Asistencia en correcciones y ajustes de las solicitudes presentadas.",
                "- Del 1 al 9 de junio: Presentación de solicitudes con penalización incrementada.\n- 9 de junio: Último día para presentar solicitudes con penalización (máximo 25%).\n- Inicio de controles administrativos y cruzados de las solicitudes.\n- Comunicación de posibles incidencias a los solicitantes.",
                "- Inicio de controles sobre el terreno en explotaciones agrícolas y ganaderas.\n- Verificación del cumplimiento de los requisitos y compromisos del PEPAC.\n- Gestión de incidencias detectadas en controles administrativos.",
                "- Continuación de los controles sobre el terreno.\n- Procesamiento de resultados de controles.\n- Preparación de informes y comunicaciones a los agricultores.",
                "- Notificación de resultados de controles a los beneficiarios.\n- Periodo para presentar alegaciones y documentación adicional.\n- Ajustes finales en los expedientes antes de la resolución.",
                "- A partir del 16 de octubre: Inicio de pagos de anticipos de las ayudas directas (hasta el 70%).\n- Publicación de resoluciones provisionales de ayudas.\n- Inicio de pagos de ciertas medidas FEADER que lo permitan.",
                "- Continuación de los pagos de anticipos.\n- Actualización y cierre de expedientes administrativos.\n- Preparación de resoluciones definitivas.",
                "- 31 de diciembre: Fecha límite para realizar ciertos pagos nacionales.\n- Finalización de trámites administrativos pendientes.\n- Planificación de la próxima campaña PAC.",
                "- Pagos finales de las ayudas directas hasta el 30 de junio del año siguiente al de la solicitud.\n- Resolución de incidencias y recursos.\n- Desarrollo y ejecución de proyectos FEADER aprobados en convocatorias anteriores.",
                "- Desarrollo y ejecución de proyectos FEADER según la programación de Aragón en el PEPAC 2023-2027.\n- Convocatorias específicas de medidas de desarrollo rural (modernización, inversiones, agroambiente, clima, LEADER, etc.).\n- Asesoramiento y formación a agricultores y ganaderos sobre prácticas sostenibles y requisitos normativos.\n- Seguimiento y evaluación de proyectos en curso."
            ],
            "Ayuda FEAGA": [
                "No", "Sí (Solicitud Única)", "Sí", "Sí", "Sí (Fecha límite Solicitud Única)",
                "Sí", "Sí", "Sí", "Sí", "Sí (Pagos anticipados de ayudas directas)",
                "Sí", "Sí (Continuación de pagos)", "Sí (Pagos finales de ayudas directas)", "No"
            ],
            "Ayuda FEADER": [
                "No", "Sí (algunas medidas FEADER incluidas en la Solicitud Única)", "Sí", "Sí",
                "Sí", "Sí", "Sí (Controles en medidas FEADER)", "Sí", "Sí",
                "Sí (En medidas que lo contemplen)", "Sí", "Sí (Proyectos con ejecución anual)",
                "Sí (Dependiendo del proyecto y convocatoria)", "Sí (Medidas de desarrollo rural y convocatorias específicas)"
            ]
        }
        self.df = pd.DataFrame(data)

    # -------- UI ----------
    def _build_ui(self):
        self.pack(fill="both", expand=True)

        # Barra de aviso (arranca oculta)
        self.alert_bar = tk.Frame(self, bg="#fff3cd", highlightbackground="#ffeeba", highlightthickness=1)
        self.alert_bar.pack(fill="x", padx=10, pady=(10,0))
        self.alert_msg = tk.Label(self.alert_bar, text="", bg="#fff3cd", fg="#856404", font=("Segoe UI", 9, "bold"))
        self.alert_msg.pack(side="left", padx=8, pady=4)
        tk.Button(self.alert_bar, text="Ver hoy", command=self._open_today, relief="groove").pack(side="right", padx=6, pady=4)
        tk.Button(self.alert_bar, text="X", command=lambda: self.alert_bar.pack_forget(), relief="flat", bg="#fff3cd").pack(side="right", padx=(0,6))
        self.alert_bar.pack_forget()  # oculta de inicio

        ttk.Label(self, text="Calendario FEAGA / FEADER – Editor y pagos (v0.8)",
                  font=("Segoe UI", 13, "bold")).pack(padx=10, pady=(6,5), anchor="w")

        split = ttk.Panedwindow(self, orient="horizontal")
        split.pack(fill="both", expand=True, padx=10, pady=10)

        # ---- Tabla izquierda + barra de herramientas
        left = ttk.Frame(split)
        split.add(left, weight=3)

        toolbar = ttk.Frame(left)
        toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0,4))
        ttk.Button(toolbar, text="Actualizar pagos (web)", command=self._update_from_web).pack(side="left")
        ttk.Button(toolbar, text="Pagos de hoy", command=self._open_today).pack(side="left", padx=(6,0))

        self.columns = ["Mes","Actividad","Ayuda FEAGA","Ayuda FEADER"]
        self.tree = ttk.Treeview(left, columns=self.columns, show="headings", selectmode="browse")
        for col in self.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=(150 if col!="Actividad" else 450), anchor="w", stretch=True)
        self.tree.grid(row=1, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(left, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(left, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        yscroll.grid(row=1, column=1, sticky="ns")
        xscroll.grid(row=2, column=0, sticky="ew")

        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        # ---- Panel derecho con tabs
        right = ttk.Frame(split)
        split.add(right, weight=2)

        self.tabs = ttk.Notebook(right)
        self.tabs.pack(fill="both", expand=True)

        # ACTIVIDAD
        tab_act = ttk.Frame(self.tabs)
        self.tabs.add(tab_act, text="Actividad")

        ttk.Label(tab_act, text="Vista/Edición rápida de 'Actividad':").grid(row=0, column=0, columnspan=3, sticky="w")

        tf = ttk.Frame(tab_act)
        tf.grid(row=1, column=0, columnspan=3, sticky="nsew", pady=(4,0))

        self.activity_text = tk.Text(tf, wrap="word")
        self.activity_text.grid(row=0, column=0, sticky="nsew")

        act_ys = ttk.Scrollbar(tf, orient="vertical", command=self.activity_text.yview)
        act_xs = ttk.Scrollbar(tf, orient="horizontal", command=self.activity_text.xview)
        self.activity_text.configure(yscrollcommand=act_ys.set, xscrollcommand=act_xs.set)
        act_ys.grid(row=0, column=1, sticky="ns")
        act_xs.grid(row=1, column=0, sticky="ew")

        tf.rowconfigure(0, weight=1)
        tf.columnconfigure(0, weight=1)

        btns = ttk.Frame(tab_act)
        btns.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(8,0))
        ttk.Button(btns, text="Aplicar a fila seleccionada", command=self.apply_activity_to_selected).grid(row=0, column=0, padx=2, pady=2, sticky="ew")
        ttk.Button(btns, text="Editar fila (modal)", command=self.edit_selected_modal).grid(row=0, column=1, padx=2, pady=2, sticky="ew")
        ttk.Button(btns, text="Añadir fila", command=self.add_row).grid(row=1, column=0, padx=2, pady=2, sticky="ew")
        ttk.Button(btns, text="Eliminar fila", command=self.delete_selected).grid(row=1, column=1, padx=2, pady=2, sticky="ew")
        ttk.Button(btns, text="Exportar a Excel…", command=self.export_to_excel).grid(row=2, column=0, padx=2, pady=6, sticky="ew")
        ttk.Button(btns, text="Restablecer datos", command=self.reset_data).grid(row=2, column=1, padx=2, pady=6, sticky="ew")
        for i in (0,1): btns.grid_columnconfigure(i, weight=1)

        tab_act.rowconfigure(1, weight=1)
        tab_act.columnconfigure(0, weight=1)

        # CALENDARIO + PAGOS
        tab_cal = ttk.Frame(self.tabs)
        self.tabs.add(tab_cal, text="Calendario")

        v_split = ttk.Panedwindow(tab_cal, orient="vertical")
        v_split.pack(fill="both", expand=True)

        cal_holder = ttk.Frame(v_split)
        v_split.add(cal_holder, weight=3)
        self.yearcal = YearCalendarFrame(
            cal_holder, year=date.today().year,
            on_date_click=self._on_calendar_click,
            on_date_double_click=self._insert_date_in_activity,
            has_events_predicate=lambda y,m,d: self.index.has_day(date(y,m,d))
        )
        self.yearcal.pack(fill="both", expand=True)

        tools = ttk.Frame(cal_holder)
        tools.pack(fill="x", pady=(4,0))
        ttk.Button(tools, text="Actualizar pagos (web)", command=self._update_from_web).pack(side="left")
        ttk.Button(tools, text="Importar circulares (PDF)…", command=self._import_pdfs).pack(side="left", padx=(6,0))
        ttk.Button(tools, text="Vaciar pagos (mantener heurística)", command=self._reset_to_heuristics).pack(side="right")

        self.pay_frame = PaymentsInfoFrame(v_split)
        v_split.add(self.pay_frame, weight=2)
        try:
            v_split.paneconfigure(cal_holder, minsize=140)
            v_split.paneconfigure(self.pay_frame, minsize=100)
        except Exception:
            pass

        self.tree.bind("<<TreeviewSelect>>", self._on_select)
        self.tree.bind("<Double-1>", self._on_double_click_cell)

    # -------- Índice de pagos / avisos ----------
    def _populate_index_from_heuristics_for_year(self, year: int):
        start_ant, end_ant, start_sal, end_sal = self.heur.get_ranges_for_campaign(year if date.today().month >= 10 else year-1)
        self.index.add_range(start_ant, end_ant, tipo="Anticipo ayudas directas", fondo="FEAGA",
                             detalle="Ventana general de anticipos (heurística).", fuente="Heurística FEGA")
        self.index.add_range(start_sal, end_sal, tipo="Saldo ayudas directas", fondo="FEAGA",
                             detalle="Ventana general de saldos (heurística).", fuente="Heurística FEGA")

    def _effective_today_items(self):
        """Items relevantes para avisar hoy (excluye 'Sin pagos...' y rellenos)."""
        today = date.today()
        items = self.index.get_day(today)
        # Añadir heurística FEAGA de hoy (si aplica)
        items += [i for i in self.heur.get_for_day(today)
                  if i["fondo"] == "FEAGA" and not i["tipo"].startswith("Sin pagos")]
        # Deduplicar
        seen, uniq = set(), []
        for it in items:
            key = (it["tipo"], it["fondo"], it["detalle"])
            if key in seen:
                continue
            seen.add(key); uniq.append(it)
        return uniq

    def _update_today_banner(self):
        items = self._effective_today_items()
        if items:
            tipos = ", ".join(sorted({it["tipo"] for it in items}))
            self.alert_msg.config(text=f"Hoy ({date.today().strftime('%d/%m/%Y')}) hay pagos vigentes: {tipos}.")
            self.alert_bar.pack(fill="x", padx=10, pady=(10,0))
        else:
            self.alert_bar.pack_forget()

    def _open_today(self):
        # Ir a la pestaña Calendario y mostrar HOY
        for i in range(self.tabs.index("end")):
            if self.tabs.tab(i, "text") == "Calendario":
                self.tabs.select(i)
                break
        self.yearcal.go_to_date(date.today())

    def _update_from_web(self):
        def run():
            try:
                scraper = FegaWebScraper()
                if not scraper.available():
                    raise RuntimeError("Falta 'requests' (y opcionalmente 'beautifulsoup4').")
                scraper.fetch_into_index(self.index, year_hint=date.today().year)
                self._refresh_calendar_async_ok("Pagos actualizados desde FEGA (web).")
            except Exception as ex:
                self._refresh_calendar_async_err(f"No se pudo completar el scraping web: {ex}")

        threading.Thread(target=run, daemon=True).start()
        messagebox.showinfo("Actualizando", "Buscando ventanas en notas FEGA...\n(Se añadirá al índice de pagos).")

    def _import_pdfs(self):
        folder = filedialog.askdirectory(title="Selecciona la carpeta con las circulares FEGA (PDF)")
        if not folder: return
        def run():
            try:
                ing = FegaPDFIngestor()
                if not ing.available(): raise RuntimeError("Falta 'PyPDF2'.")
                ingest_folder = Path(folder)
                ing.ingest_folder(ingest_folder, self.index, default_year=date.today().year)
                self._refresh_calendar_async_ok(f"Se importaron PDFs de {folder}")
            except Exception as ex:
                self._refresh_calendar_async_err(f"No se pudieron importar los PDFs: {ex}")
        threading.Thread(target=run, daemon=True).start()
        messagebox.showinfo("Importando", "Leyendo circulares PDF… Esto puede tardar unos segundos.")

    def _reset_to_heuristics(self):
        self.index.clear()
        self._populate_index_from_heuristics_for_year(self.yearcal.current_year)
        self.yearcal.refresh()
        self._update_today_banner()
        messagebox.showinfo("Pagos", "Índice reiniciado (sólo heurística).")

    def _refresh_calendar_async_ok(self, msg):
        def _do():
            self.yearcal.refresh()
            self._update_today_banner()
            messagebox.showinfo("Listo", msg)
        self.after(0, _do)

    def _refresh_calendar_async_err(self, msg):
        self.after(0, lambda: messagebox.showerror("Error", msg))

    # -------- Integración calendario ----------
    def _on_calendar_click(self, y,m,d):
        dt = date(y,m,d)
        items = self.index.get_day(dt) + [i for i in self.heur.get_for_day(dt) if not i["tipo"].startswith("Sin pagos")]
        # Deduplicar
        seen=set(); uniq=[]
        for it in items:
            key=(it["tipo"], it["fondo"], it["detalle"])
            if key in seen: continue
            seen.add(key); uniq.append(it)
        self.pay_frame.show(dt, uniq)

    def _insert_date_in_activity(self, y,m,d):
        self.activity_text.insert("insert", f"- {d:02d}/{m:02d}/{y}: ")
        self.activity_text.see("insert")

    # -------- Lógica tabla ----------
    def _load_into_tree(self, df: pd.DataFrame):
        self.tree.delete(*self.tree.get_children())
        for _, row in df.iterrows():
            self.tree.insert("", "end", values=[row[c] for c in self.columns])

    def _get_selected_item(self):
        sel = self.tree.selection()
        return sel[0] if sel else None

    def _on_select(self, _=None):
        item = self._get_selected_item()
        if not item: return
        vals = self.tree.item(item, "values")
        self.activity_text.delete("1.0", "end")
        self.activity_text.insert("1.0", vals[1])

    def edit_selected_modal(self):
        item = self._get_selected_item()
        if not item: messagebox.showinfo("Editar fila","Selecciona primero una fila."); return
        vals = list(self.tree.item(item,"values"))
        win = tk.Toplevel(self); win.title("Editar fila"); win.transient(self.winfo_toplevel()); win.grab_set()
        ttk.Label(win, text="Mes:").grid(row=0, column=0, sticky="w", padx=6, pady=(8,2))
        e_mes = ttk.Entry(win, width=50); e_mes.grid(row=0, column=1, sticky="ew", padx=6, pady=(8,2)); e_mes.insert(0, vals[0])
        ttk.Label(win, text="Actividad:").grid(row=1, column=0, sticky="nw", padx=6, pady=(8,2))
        t_act = tk.Text(win, width=80, height=15, wrap="word"); t_act.grid(row=1, column=1, sticky="ew", padx=6, pady=(8,2)); t_act.insert("1.0", vals[1])
        ttk.Label(win, text="Ayuda FEAGA:").grid(row=2, column=0, sticky="w", padx=6, pady=(8,2))
        e_feaga = ttk.Entry(win, width=50); e_feaga.grid(row=2, column=1, sticky="ew", padx=6, pady=(8,2)); e_feaga.insert(0, vals[2])
        ttk.Label(win, text="Ayuda FEADER:").grid(row=3, column=0, sticky="w", padx=6, pady=(8,2))
        e_feader = ttk.Entry(win, width=50); e_feader.grid(row=3, column=1, sticky="ew", padx=6, pady=(8,2)); e_feader.insert(0, vals[3])
        btns = ttk.Frame(win); btns.grid(row=4, column=0, columnspan=2, sticky="ew", padx=6, pady=10)
        btns.grid_columnconfigure(0, weight=1); btns.grid_columnconfigure(1, weight=1)
        def aceptar():
            new_vals=[e_mes.get().strip(), t_act.get("1.0","end").rstrip("\n"), e_feaga.get().strip(), e_feader.get().strip()]
            self.tree.item(item, values=new_vals); self._on_select(); win.destroy()
        ttk.Button(btns, text="Aceptar", command=aceptar).grid(row=0, column=0, padx=5)
        ttk.Button(btns, text="Cancelar", command=win.destroy).grid(row=0, column=1, padx=5)
        for r in range(4): win.grid_rowconfigure(r, weight=0)
        win.grid_rowconfigure(1, weight=1); win.grid_columnconfigure(1, weight=1)

    def _on_double_click_cell(self, event):
        item = self.tree.identify_row(event.y); column = self.tree.identify_column(event.x)
        if not item or not column: return
        col_index = int(column.replace("#","")) - 1; col_name = self.columns[col_index]
        vals = list(self.tree.item(item,"values"))
        if col_name == "Actividad": self.edit_selected_modal(); return
        x,y,w,h = self.tree.bbox(item, column)
        top = tk.Toplevel(self); top.overrideredirect(True); top.geometry(f"{w}x{h}+{self.tree.winfo_rootx()+x}+{self.tree.winfo_rooty()+y}")
        entry = ttk.Entry(top); entry.insert(0, vals[col_index]); entry.select_range(0,'end'); entry.focus(); entry.pack(fill="both", expand=True)
        def save_and_close(_=None):
            vals[col_index]=entry.get(); self.tree.item(item, values=vals); top.destroy(); self._on_select()
        entry.bind("<Return>", save_and_close); entry.bind("<Escape>", lambda e: top.destroy()); entry.bind("<FocusOut>", save_and_close)

    def apply_activity_to_selected(self):
        item = self._get_selected_item()
        if not item: messagebox.showinfo("Aplicar","Selecciona primero una fila."); return
        vals = list(self.tree.item(item,"values"))
        vals[1] = self.activity_text.get("1.0","end").rstrip("\n")
        self.tree.item(item, values=vals)
        messagebox.showinfo("Aplicar","Actividad actualizada en la fila seleccionada.")

    def add_row(self):
        self.tree.insert("", "end", values=["","","",""])
        children = self.tree.get_children()
        if children: self.tree.selection_set(children[-1]); self.tree.see(children[-1])

    def delete_selected(self):
        item = self._get_selected_item()
        if not item: messagebox.showinfo("Eliminar fila","Selecciona primero una fila."); return
        self.tree.delete(item); self.activity_text.delete("1.0","end")

    def export_to_excel(self):
        df = self._df_from_tree()
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")],
                                                 title="Guardar como", initialfile="tabla_FEAGA_FEADER_Aragon.xlsx")
        if not file_path: return
        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Calendario')
                workbook = writer.book; worksheet = writer.sheets['Calendario']
                wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
                actividad_col_idx = df.columns.get_loc('Actividad')
                actividad_col_letter = self._col_letter_from_index(actividad_col_idx)
                worksheet.set_column(f'{actividad_col_letter}:{actividad_col_letter}', 50, wrap_format)
                for idx, row in df.iterrows():
                    num_lines = str(row['Actividad']).count('\n') + 1
                    worksheet.set_row(idx + 1, 15 * num_lines)
            messagebox.showinfo("Exportación completada", f"Archivo guardado en:\n{file_path}")
        except Exception as ex:
            messagebox.showerror("Error al exportar", f"No se pudo exportar el Excel.\n\nDetalle: {ex}")

    def reset_data(self):
        if messagebox.askyesno("Restablecer", "¿Restablecer los datos originales? Se perderán los cambios no guardados."):
            self._build_dataset(); self._load_into_tree(self.df); self.activity_text.delete("1.0","end")

    def _df_from_tree(self) -> pd.DataFrame:
        rows = [self.tree.item(i,"values") for i in self.tree.get_children()]
        return pd.DataFrame(rows, columns=self.columns)

    @staticmethod
    def _col_letter_from_index(idx0: int) -> str:
        idx = idx0; letters = ""
        while True:
            idx, rem = divmod(idx, 26); letters = chr(ord('A') + rem) + letters
            if idx == 0: break
            idx -= 1
        return letters

# =========================
#  App
# =========================
def main():
    root = tk.Tk()
    root.title("Editor Calendario FEAGA/FEADER (v0.8)")
    root.geometry("1280x820")
    try:
        from ctypes import windll; windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    try:
        import platform
        ttk.Style().theme_use("winnative" if platform.system()=="Windows" else "clam")
    except Exception:
        pass
    app = CalendarioFEAGA_FEADERFrame(root)
    app.pack(fill="both", expand=True)
    root.mainloop()

if __name__ == "__main__":
    main()
