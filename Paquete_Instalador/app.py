"""
Cargador de Horas → Redmine  |  HM Consulting
===============================================
v1.0.0  —  Autor: HM Consulting

Requisitos:
    pip install pandas openpyxl requests pillow
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json, os, sys, threading
import pandas as pd
import requests

# ============================================================
#  METADATA
# ============================================================

APP_NAME    = "Cargador de Horas Redmine"
APP_VERSION = "1.0.1"
APP_AUTHOR  = "HM Consulting"

# ============================================================
#  RUTAS
# ============================================================

def get_app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def get_data_dir():
    base = os.environ.get("APPDATA", get_app_dir())
    path = os.path.join(base, "HMConsulting", "CargadorHoras")
    os.makedirs(path, exist_ok=True)
    return path

def get_asset(name):
    return os.path.join(get_app_dir(), name)

CONFIG_FILE   = os.path.join(get_data_dir(), "config.json")
CLIENTES_FILE = os.path.join(get_data_dir(), "clientes.json")

# ============================================================
#  LISTAS DE OPCIONES
# ============================================================

ACTIVIDADES = [
    "Soporte", "Análisis", "Configuración", "Desarrollo",
    "Diseño", "Documentación", "Capacitación", "Migración de Datos",
    "Pruebas Unitarias", "Pruebas Integrales", "Reunión", "Otra",
]

UBICACIONES = [
    "Oficina Rosario", "Oficina BsAs", "Oficina Venado",
    "Remoto", "On Site",
]

DIAS_SEMANA = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
TIPOS_HORA  = ["Funcional", "SSFF", "ABAP"]

# ============================================================
#  CONFIG
# ============================================================

CONFIG_DEFAULT = {
    "redmine_url":      "http://206.189.195.128/redmine",
    "api_key":          "",
    "archivo_excel":    "",
    "hoja":             "Horas",
    "actividad":        "Soporte",
    "tipo_hora":        "Funcional",
    "ubicacion":        "Oficina Rosario",
    "ubicacion_remoto": "Remoto",
    "dia_remoto":       4,
}

def cargar_config():
    if os.path.exists(CONFIG_FILE):
        try:
            cfg = CONFIG_DEFAULT.copy()
            cfg.update(json.load(open(CONFIG_FILE, encoding="utf-8")))
            return cfg
        except Exception:
            pass
    return CONFIG_DEFAULT.copy()

def guardar_config(cfg):
    json.dump(cfg, open(CONFIG_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

def cargar_clientes():
    if os.path.exists(CLIENTES_FILE):
        try:
            return json.load(open(CLIENTES_FILE, encoding="utf-8"))
        except Exception:
            pass
    return []

def guardar_clientes(lista):
    json.dump(lista, open(CLIENTES_FILE, "w", encoding="utf-8"), ensure_ascii=False, indent=2)

# ============================================================
#  LÓGICA REDMINE
# ============================================================

def hdrs(api_key):
    return {"X-Redmine-API-Key": api_key, "Content-Type": "application/json"}

def obtener_proyectos(url, api_key):
    r = requests.get(f"{url}/projects.json?limit=100", headers=hdrs(api_key))
    if r.status_code == 200:
        return r.json().get("projects", []), None
    return None, f"HTTP {r.status_code}"

def obtener_id_actividad(url, api_key, nombre):
    r = requests.get(f"{url}/enumerations/time_entry_activities.json", headers=hdrs(api_key))
    if r.status_code != 200:
        return None, f"HTTP {r.status_code}"
    for a in r.json().get("time_entry_activities", []):
        if a["name"].lower() == nombre.lower():
            return a["id"], None
    return None, f"Actividad '{nombre}' no encontrada"

def obtener_mi_id(url, api_key):
    r = requests.get(f"{url}/users/current.json", headers=hdrs(api_key))
    if r.status_code == 200:
        return r.json()["user"]["id"], None
    return None, f"HTTP {r.status_code}"

def crear_issue(url, api_key, proyecto_id, titulo, usuario_id):
    payload = {"issue": {
        "project_id": proyecto_id, "subject": titulo,
        "tracker_id": 1, "assigned_to_id": usuario_id,
    }}
    r = requests.post(f"{url}/issues.json", headers=hdrs(api_key), data=json.dumps(payload))
    if r.status_code == 201:
        return r.json()["issue"]["id"], None
    return None, r.text

def cargar_entrada(url, api_key, issue_id, fecha, horas, comentario,
                   act_id, hs_cliente, tipo_hora, ubicacion):
    entry = {"time_entry": {
        "issue_id": int(issue_id), "spent_on": fecha,
        "hours": float(horas), "comments": comentario, "activity_id": act_id,
        "custom_fields": [
            {"id": 3, "value": str(hs_cliente)},
            {"id": 2, "value": tipo_hora},
            {"id": 4, "value": ubicacion},
        ]
    }}
    r = requests.post(f"{url}/time_entries.json", headers=hdrs(api_key), data=json.dumps(entry))
    return r.status_code, r.text

def armar_titulo(tema, descripcion):
    t = "" if pd.isna(tema)        else str(tema).strip()
    d = "" if pd.isna(descripcion) else str(descripcion).strip()
    if t and d: return f"{t} - {d}"
    return t or d

def ejecutar_carga(cfg, clientes_lista, log, on_done):
    def L(m): log(m)
    try:
        # Mapeo alias → proyecto_id (soporta múltiples nombres por coma)
        mapeo = {}
        for c in clientes_lista:
            nombres = c.get("nombres_excel", c.get("nombre_excel", ""))
            for nombre in [n.strip() for n in nombres.split(",") if n.strip()]:
                mapeo[nombre] = c["proyecto_id"]
        try:
            df = pd.read_excel(cfg["archivo_excel"], sheet_name=cfg["hoja"])
        except Exception as e:
            L(f"❌ No se pudo leer el Excel: {e}"); on_done(False); return
        L(f"✔ Excel cargado: {len(df)} filas")

        act_id, err = obtener_id_actividad(cfg["redmine_url"], cfg["api_key"], cfg["actividad"])
        if err: L(f"❌ {err}"); on_done(False); return
        L(f"✔ Actividad '{cfg['actividad']}' (ID: {act_id})")

        usr_id, err = obtener_mi_id(cfg["redmine_url"], cfg["api_key"])
        if err: L(f"❌ {err}"); on_done(False); return
        L(f"✔ Usuario ID: {usr_id}\n")

        ok = errores = omitidas = creados = 0
        cache = {}

        for idx, row in df.iterrows():
            fn = idx + 2
            titulo = armar_titulo(row.get("Tema"), row.get("Descripcion"))
            id_redmine = row.get("ID_Redmine")
            tiene_id = not (pd.isna(id_redmine) or str(id_redmine).strip() == "")

            if not tiene_id:
                cliente = str(row.get("Cliente", "")).strip()
                proyecto_id = mapeo.get(cliente)
                if proyecto_id is None:
                    L(f"  Fila {fn}: ⚠ Cliente '{cliente}' no configurado — omitida")
                    omitidas += 1; continue
                ck = (proyecto_id, titulo)
                if ck in cache:
                    id_redmine = cache[ck]
                    L(f"  Fila {fn}: 🔗 Issue ya creado #{id_redmine}")
                else:
                    nid, err = crear_issue(cfg["redmine_url"], cfg["api_key"],
                                           proyecto_id, titulo, usr_id)
                    if err:
                        L(f"  Fila {fn}: ❌ Error creando issue — {err[:100]}")
                        errores += 1; continue
                    cache[ck] = nid; id_redmine = nid; creados += 1
                    L(f"  Fila {fn}: ✨ Issue #{nid} creado — '{titulo[:40]}'")

            horas = row.get("HS Trabajadas")
            if pd.isna(horas): horas = 0

            fecha_raw = row.get("Fecha")
            try:
                fdt = pd.to_datetime(fecha_raw)
                fecha_str = fdt.strftime("%Y-%m-%d")
            except Exception:
                L(f"  Fila {fn}: ❌ Fecha inválida"); errores += 1; continue

            ubicacion  = cfg["ubicacion_remoto"] if fdt.weekday() == cfg["dia_remoto"] \
                         else cfg["ubicacion"]
            hs_cliente = row.get("HS Cliente", 0)
            if pd.isna(hs_cliente): hs_cliente = 0

            status, resp = cargar_entrada(
                cfg["redmine_url"], cfg["api_key"], id_redmine, fecha_str,
                horas, titulo, act_id, hs_cliente, cfg["tipo_hora"], ubicacion)
            if status == 201:
                L(f"  Fila {fn}: ✔ Issue #{int(id_redmine)} | {fecha_str} | {horas}hs | {ubicacion}")
                ok += 1
            else:
                L(f"  Fila {fn}: ❌ HTTP {status} — {resp[:100]}"); errores += 1

        L("\n" + "=" * 50)
        L(f"  ✨ Issues creados  : {creados}")
        L(f"  ✔ Horas cargadas  : {ok}")
        L(f"  ❌ Con errores     : {errores}")
        L(f"  ⏭ Omitidas        : {omitidas}")
        L("=" * 50)
        on_done(True)
    except Exception as e:
        L(f"\n❌ Error inesperado: {e}"); on_done(False)

# ============================================================
#  COLORES
# ============================================================

AZUL_OSCURO = "#1a2a4a"
AZUL_CLARO  = "#00aadd"
BG          = "#f0f4f8"

# ============================================================
#  APP
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME}  v{APP_VERSION}")
        self.resizable(False, False)
        self.configure(bg=BG)

        ico = get_asset("HM_Icono.ico")
        if os.path.exists(ico):
            try: self.iconbitmap(ico)
            except Exception: pass

        self.cfg      = cargar_config()
        self.clientes = cargar_clientes()
        self._cli_rows = []

        self._build()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _build(self):
        # Header con logo
        hf = tk.Frame(self, bg=AZUL_OSCURO)
        hf.pack(fill="x")
        # Lado izquierdo blanco para el logo
        hf_left = tk.Frame(hf, bg="white")
        hf_left.pack(side="left")
        logo_path = get_asset("logo_app.png")
        if os.path.exists(logo_path):
            try:
                from PIL import Image, ImageTk
                img   = Image.open(logo_path).resize((240, 56), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                lbl   = tk.Label(hf_left, image=photo, bg="white")
                lbl.image = photo
                lbl.pack(padx=16, pady=4)
            except Exception:
                tk.Label(hf_left, text=APP_AUTHOR, font=("Segoe UI", 13, "bold"),
                         bg="white", fg=AZUL_OSCURO).pack(padx=16, pady=8)
        else:
            tk.Label(hf_left, text=APP_AUTHOR, font=("Segoe UI", 13, "bold"),
                     bg="white", fg=AZUL_OSCURO).pack(padx=16, pady=8)
        # Lado derecho azul
        tk.Label(hf, text=f"v{APP_VERSION}", font=("Segoe UI", 9),
                 bg=AZUL_OSCURO, fg="#aaccee").pack(side="right", padx=16)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)
        self._tab_config(nb)
        self._tab_clientes(nb)
        self._tab_ejecutar(nb)

    # --- TAB CONFIGURACIÓN ---
    def _tab_config(self, nb):
        f = ttk.Frame(nb, padding=15)
        nb.add(f, text="⚙  Configuración")

        def lbl(t, r, bold=False):
            ttk.Label(f, text=t,
                      font=("Segoe UI", 10, "bold") if bold else ("Segoe UI", 9)
                      ).grid(row=r, column=0, sticky="w", pady=3)

        def entry(var, r, show=""):
            ttk.Entry(f, textvariable=var, width=46, show=show
                      ).grid(row=r, column=1, sticky="ew", padx=(10,0))

        def combo(var, vals, r, w=44):
            ttk.Combobox(f, textvariable=var, values=vals, width=w, state="readonly"
                         ).grid(row=r, column=1, sticky="w", padx=(10,0))

        lbl("Conexión a Redmine", 0, bold=True)
        self.v_url     = tk.StringVar(value=self.cfg["redmine_url"])
        self.v_api_key = tk.StringVar(value=self.cfg["api_key"])
        lbl("URL de Redmine", 1); entry(self.v_url, 1)
        lbl("API Key",         2); entry(self.v_api_key, 2, show="•")

        ttk.Separator(f).grid(row=3, column=0, columnspan=2, sticky="ew", pady=8)
        lbl("Archivo Excel", 4, bold=True)
        self.v_excel = tk.StringVar(value=self.cfg["archivo_excel"])
        self.v_hoja  = tk.StringVar(value=self.cfg["hoja"])
        lbl("Ruta del Excel", 5)
        xf = ttk.Frame(f)
        xf.grid(row=5, column=1, sticky="ew", padx=(10,0))
        ttk.Entry(xf, textvariable=self.v_excel, width=37).pack(side="left")
        ttk.Button(xf, text="📂", width=3, command=self._browse_excel).pack(side="left", padx=4)
        lbl("Hoja de horas", 6); entry(self.v_hoja, 6)

        ttk.Separator(f).grid(row=7, column=0, columnspan=2, sticky="ew", pady=8)
        lbl("Valores por defecto", 8, bold=True)
        self.v_actividad  = tk.StringVar(value=self.cfg["actividad"])
        self.v_tipo_hora  = tk.StringVar(value=self.cfg["tipo_hora"])
        self.v_ubicacion  = tk.StringVar(value=self.cfg["ubicacion"])
        self.v_ub_remoto  = tk.StringVar(value=self.cfg["ubicacion_remoto"])
        self.v_dia_remoto = tk.StringVar(value=DIAS_SEMANA[self.cfg["dia_remoto"]])

        lbl("Actividad",              9);  combo(self.v_actividad, ACTIVIDADES, 9)
        lbl("Tipo de Hora",           10); combo(self.v_tipo_hora, TIPOS_HORA, 10)
        lbl("Ubicación (normal)",     11); combo(self.v_ubicacion, UBICACIONES, 11)
        lbl("Ubicación (día remoto)", 12); combo(self.v_ub_remoto, UBICACIONES, 12)
        lbl("Día de trabajo remoto",  13); combo(self.v_dia_remoto, DIAS_SEMANA, 13, w=15)

        ttk.Separator(f).grid(row=14, column=0, columnspan=2, sticky="ew", pady=8)
        ttk.Button(f, text="💾  Guardar configuración",
                   command=self._guardar_config).grid(row=15, column=0, columnspan=2, pady=4)
        f.columnconfigure(1, weight=1)

    def _browse_excel(self):
        p = filedialog.askopenfilename(
            title="Seleccioná tu Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Todos", "*.*")])
        if p: self.v_excel.set(p)

    def _guardar_config(self):
        self.cfg.update({
            "redmine_url":      self.v_url.get().strip(),
            "api_key":          self.v_api_key.get().strip(),
            "archivo_excel":    self.v_excel.get().strip(),
            "hoja":             self.v_hoja.get().strip(),
            "actividad":        self.v_actividad.get(),
            "tipo_hora":        self.v_tipo_hora.get(),
            "ubicacion":        self.v_ubicacion.get(),
            "ubicacion_remoto": self.v_ub_remoto.get(),
            "dia_remoto":       DIAS_SEMANA.index(self.v_dia_remoto.get()),
        })
        guardar_config(self.cfg)
        messagebox.showinfo("Guardado", "✔ Configuración guardada.")

    # --- TAB CLIENTES ---
    def _tab_clientes(self, nb):
        f = ttk.Frame(nb, padding=15)
        nb.add(f, text="🏢  Clientes")

        ttk.Label(f, text="Conectate a Redmine para ver los proyectos disponibles.\n"
                          "Tildá los que usás y escribí los nombres de tu Excel separados por coma (ej: Zurich, GSK).",
                  font=("Segoe UI", 9), foreground="gray").pack(anchor="w")

        ttk.Button(f, text="🔄  Cargar proyectos desde Redmine",
                   command=self._cargar_proyectos).pack(anchor="w", pady=(8,4))

        self.lbl_cli = ttk.Label(f, text="", foreground="gray", font=("Segoe UI", 8))
        self.lbl_cli.pack(anchor="w")

        # Encabezados de tabla
        hdr = tk.Frame(f, bg=AZUL_OSCURO)
        hdr.pack(fill="x", pady=(6,0))
        for txt, w in [("✔", 3), ("Proyecto en Redmine", 22), ("Nombres en Excel (separados por coma)", 38)]:
            tk.Label(hdr, text=txt, width=w, bg=AZUL_OSCURO, fg="white",
                     font=("Segoe UI", 9, "bold"), anchor="w").pack(side="left", padx=4)

        # Canvas scrollable dentro de un frame para no interferir con el botón
        table_frame = tk.Frame(f)
        table_frame.pack(fill="both", expand=True)
        canvas = tk.Canvas(table_frame, height=240, bg="white", highlightthickness=0)
        sb     = ttk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
        self.cli_inner = tk.Frame(canvas, bg="white")
        self.cli_inner.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0,0), window=self.cli_inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # Precargar guardados
        for c in self.clientes:
            nombres = c.get("nombres_excel", c.get("nombre_excel", ""))
            self._fila_cliente({"id": c["proyecto_id"], "name": c["proyecto_nombre"]},
                               True, nombres)
        if self.clientes:
            self.lbl_cli.config(text=f"✔ {len(self.clientes)} clientes guardados.")

        # Botón FUERA del scroll, siempre visible abajo
        ttk.Separator(f).pack(fill="x", pady=(8,0))
        ttk.Button(f, text="💾  Guardar configuración de clientes",
                   command=self._guardar_clientes).pack(anchor="center", pady=(8,0))

    def _fila_cliente(self, proyecto, tildado=False, nombres_excel=""):
        rf = tk.Frame(self.cli_inner, bg="white")
        rf.pack(fill="x", pady=1)
        vc = tk.BooleanVar(value=tildado)
        vn = tk.StringVar(value=nombres_excel or proyecto["name"])
        ttk.Checkbutton(rf, variable=vc).pack(side="left", padx=4)
        tk.Label(rf, text=proyecto["name"][:28], width=22, bg="white",
                 anchor="w", font=("Segoe UI", 9)).pack(side="left")
        ttk.Entry(rf, textvariable=vn, width=38).pack(side="left", padx=4)
        self._cli_rows.append((vc, vn, proyecto))

    def _cargar_proyectos(self):
        url = self.cfg.get("redmine_url","")
        key = self.cfg.get("api_key","")
        if not key:
            messagebox.showwarning("Falta API Key", "Guardá tu API Key en Configuración primero.")
            return
        self.lbl_cli.config(text="⏳ Cargando proyectos...")
        self.update_idletasks()

        def fetch():
            proyectos, err = obtener_proyectos(url, key)
            if err:
                self.lbl_cli.config(text=f"❌ Error: {err}"); return
            for w in self.cli_inner.winfo_children(): w.destroy()
            self._cli_rows.clear()
            # Mapeo proyecto_id → nombres guardados
            guardados = {c["proyecto_id"]: c.get("nombres_excel", c.get("nombre_excel",""))
                         for c in self.clientes}
            for p in sorted(proyectos, key=lambda x: x["name"]):
                self._fila_cliente(p, p["id"] in guardados, guardados.get(p["id"],""))
            self.lbl_cli.config(text=f"✔ {len(proyectos)} proyectos cargados.")
        threading.Thread(target=fetch, daemon=True).start()

    def _guardar_clientes(self):
        res = []
        for vc, vn, p in self._cli_rows:
            nombres_raw = vn.get().strip()
            if vc.get() and nombres_raw:
                # Guardar lista de nombres separados por coma
                res.append({
                    "nombres_excel": nombres_raw,
                    "proyecto_id":   p["id"],
                    "proyecto_nombre": p["name"]
                })
        self.clientes = res
        guardar_clientes(res)
        total_alias = sum(len(c["nombres_excel"].split(",")) for c in res)
        messagebox.showinfo("Guardado", f"✔ {len(res)} proyecto(s) guardados con {total_alias} nombre(s) en total.")

    # --- TAB EJECUTAR ---
    def _tab_ejecutar(self, nb):
        f = ttk.Frame(nb, padding=15)
        nb.add(f, text="▶  Ejecutar carga")
        ttk.Label(f, text="Log de ejecución",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.log_box = scrolledtext.ScrolledText(
            f, width=72, height=22, state="disabled",
            font=("Consolas", 9), bg="#1e1e1e", fg="#d4d4d4")
        self.log_box.pack(fill="both", expand=True, pady=(6,10))
        bf = ttk.Frame(f)
        bf.pack(fill="x")
        self.btn_run = ttk.Button(bf, text="▶  Iniciar carga", command=self._iniciar)
        self.btn_run.pack(side="left")
        ttk.Button(bf, text="🗑  Limpiar", command=self._limpiar_log).pack(side="left", padx=8)
        self.status_var = tk.StringVar(value="Listo.")
        ttk.Label(bf, textvariable=self.status_var, foreground="gray").pack(side="right")

    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg+"\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _limpiar_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0","end")
        self.log_box.configure(state="disabled")

    def _iniciar(self):
        if not self.cfg.get("api_key"):
            messagebox.showwarning("Falta configuración","Completá la API Key."); return
        if not self.cfg.get("archivo_excel"):
            messagebox.showwarning("Falta configuración","Seleccioná el Excel."); return
        if not self.clientes:
            messagebox.showwarning("Sin clientes","Configurá clientes en la pestaña Clientes."); return
        self.btn_run.configure(state="disabled")
        self.status_var.set("⏳ Cargando...")
        self._log("="*50 + "\n  Iniciando carga...\n" + "="*50)
        def on_done(ok):
            self.btn_run.configure(state="normal")
            self.status_var.set("✔ Finalizado." if ok else "❌ Finalizado con errores.")
        threading.Thread(target=ejecutar_carga,
                         args=(self.cfg, self.clientes, self._log, on_done),
                         daemon=True).start()

# ============================================================
#  ENTRY POINT
# ============================================================

if __name__ == "__main__":
    if not any("pyinstaller" in a.lower() for a in sys.argv):
        App().mainloop()
