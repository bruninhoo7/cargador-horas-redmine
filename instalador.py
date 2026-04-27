"""
Instalador gráfico — Cargador de Horas Redmine
HM Consulting  v1.1.1
4 pantallas: Bienvenida → Disclaimer → Instalación → Finalizado
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading, subprocess, sys, os, shutil

APP_NAME    = "Cargador de Horas Redmine"
APP_VERSION = "1.1.1"
APP_AUTHOR  = "HM Consulting"
EXE_NAME    = "CargadorHorasRedmine.exe"
AZUL_OSCURO = "#1a2a4a"
AZUL_CLARO  = "#00aadd"
BG          = "#f5f5f5"

DEFAULT_DIR = os.path.join(os.environ.get("LOCALAPPDATA", "C:\\"), "CargadorHorasRedmine")
ASSETS      = ["logo_app.png", "logo_instalador.png", "HM_Icono.ico", "icono_acerca.png"]

DISCLAIMER_TEXT = """AVISO LEGAL Y TERMINOS DE USO

Esta aplicacion, "Cargador de Horas Redmine", fue desarrollada con fines exclusivamente \
practicos y operativos para facilitar la carga de horas en la plataforma Redmine. \
Su creacion responde a una necesidad de automatizacion interna y no persigue ningun \
proposito comercial, malicioso ni de recopilacion indebida de informacion.

USO DE DATOS
La aplicacion utiliza unicamente los datos que el propio usuario ingresa (credenciales \
de API, rutas de archivos y registros de horas). Dichos datos se almacenan localmente \
en el equipo del usuario y se transmiten exclusivamente hacia la instancia de Redmine \
configurada por el mismo usuario. En ningun caso se envian datos a servidores externos \
ni a terceros.

CARACTER NO OFICIAL
Esta herramienta NO es un producto oficial de HM Consulting ni de ninguna plataforma \
de gestion de proyectos. Es una utilidad desarrollada internamente para uso practico \
del equipo, sin garantias de soporte formal ni de continuidad.

LIMITACION DE RESPONSABILIDAD
El uso de esta aplicacion es bajo responsabilidad exclusiva del usuario. Los \
desarrolladores no se hacen responsables por perdidas de datos, errores en registros \
de horas, o cualquier inconveniente derivado del uso incorrecto de la herramienta.

Al continuar con la instalacion, el usuario declara haber leido y aceptado los \
presentes terminos."""


def get_base_dir():
    if getattr(sys, "frozen", False):
        return sys._MEIPASS
    return os.path.dirname(os.path.abspath(__file__))

def get_desktop_path():
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders")
        desktop, _ = winreg.QueryValueEx(key, "Desktop")
        winreg.CloseKey(key)
        return desktop
    except Exception:
        return os.path.join(os.environ.get("USERPROFILE", ""), "Desktop")

def crear_acceso_directo(exe_path, shortcut_path):
    import tempfile
    work_dir = os.path.dirname(exe_path)
    ps_lines = [
        "$ws = New-Object -ComObject WScript.Shell",
        f"$s = $ws.CreateShortcut('{shortcut_path.replace(chr(39), chr(39)*2)}')",
        f"$s.TargetPath = '{exe_path.replace(chr(39), chr(39)*2)}'",
        f"$s.WorkingDirectory = '{work_dir.replace(chr(39), chr(39)*2)}'",
        f"$s.IconLocation = '{exe_path.replace(chr(39), chr(39)*2)}'",
        "$s.Save()",
    ]
    ps_file = tempfile.mktemp(suffix=".ps1")
    with open(ps_file, "w", encoding="utf-8") as f:
        f.write("\n".join(ps_lines))
    CREATE_NO_WINDOW = 0x08000000
    subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", ps_file],
                   capture_output=True,
                   creationflags=CREATE_NO_WINDOW)
    try: os.remove(ps_file)
    except: pass

def instalar(destino, crear_shortcut, log_cb, on_done):
    def L(m): log_cb(m)
    try:
        base = get_base_dir()
        if os.path.exists(destino):
            L("Eliminando version anterior...")
            for f in os.listdir(destino):
                fp = os.path.join(destino, f)
                try:
                    if os.path.isfile(fp): os.remove(fp)
                except: pass
            L("   OK")

        os.makedirs(destino, exist_ok=True)
        L(f"Carpeta: {destino}")

        L("\nInstalando aplicacion...")
        exe_src = os.path.join(base, EXE_NAME)
        if not os.path.exists(exe_src):
            L(f"ERROR: No se encontro {EXE_NAME}"); on_done(False, None); return
        exe_path = os.path.join(destino, EXE_NAME)
        shutil.copy2(exe_src, exe_path)
        L(f"   OK: {EXE_NAME}")

        for asset in ASSETS:
            src = os.path.join(base, asset)
            if os.path.exists(src):
                shutil.copy2(src, os.path.join(destino, asset))
                L(f"   OK: {asset}")

        if crear_shortcut:
            L("\nCreando acceso directo...")
            desktop  = get_desktop_path()
            shortcut = os.path.join(desktop, f"{APP_NAME}.lnk")
            crear_acceso_directo(exe_path, shortcut)
            L("   OK" if os.path.exists(shortcut) else "   No se pudo crear")

        L("\n" + "=" * 42)
        L(f"  Instalacion completada  v{APP_VERSION}")
        L("=" * 42)
        on_done(True, exe_path)

    except Exception as e:
        L(f"\nError: {e}"); on_done(False, None)


# ============================================================
#  VENTANA BASE
# ============================================================

class Installer(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Instalador - {APP_NAME}  v{APP_VERSION}")
        self.resizable(False, False)
        self.configure(bg=BG)

        ico = os.path.join(get_base_dir(), "HM_Icono.ico")
        if os.path.exists(ico):
            try: self.iconbitmap(ico)
            except: pass

        self._exe_path = None
        self._container = tk.Frame(self, bg=BG)
        self._container.pack(fill="both", expand=True)

        self._show_screen(ScreenBienvenida)
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

    def _show_screen(self, ScreenClass, **kwargs):
        for w in self._container.winfo_children():
            w.destroy()
        ScreenClass(self._container, self, **kwargs).pack(fill="both", expand=True)

    def _make_header(self, parent, subtitle=""):
        hf = tk.Frame(parent, bg="#ffffff")
        hf.pack(fill="x")
        logo_path = os.path.join(get_base_dir(), "logo_instalador.png")
        if os.path.exists(logo_path):
            try:
                from PIL import Image, ImageTk
                img   = Image.open(logo_path).convert("RGB")
                img   = img.resize((320, 92), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                lbl   = tk.Label(hf, image=photo, bg="#ffffff")
                lbl.image = photo
                lbl.pack(padx=20, pady=8)
            except:
                tk.Label(hf, text=APP_AUTHOR, font=("Segoe UI", 16, "bold"),
                         bg="#ffffff", fg=AZUL_OSCURO).pack(pady=16)
        else:
            tk.Label(hf, text=APP_AUTHOR, font=("Segoe UI", 16, "bold"),
                     bg="#ffffff", fg=AZUL_OSCURO).pack(pady=16)

        tk.Label(parent, text=subtitle, bg=AZUL_OSCURO, fg="#aaccee",
                 font=("Segoe UI", 9)).pack(fill="x", ipady=4)

    def _make_nav(self, parent, back_cmd=None, next_cmd=None,
                  back_text="< Atras", next_text="Siguiente >",
                  next_state="normal"):
        nf = tk.Frame(parent, bg="#e8e8e8", pady=8, padx=16)
        nf.pack(fill="x", side="bottom")
        tk.Frame(nf, bg="#cccccc", height=1).pack(fill="x", side="top", pady=(0,8))
        if back_cmd:
            ttk.Button(nf, text=back_text, command=back_cmd).pack(side="left")
        if next_cmd:
            btn = ttk.Button(nf, text=next_text, command=next_cmd, state=next_state)
            btn.pack(side="right")
            return btn
        return None


# ============================================================
#  PANTALLA 1 — BIENVENIDA
# ============================================================

class ScreenBienvenida(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self.app = app
        app._make_header(self, subtitle=f"Asistente de instalacion  -  v{APP_VERSION}")

        body = tk.Frame(self, bg=BG, padx=30, pady=20)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Bienvenido al instalador de",
                 font=("Segoe UI", 11), bg=BG, fg="#555").pack(pady=(10, 4))
        tk.Label(body, text=APP_NAME,
                 font=("Segoe UI", 15, "bold"), bg=BG, fg=AZUL_OSCURO).pack()
        tk.Label(body, text=APP_AUTHOR,
                 font=("Segoe UI", 11), bg=BG, fg=AZUL_CLARO).pack(pady=(2, 20))

        tk.Frame(body, bg="#e0e0e0", height=1).pack(fill="x", pady=8)

        tk.Label(body,
                 text="Este asistente te guidara paso a paso\n"
                      "para instalar la aplicacion en tu equipo.\n\n"
                      "Hace click en Siguiente para continuar.",
                 font=("Segoe UI", 10), bg=BG, fg="#444",
                 justify="center").pack(pady=10)

        app._make_nav(self,
                      next_cmd=lambda: app._show_screen(ScreenDisclaimer),
                      next_text="Siguiente >")


# ============================================================
#  PANTALLA 2 — DISCLAIMER
# ============================================================

class ScreenDisclaimer(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self.app = app
        app._make_header(self, subtitle="Aviso legal y terminos de uso")

        body = tk.Frame(self, bg=BG, padx=20, pady=10)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Lee los siguientes terminos antes de continuar:",
                 font=("Segoe UI", 9, "bold"), bg=BG).pack(anchor="w", pady=(0, 6))

        txt_frame = tk.Frame(body, bg=BG)
        txt_frame.pack(fill="both", expand=True)
        sb = ttk.Scrollbar(txt_frame)
        sb.pack(side="right", fill="y")
        txt = tk.Text(txt_frame, wrap="word", font=("Segoe UI", 8),
                      yscrollcommand=sb.set, bg="#f9f9f9",
                      relief="solid", bd=1, padx=10, pady=8, height=14)
        txt.pack(side="left", fill="both", expand=True)
        sb.config(command=txt.yview)
        txt.insert("1.0", DISCLAIMER_TEXT)
        txt.config(state="disabled")

        self.v_acepto = tk.BooleanVar(value=False)
        ttk.Checkbutton(body,
                        text="He leido y acepto los terminos descritos anteriormente.",
                        variable=self.v_acepto,
                        command=self._toggle_next).pack(anchor="w", pady=(8, 0))

        self._btn_next = app._make_nav(
            self,
            back_cmd=lambda: app._show_screen(ScreenBienvenida),
            next_cmd=self._continuar,
            next_text="Acepto >",
            next_state="disabled")

    def _toggle_next(self):
        if self._btn_next:
            self._btn_next.config(state="normal" if self.v_acepto.get() else "disabled")

    def _continuar(self):
        if self.v_acepto.get():
            self.app._show_screen(ScreenInstalacion)


# ============================================================
#  PANTALLA 3 — INSTALACION
# ============================================================

class ScreenInstalacion(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self.app = app
        app._make_header(self, subtitle="Configuracion de la instalacion")

        body = tk.Frame(self, bg=BG, padx=20, pady=12)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="Carpeta de instalacion",
                 font=("Segoe UI", 10, "bold"), bg=BG).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))

        self.v_dest = tk.StringVar(value=DEFAULT_DIR)
        ttk.Entry(body, textvariable=self.v_dest, width=46).grid(
            row=1, column=0, columnspan=2, sticky="ew")
        ttk.Button(body, text="...", width=3,
                   command=self._browse).grid(row=1, column=2, padx=(6, 0))

        tk.Frame(body, bg="#e0e0e0", height=1).grid(
            row=2, column=0, columnspan=3, sticky="ew", pady=10)

        self.v_shortcut = tk.BooleanVar(value=True)
        ttk.Checkbutton(body, text="Crear acceso directo en el escritorio",
                        variable=self.v_shortcut).grid(
            row=3, column=0, columnspan=3, sticky="w")

        tk.Frame(body, bg="#e0e0e0", height=1).grid(
            row=4, column=0, columnspan=3, sticky="ew", pady=10)

        # Log colapsable
        self.v_log_expanded = tk.BooleanVar(value=False)
        log_toggle = tk.Label(body, text="+ Ver log de instalacion",
                              font=("Segoe UI", 9), bg=BG, fg=AZUL_CLARO,
                              cursor="hand2")
        log_toggle.grid(row=5, column=0, columnspan=3, sticky="w")
        log_toggle.bind("<Button-1>", self._toggle_log)
        self._log_toggle_lbl = log_toggle

        self.log_frame = tk.Frame(body, bg=BG)
        self.log_box = tk.Text(self.log_frame, width=56, height=7,
                               font=("Consolas", 7), bg="#1e1e1e", fg="#d4d4d4",
                               state="disabled", relief="flat")
        self.log_box.pack(fill="both", expand=True)

        self.status_var = tk.StringVar(value="")
        tk.Label(body, textvariable=self.status_var, bg=BG, fg="gray",
                 font=("Segoe UI", 8)).grid(row=7, column=0, columnspan=3, sticky="w", pady=(6,0))

        body.columnconfigure(1, weight=1)

        self._btn_next = app._make_nav(
            self,
            back_cmd=lambda: app._show_screen(ScreenDisclaimer),
            next_cmd=self._iniciar,
            next_text="Instalar >")

    def _browse(self):
        p = filedialog.askdirectory(title="Elegi la carpeta de instalacion")
        if p: self.v_dest.set(p.replace("/", "\\"))

    def _toggle_log(self, event=None):
        if self.v_log_expanded.get():
            self.log_frame.grid_remove()
            self._log_toggle_lbl.config(text="+ Ver log de instalacion")
            self.v_log_expanded.set(False)
        else:
            self.log_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(4, 0))
            self._log_toggle_lbl.config(text="- Ocultar log de instalacion")
            self.v_log_expanded.set(True)

    def _log(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _iniciar(self):
        dest = self.v_dest.get().strip()
        if not dest:
            messagebox.showwarning("Falta carpeta", "Selecciona una carpeta."); return
        if self._btn_next:
            self._btn_next.config(state="disabled")
        self.status_var.set("Instalando...")

        def on_done(ok, exe_path):
            self.app._exe_path = exe_path
            if ok:
                self.status_var.set("Completado.")
                self.app._show_screen(ScreenFinalizado)
            else:
                self.status_var.set("Error en la instalacion.")
                if self._btn_next:
                    self._btn_next.config(state="normal")

        threading.Thread(target=instalar,
                         args=(dest, self.v_shortcut.get(), self._log, on_done),
                         daemon=True).start()


# ============================================================
#  PANTALLA 4 — FINALIZADO
# ============================================================

class ScreenFinalizado(tk.Frame):
    def __init__(self, parent, app):
        super().__init__(parent, bg=BG)
        self.app = app
        app._make_header(self, subtitle="Instalacion completada")

        body = tk.Frame(self, bg=BG, padx=30, pady=20)
        body.pack(fill="both", expand=True)

        tk.Label(body, text="OK", font=("Segoe UI", 48, "bold"),
                 fg="#2ecc71", bg=BG).pack(pady=(10, 4))

        tk.Label(body, text="Instalacion exitosa!",
                 font=("Segoe UI", 14, "bold"), bg=BG, fg=AZUL_OSCURO).pack()
        tk.Label(body, text=f"{APP_NAME}  v{APP_VERSION}",
                 font=("Segoe UI", 10), bg=BG, fg="#555").pack(pady=(2, 16))

        tk.Frame(body, bg="#e0e0e0", height=1).pack(fill="x", pady=8)

        self.v_ejecutar = tk.BooleanVar(value=True)
        ttk.Checkbutton(body,
                        text=f"Ejecutar {APP_NAME} al cerrar el instalador",
                        variable=self.v_ejecutar).pack(pady=(10, 0))

        app._make_nav(self,
                      next_cmd=self._finalizar,
                      next_text="Finalizar",
                      back_cmd=None)

    def _finalizar(self):
        if self.v_ejecutar.get() and self.app._exe_path:
            try:
                subprocess.Popen([self.app._exe_path])
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo ejecutar la app:\n{e}")
        self.app.destroy()


# ============================================================
#  ENTRY POINT
# ============================================================

if __name__ == "__main__":
    if not any("pyinstaller" in a.lower() for a in sys.argv):
        Installer().mainloop()
