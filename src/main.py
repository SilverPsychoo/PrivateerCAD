from __future__ import annotations

import csv
import os
import threading
from tkinter import filedialog, messagebox

import customtkinter as ctk
from PIL import Image

import analizador
import extractor
from config import CAD_EXTENSIONS
from extractor_solidworks import SolidWorksSession

# ─── Paleta ──────────────────────────────────────────────────────────────────
C_BG        = "#0a0a12"
C_PANEL     = "#12121e"
C_CARD      = "#1a1a2e"
C_BORDER    = "#2a2a3e"
C_RED       = "#c0392b"
C_RED_HOVER = "#e74c3c"
C_RED_DIM   = "#7b241c"
C_GRAY      = "#8fa3b1"
C_TEXT      = "#dce8f0"
C_TEXT_DIM  = "#5d7080"
C_GREEN     = "#27ae60"
C_ORANGE    = "#e67e22"
C_YELLOW    = "#f1c40f"
C_BLUE      = "#5dade2"
C_WHITE     = "#ffffff"

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("dark-blue")


def _divider(parent) -> ctk.CTkFrame:
    return ctk.CTkFrame(parent, height=1, fg_color=C_BORDER, corner_radius=0)


# ─── App (composición, no herencia — compatibilidad Python 3.12) ─────────────

class PrivateerCAD:
    def __init__(self):
        # Usar instancia directa evita el bug de recursión en Python 3.12
        self.root = ctk.CTk()
        self.root.title("PrivateerCAD")
        self.root.geometry("1280x820")
        self.root.minsize(1100, 700)
        self.root.configure(fg_color=C_BG)

        self.df_actual = None
        self.relaciones_actuales: list = []
        self.modo_solidworks = False
        self.sw_session = None

        self._build_ui()
        self._ask_mode()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def mainloop(self):
        self.root.mainloop()

    # ─── UI ──────────────────────────────────────────────────────────────────

    def _build_ui(self):
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_main()
        self._build_statusbar()

    def _build_sidebar(self):
        sidebar = ctk.CTkFrame(self.root, width=224, fg_color=C_PANEL, corner_radius=0)
        sidebar.grid(row=0, column=0, rowspan=2, sticky="nsew")
        sidebar.grid_propagate(False)

        try:
            logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png")
            if os.path.isfile(logo_path):
                pil_img = Image.open(logo_path)
                ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(144, 78))
                ctk.CTkLabel(sidebar, image=ctk_img, text="").pack(pady=(26, 2))
        except Exception:
            ctk.CTkLabel(sidebar, text="⚓",
                         font=ctk.CTkFont(size=42), text_color=C_RED).pack(pady=(26, 2))

        ctk.CTkLabel(sidebar, text="PRIVATEERCAD",
                     font=ctk.CTkFont("Courier New", 15, "bold"),
                     text_color=C_RED).pack(pady=(2, 1))
        ctk.CTkLabel(sidebar, text="Inspector Forense CAD",
                     font=ctk.CTkFont("Courier New", 9),
                     text_color=C_TEXT_DIM).pack(pady=(0, 20))

        _divider(sidebar).pack(fill="x", padx=18, pady=(0, 16))

        self.lbl_modo = ctk.CTkLabel(sidebar, text="○  Iniciando…",
                                      font=ctk.CTkFont("Courier New", 10),
                                      text_color=C_TEXT_DIM, anchor="w")
        self.lbl_modo.pack(padx=20, fill="x", pady=(0, 16))

        _divider(sidebar).pack(fill="x", padx=18, pady=(0, 18))

        pk = dict(padx=14, pady=4, fill="x")

        self.btn_archivo = ctk.CTkButton(
            sidebar, text="  📄  Analizar Pieza",
            font=ctk.CTkFont("Courier New", 12, "bold"), height=44,
            fg_color=C_CARD, hover_color=C_RED_DIM, text_color=C_TEXT,
            corner_radius=6, anchor="w", command=self._analizar_archivo)
        self.btn_archivo.pack(**pk)

        self.btn_carpeta = ctk.CTkButton(
            sidebar, text="  📁  Analizar Grupo",
            font=ctk.CTkFont("Courier New", 12, "bold"), height=44,
            fg_color=C_RED, hover_color=C_RED_HOVER, text_color=C_WHITE,
            corner_radius=6, anchor="w", command=self._analizar_carpeta)
        self.btn_carpeta.pack(**pk)

        _divider(sidebar).pack(fill="x", padx=18, pady=(18, 14))

        sec = dict(font=ctk.CTkFont("Courier New", 11), height=36,
                   fg_color=C_CARD, hover_color=C_BORDER,
                   text_color=C_GRAY, corner_radius=6, anchor="w")

        self.btn_grafo = ctk.CTkButton(
            sidebar, text="  🕸  Red de distribución",
            state="disabled", command=self._ver_grafo, **sec)
        self.btn_grafo.pack(**pk)

        self.btn_exportar = ctk.CTkButton(
            sidebar, text="  💾  Exportar CSV",
            state="disabled", command=self._exportar_csv, **sec)
        self.btn_exportar.pack(**pk)

        self.btn_limpiar = ctk.CTkButton(
            sidebar, text="  ✕   Limpiar pantalla",
            command=self._limpiar, **sec)
        self.btn_limpiar.pack(**pk)

        ctk.CTkLabel(sidebar, text="v1.0",
                     font=ctk.CTkFont("Courier New", 9),
                     text_color=C_TEXT_DIM).pack(side="bottom", pady=10)

    def _build_main(self):
        main = ctk.CTkFrame(self.root, fg_color=C_BG, corner_radius=0)
        main.grid(row=0, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(main, fg_color=C_PANEL, height=50, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        header.grid_columnconfigure(1, weight=1)

        self.lbl_seccion = ctk.CTkLabel(
            header, text="RESULTADOS",
            font=ctk.CTkFont("Courier New", 13, "bold"), text_color=C_RED)
        self.lbl_seccion.grid(row=0, column=0, padx=20, pady=14)

        self.lbl_header_sub = ctk.CTkLabel(
            header, text="",
            font=ctk.CTkFont("Courier New", 10),
            text_color=C_TEXT_DIM, anchor="w")
        self.lbl_header_sub.grid(row=0, column=1, padx=8, sticky="w")

        self.textbox = ctk.CTkTextbox(
            main, font=ctk.CTkFont("Courier New", 12),
            fg_color=C_CARD, text_color=C_TEXT,
            scrollbar_button_color=C_BORDER,
            scrollbar_button_hover_color=C_RED_DIM,
            corner_radius=0, border_width=0, wrap="word")
        self.textbox.grid(row=1, column=0, sticky="nsew")

        for tag, color in (
            ("rojo",     "#e74c3c"),
            ("naranja",  "#e67e22"),
            ("verde",    "#2ecc71"),
            ("gris",     C_TEXT_DIM),
            ("azul",     C_BLUE),
            ("titulo",   C_RED),
            ("sep",      "#2a3a4a"),
            ("amarillo", C_YELLOW),
            ("blanco",   C_TEXT),
        ):
            self.textbox.tag_config(tag, foreground=color)

        self._bienvenida()

    def _build_statusbar(self):
        bar = ctk.CTkFrame(self.root, fg_color=C_PANEL, height=30, corner_radius=0)
        bar.grid(row=1, column=1, sticky="ew")
        bar.grid_propagate(False)
        bar.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(
            bar, text="Listo.",
            font=ctk.CTkFont("Courier New", 10),
            text_color=C_TEXT_DIM, anchor="w")
        self.lbl_status.grid(row=0, column=0, padx=16, pady=6, sticky="w")

        self.progress = ctk.CTkProgressBar(
            bar, width=150, height=5,
            fg_color=C_BORDER, progress_color=C_RED, corner_radius=2)
        self.progress.grid(row=0, column=1, padx=16, pady=10)
        self.progress.set(0)

    # ─── Modo ────────────────────────────────────────────────────────────────

    def _ask_mode(self):
        try:
            usar_sw = messagebox.askyesno(
                "Modo de análisis",
                "¿Usar la API de SolidWorks?\n\n"
                "Sí → Lee árbol de operaciones y metadatos completos\n"
                "No → Solo metadatos del sistema (sin árbol de operaciones)")
        except Exception:
            usar_sw = False

        if usar_sw:
            try:
                self.sw_session = SolidWorksSession()
                self.sw_session.connect()
                self.modo_solidworks = True
                self.lbl_modo.configure(text="●  SolidWorks API", text_color=C_GREEN)
            except Exception:
                self.modo_solidworks = False
                self.sw_session = None
                self.lbl_modo.configure(text="●  Windows / OLE", text_color=C_ORANGE)
                messagebox.showwarning(
                    "SolidWorks no disponible",
                    "No se pudo conectar a SolidWorks.\n"
                    "Se usará el modo Windows/OLE.")
        else:
            self.lbl_modo.configure(text="●  Windows / OLE", text_color=C_ORANGE)

    # ─── Analizar pieza ──────────────────────────────────────────────────────

    def _analizar_archivo(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo SolidWorks",
            filetypes=[("SolidWorks", "*.sldprt *.sldasm"), ("Todos", "*.*")])
        if not path:
            return

        nombre = os.path.basename(path)
        self.lbl_seccion.configure(text="ANÁLISIS DE PIEZA")
        self.lbl_header_sub.configure(text=nombre)
        self._lock_buttons()
        self._set_status(f"Analizando {nombre}…")
        self.progress.set(0.3)
        self.df_actual = None
        self.relaciones_actuales = []

        def _run():
            d = extractor.extraer_archivo(path, self.modo_solidworks, self.sw_session)
            self.root.after(0, lambda: self._render_archivo(d, nombre))

        threading.Thread(target=_run, daemon=True).start()

    def _render_archivo(self, datos, nombre):
        self.textbox.configure(state="normal")
        self.textbox.delete("0.0", "end")

        if not datos:
            self._w("El archivo no es un formato CAD compatible.\n", "rojo")
            self._unlock_buttons()
            return

        _, estado = analizador.diagnostico_unico(datos)

        autor       = datos.get("Autor_Original", "Desconocido")
        prop_win    = datos.get("Propietario_Windows", "")
        fc_sw       = datos.get("SW_Created_Date") or datos.get("Fecha_Creacion_SW", "")
        fs_sw       = datos.get("SW_Saved_Date")   or datos.get("Fecha_Ultimo_Guardado_SW", "")
        feats       = int(datos.get("Feature_Count") or 0)
        conf        = int(datos.get("Confidence") or 0)
        maquina     = prop_win or datos.get("Nombre_Maquina", "")
        feat_types  = datos.get("Feature_Types", "")

        self._sep()
        self._w(f"  {nombre}\n", "titulo")
        self._sep()
        self._w("\n")

        self._w("  IDENTIDAD\n", "azul")
        self._w("  " + "─" * 38 + "\n", "sep")
        self._fld("  Autor SW            ", autor)
        self._fld("  Máquina / Propietario", maquina or "No disponible")
        self._fld("  Confianza lectura   ", f"{conf}/100")
        self._fld("  Modo               ", datos.get("Modo", ""))
        self._w("\n")

        self._w("  FECHAS INTERNAS SW  (no cambian al copiar el archivo)\n", "azul")
        self._w("  " + "─" * 38 + "\n", "sep")
        self._fld("  Creado en SW        ", fc_sw or "No disponible")
        self._fld("  Último guardado SW  ", fs_sw or "No disponible")
        self._fld("  Modificación Windows", datos.get("Fecha_Modificacion", ""))
        self._w("\n")

        self._w("  ÁRBOL DE OPERACIONES\n", "azul")
        self._w("  " + "─" * 38 + "\n", "sep")

        if feats > 0 and feat_types:
            ops = [t.strip() for t in feat_types.split(">") if t.strip()]
            self._fld("  Total operaciones   ", str(len(ops)))
            self._w("\n")
            for i, op in enumerate(ops[:30], 1):
                self._w(f"    {i:>2}.  {op}\n", "gris")
            if len(ops) > 30:
                self._w(f"    … y {len(ops)-30} operaciones más\n", "gris")
        else:
            self._fld("  Total operaciones   ", "No disponible")
            if not self.modo_solidworks:
                self._w("  (Activa el modo SolidWorks API para leer el árbol)\n", "gris")
        self._w("\n")

        self._w("  VEREDICTO INDIVIDUAL\n", "azul")
        self._w("  " + "─" * 38 + "\n", "sep")
        v_map = {
            "LIMPIO":            ("🟢  Sin anomalías detectadas.", "verde"),
            "EVIDENCIA_PARCIAL": ("🟠  Datos parciales — compara en grupo.", "naranja"),
            "BAJA_CONFIANZA":    ("⚪  Evidencia insuficiente.", "gris"),
            "ERROR":             ("🔴  Error de lectura.", "rojo"),
        }
        txt, tag = v_map.get(estado, ("🔵  Analiza junto con el grupo.", "azul"))
        self._w(f"  {txt}\n\n", tag)
        self._w("  Para detectar plagio usa 'Analizar Grupo'\n"
                "  con todos los trabajos del grupo.\n", "gris")

        if fc_sw:
            self._w(f"\n  💡  Fecha SW '{fc_sw}' es la huella de esta sesión.\n", "amarillo")

        # Mostrar error de SW si existe (para diagnóstico)
        if datos.get("Error"):
            self._w(f"\n  ❌  Error SW: {datos['Error']}\n", "rojo")
        if datos.get("SolidWorks_Error"):
            self._w(f"  ❌  Error interno SW: {datos['SolidWorks_Error']}\n", "rojo")

        self._sep()
        self._unlock_buttons(msg=f"Análisis completado — {nombre}")

    # ─── Analizar carpeta ────────────────────────────────────────────────────

    def _analizar_carpeta(self):
        ruta = filedialog.askdirectory(title="Seleccionar carpeta del grupo")
        if not ruta:
            return

        self.lbl_seccion.configure(text="ANÁLISIS DE GRUPO")
        self.lbl_header_sub.configure(text=os.path.basename(ruta))
        self._lock_buttons()
        self.df_actual = None
        self.relaciones_actuales = []
        self.textbox.configure(state="normal")
        self.textbox.delete("0.0", "end")
        self._w("Escaneando carpeta…\n", "gris")

        def _run():
            archivos = []
            for raiz, _, files in os.walk(ruta):
                for f in sorted(files):
                    fp = os.path.join(raiz, f)
                    if fp.lower().endswith(CAD_EXTENSIONS):
                        archivos.append(fp)

            total = len(archivos)
            if total == 0:
                self.root.after(0, lambda: self._w(
                    "No se encontraron archivos .sldprt o .sldasm.\n", "rojo"))
                self.root.after(0, self._unlock_buttons)
                return

            self.root.after(0, lambda: self._set_status(f"Extrayendo {total} archivos…"))
            datos = []
            for i, fp in enumerate(archivos):
                prog = (i + 1) / total * 0.85
                nb   = os.path.basename(fp)
                self.root.after(0, lambda p=prog, n=nb, idx=i+1: (
                    self.progress.set(p),
                    self._set_status(f"{idx}/{total}  {n}")))
                d = extractor.extraer_archivo(fp, self.modo_solidworks, self.sw_session)
                if d:
                    datos.append(d)

            self.root.after(0, lambda: self._set_status("Comparando archivos…"))
            self.root.after(0, lambda: self.progress.set(0.95))

            df, reporte, relaciones = analizador.analizar_lote(datos)
            self.root.after(0, lambda: self._render_lote(df, reporte, relaciones, total))

        threading.Thread(target=_run, daemon=True).start()

    def _render_lote(self, df, reporte, relaciones, total):
        self.df_actual = df
        self.relaciones_actuales = relaciones

        self.textbox.configure(state="normal")
        self.textbox.delete("0.0", "end")

        for line in reporte.splitlines():
            if any(k in line for k in ("ALTO RIESGO", "🔴", "🚨", "COPIAS", "MISMO ORIGEN")):
                tag = "rojo"
            elif any(k in line for k in ("SOSPECHOSO", "🟠", "⚠️", "MISMO NÚMERO")):
                tag = "naranja"
            elif any(k in line for k in ("SIN ANOMALÍAS", "🟢")):
                tag = "verde"
            elif any(k in line for k in ("🦠", "DISTRIBUIDOR")):
                tag = "amarillo"
            elif any(k in line for k in ("👤",)):
                tag = "azul"
            elif "─" * 8 in line:
                tag = "sep"
            elif any(k in line for k in ("REPORTE", "DETALLE POR ARCHIVO")):
                tag = "titulo"
            elif any(k in line for k in ("⚪", "BAJA CONFIANZA")):
                tag = "gris"
            else:
                tag = ""
            self._w(line + "\n", tag)

        if relaciones:
            self.btn_grafo.configure(state="normal", fg_color=C_RED,
                                      hover_color=C_RED_HOVER, text_color=C_WHITE)
        if df is not None and not df.empty:
            self.btn_exportar.configure(state="normal")

        n_alto = int((df["Puntaje_Sospecha"] >= 80).sum()) if df is not None else 0
        msg = f"{total} archivos procesados"
        if n_alto:
            msg += f" · {n_alto} en ALTO RIESGO"
        self._unlock_buttons(msg=msg)

    # ─── Grafo ───────────────────────────────────────────────────────────────

    def _ver_grafo(self):
        if self.df_actual is not None and self.relaciones_actuales:
            analizador.mostrar_grafo(self.df_actual, self.relaciones_actuales)

    # ─── Exportar CSV ────────────────────────────────────────────────────────

    def _exportar_csv(self):
        if self.df_actual is None or self.df_actual.empty:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV", "*.csv")],
            initialfile="reporte_privateercad.csv")
        if not path:
            return
        cols_wanted = [
            "Archivo", "Estado", "Puntaje_Sospecha", "Autor_Original",
            "SW_Created_Date", "SW_Saved_Date", "Feature_Count",
            "Detalle_Sospecha", "Posible_Fuente", "Hash_Corto",
            "Tamano_Bytes", "Ruta_Completa",
        ]
        cols = [c for c in cols_wanted if c in self.df_actual.columns]
        try:
            self.df_actual[cols].to_csv(
                path, index=False, encoding="utf-8-sig", quoting=csv.QUOTE_ALL)
            self._set_status(f"Exportado → {os.path.basename(path)}")
        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))

    # ─── Helpers de escritura ────────────────────────────────────────────────

    def _w(self, text: str, tag: str = "") -> None:
        self.textbox.configure(state="normal")
        if tag:
            self.textbox.insert("end", text, tag)
        else:
            self.textbox.insert("end", text)
        self.textbox.see("end")

    def _fld(self, label: str, value: str) -> None:
        self.textbox.insert("end", label, "gris")
        low = (value or "").lower()
        if low in ("desconocido", "no disponible", ""):
            tag = "gris"
        elif "error" in low:
            tag = "rojo"
        else:
            tag = "blanco"
        self.textbox.insert("end", f": {value}\n", tag)

    def _sep(self) -> None:
        self.textbox.insert("end", "  " + "─" * 60 + "\n", "sep")

    def _set_status(self, msg: str) -> None:
        self.lbl_status.configure(text=msg)
        self.root.update_idletasks()

    def _lock_buttons(self) -> None:
        for btn in (self.btn_archivo, self.btn_carpeta,
                    self.btn_grafo, self.btn_exportar):
            btn.configure(state="disabled")

    def _unlock_buttons(self, msg: str = "Listo.") -> None:
        self.btn_archivo.configure(state="normal")
        self.btn_carpeta.configure(state="normal")
        self.progress.set(1.0)
        self._set_status(msg)
        self.root.after(2500, lambda: self.progress.set(0))

    def _limpiar(self) -> None:
        self.textbox.configure(state="normal")
        self.textbox.delete("0.0", "end")
        self.lbl_seccion.configure(text="RESULTADOS")
        self.lbl_header_sub.configure(text="")
        self.df_actual = None
        self.relaciones_actuales = []
        self.btn_grafo.configure(state="disabled", fg_color=C_CARD, text_color=C_GRAY)
        self.btn_exportar.configure(state="disabled")
        self.progress.set(0)
        self._set_status("Listo.")
        self._bienvenida()

    def _bienvenida(self) -> None:
        lines = [
            ("",                                                             ""),
            ("  PrivateerCAD — Inspector Forense SolidWorks",               "titulo"),
            ("",                                                             ""),
            ("  " + "─" * 52,                                               "sep"),
            ("",                                                             ""),
            ("  📄  Analizar Pieza",                                         "azul"),
            ("      Autor · Fechas internas SW · Árbol de operaciones",      "gris"),
            ("      Detalles completos de un archivo individual.",            "gris"),
            ("",                                                             ""),
            ("  📁  Analizar Grupo",                                         "azul"),
            ("      Compara todos los archivos de una carpeta.",              "gris"),
            ("      Detecta copias por fecha SW, hash y árbol de operaciones.","gris"),
            ("      Identifica al distribuidor original.",                    "gris"),
            ("",                                                             ""),
            ("  🕸  Red de distribución",                                    "azul"),
            ("      Grafo visual: quién le pasó el archivo a quién.",        "gris"),
            ("",                                                             ""),
            ("  " + "─" * 52,                                               "sep"),
            ("",                                                             ""),
            ("  Criterios de detección:",                                    "gris"),
            ("    ·  Misma fecha de creación SW  → mismo origen",           "gris"),
            ("    ·  Misma fecha de guardado SW  → copia directa",          "gris"),
            ("    ·  Hash SHA-256 idéntico        → copia byte a byte",     "gris"),
            ("    ·  Árbol de operaciones idéntico → mismo diseño",         "gris"),
            ("    ·  Mismo usuario en varias piezas",                       "gris"),
            ("",                                                             ""),
        ]
        for text, tag in lines:
            self._w(text + "\n", tag)

    # ─── Cierre ──────────────────────────────────────────────────────────────

    def _on_close(self) -> None:
        try:
            if self.sw_session is not None:
                self.sw_session.close()
        except Exception:
            pass
        self.root.destroy()


if __name__ == "__main__":
    app = PrivateerCAD()
    app.mainloop()
