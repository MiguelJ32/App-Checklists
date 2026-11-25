import ctypes
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import stat
import sys
import tempfile
import threading
import win32con
import win32gui
import win32print
import win32ui
import customtkinter as ctk
from PIL import Image, ImageOps, ImageWin
import pandas as pd
import pymupdf as fitz

# Configuración visual de CustomTkinter
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")


def ocultar_consola():
  """Desvincula y cierra la ventana de consola en Windows 10/11 sin activar falsos positivos."""
  if os.name == "nt":
    try:
      hwnd = ctypes.windll.kernel32.GetConsoleWindow()
      if hwnd != 0:
        ctypes.windll.user32.ShowWindow(hwnd, 0)
        ctypes.windll.kernel32.FreeConsole()
    except Exception:
      pass


def obtener_ruta_recurso(nombre_archivo):
  """Obtiene la ruta absoluta del recurso, compatible con desarrollo y PyInstaller (--onedir / --onefile)."""
  if hasattr(sys, "_MEIPASS"):
    return str(Path(sys._MEIPASS) / nombre_archivo)
  return str(Path(__file__).resolve().parent / nombre_archivo)


if getattr(sys, "frozen", False):
  BASE_DIR = Path(sys.executable).resolve().parent
else:
  BASE_DIR = Path(__file__).resolve().parent

CONFIG_FILE = BASE_DIR / "config.json"
# Hash SHA-256 correspondiente a "Mike"
PASSWORD_HASH = (
    "9dc415325a95c6e2558bf141a8772a175de49b08f0a027c8720ad942d6ec63f7"
)
ICON_PATH = obtener_ruta_recurso("logo2.ico")

DEFAULTS = {
    "RUTA_EXCEL": (
        r"J:\1 OEE\Miguel Jacinto\Programas\BuscadorCL-Windows\Relación CL.xlsx"
    ),
    "CARPETA_CHECKLISTS": r"J:\CL SAP ENSAMBLE",
    "PERMITIR_COLOR": False,
    "IMPRESORA_PREDETERMINADA": "",
}


def cargar_configuracion():
  """Carga las configuraciones desde config.json. Si no existe, lo crea con valores por defecto."""
  if not CONFIG_FILE.exists():
    guardar_configuracion(
        DEFAULTS["RUTA_EXCEL"],
        DEFAULTS["CARPETA_CHECKLISTS"],
        DEFAULTS["PERMITIR_COLOR"],
        DEFAULTS["IMPRESORA_PREDETERMINADA"],
    )
    return (
        DEFAULTS["RUTA_EXCEL"],
        DEFAULTS["CARPETA_CHECKLISTS"],
        DEFAULTS["PERMITIR_COLOR"],
        DEFAULTS["IMPRESORA_PREDETERMINADA"],
    )

  try:
    with CONFIG_FILE.open("r", encoding="utf-8") as f:
      data = json.load(f)
      return (
          data.get("RUTA_EXCEL", DEFAULTS["RUTA_EXCEL"]),
          data.get("CARPETA_CHECKLISTS", DEFAULTS["CARPETA_CHECKLISTS"]),
          data.get("PERMITIR_COLOR", DEFAULTS["PERMITIR_COLOR"]),
          data.get(
              "IMPRESORA_PREDETERMINADA", DEFAULTS["IMPRESORA_PREDETERMINADA"]
          ),
      )
  except Exception:
    return (
        DEFAULTS["RUTA_EXCEL"],
        DEFAULTS["CARPETA_CHECKLISTS"],
        DEFAULTS["PERMITIR_COLOR"],
        DEFAULTS["IMPRESORA_PREDETERMINADA"],
    )


def guardar_configuracion(
    ruta_excel,
    carpeta_checklists,
    permitir_color=False,
    impresora_predeterminada="",
):
  """Guarda la configuración de forma persistente en config.json."""
  data = {
      "RUTA_EXCEL": ruta_excel,
      "CARPETA_CHECKLISTS": carpeta_checklists,
      "PERMITIR_COLOR": permitir_color,
      "IMPRESORA_PREDETERMINADA": impresora_predeterminada,
  }
  try:
    with CONFIG_FILE.open("w", encoding="utf-8") as f:
      json.dump(data, f, indent=4, ensure_ascii=False)
  except Exception as e:
    print(f"Error al guardar la configuración: {e}")


def normalizar_ruta_larga(ruta):
  r"""Convierte una ruta a absoluta y añade el prefijo \\?\ en Windows para soportar rutas largas."""
  ruta_abs = os.path.abspath(ruta)
  if os.name == "nt" and not ruta_abs.startswith("\\\\?\\"):
    return "\\\\?\\" + ruta_abs
  return ruta_abs


def normalizar_texto_busqueda(texto):
  return re.sub(r"[^A-Za-z0-9]", "", str(texto)).upper()


def quitar_solo_lectura(ruta):
  try:
    os.chmod(ruta, stat.S_IWRITE)
  except Exception:
    pass


def obtener_impresoras_sistema():
  try:
    impresoras = [
        printer[2]
        for printer in win32print.EnumPrinters(
            win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
        )
    ]
    impresora_defecto = win32print.GetDefaultPrinter()
    if impresora_defecto in impresoras:
      impresoras.remove(impresora_defecto)
      impresoras.insert(0, impresora_defecto)
    return impresoras
  except Exception:
    return ["Impresora Predeterminada"]


def mandar_a_imprimir_configurado(lista_pdfs, nombre_impresora, a_color=True):
  """Envía cada PDF como un trabajo de impresión independiente para respetar el Dúplex por documento."""
  hprinter = win32print.OpenPrinter(nombre_impresora)
  try:
    properties = win32print.GetPrinter(hprinter, 2)
    devmode = properties["pDevMode"]

    if devmode is not None:
      devmode.Duplex = 2
      devmode.Color = 2 if a_color else 1
      devmode.Fields |= win32con.DM_DUPLEX | win32con.DM_COLOR

    hdc = win32gui.CreateDC("WINSPOOL", nombre_impresora, devmode)
    dc = win32ui.CreateDCFromHandle(hdc)

    try:
      width_px = dc.GetDeviceCaps(win32con.PHYSICALWIDTH)
      height_px = dc.GetDeviceCaps(win32con.PHYSICALHEIGHT)
      dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)

      factor_escala = 0.95
      print_w = int(width_px * factor_escala)
      print_h = int(height_px * factor_escala)
      offset_x = (width_px - print_w) // 2
      offset_y = (height_px - print_h) // 2

      dpi_impresion = min(max(dpi_x, 150), 300)

      for idx, ruta_pdf in enumerate(lista_pdfs, 1):
        ruta_limpia = ruta_pdf.replace("\\\\?\\", "")
        with fitz.open(ruta_limpia) as doc:
          dc.StartDoc(f"Checklist Documento {idx}")
          try:
            for page in doc:
              dc.StartPage()
              rect = page.rect
              es_horizontal = rect.width > rect.height

              try:
                pix = page.get_pixmap(dpi=dpi_impresion)
              except Exception:
                pix = page.get_pixmap(dpi=100)

              img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

              if es_horizontal and width_px < height_px:
                img = img.rotate(270, expand=True)
              elif not es_horizontal and width_px > height_px:
                img = img.rotate(90, expand=True)

              if not a_color:
                img = ImageOps.grayscale(img).convert("RGB")

              dib_win = ImageWin.Dib(img)
              dib_win.draw(
                  hdc,
                  (offset_x, offset_y, offset_x + print_w, offset_y + print_h),
              )
              dc.EndPage()
          finally:
            dc.EndDoc()
    finally:
      dc.DeleteDC()
  finally:
    win32print.ClosePrinter(hprinter)


# ==============================================================================
# VENTANA MODAL DE CONFIGURACIÓN
# ==============================================================================
class VentanaConfiguracionRutas(ctk.CTkToplevel):

  def __init__(
      self,
      parent,
      ruta_excel_actual,
      carpeta_pdfs_actual,
      permitir_color_actual,
  ):
    super().__init__(parent)

    self.parent = parent
    self.title("⚙ Configuración del Sistema (Administrador)")
    self.geometry("760x420")
    self.minsize(650, 380)
    self.configure(fg_color="#F1F5F9")

    if Path(ICON_PATH).exists():
      try:
        self.iconbitmap(ICON_PATH)
      except Exception:
        pass

    self.transient(parent)
    self.grab_set()

    self.ruta_excel_val = ctk.StringVar(value=ruta_excel_actual)
    self.carpeta_pdfs_val = ctk.StringVar(value=carpeta_pdfs_actual)
    self.permitir_color_val = ctk.BooleanVar(value=permitir_color_actual)

    self._crear_interfaz()

  def _crear_interfaz(self):
    # Banner modal
    header = ctk.CTkFrame(self, fg_color="#003E6B", corner_radius=0, height=54)
    header.pack(fill="x")
    header.pack_propagate(False)

    ctk.CTkLabel(
        header,
        text="⚙  Ajustes Generales del Sistema",
        font=("Roboto", 16, "bold"),
        text_color="#FFFFFF",
    ).pack(side="left", padx=20)

    body = ctk.CTkFrame(
        self,
        fg_color="#FFFFFF",
        corner_radius=10,
        border_width=1,
        border_color="#CBD5E1",
    )
    body.pack(fill="both", expand=True, padx=20, pady=18)

    # Campo 1: Ruta Excel
    frame_excel = ctk.CTkFrame(body, fg_color="transparent")
    frame_excel.pack(fill="x", padx=20, pady=(18, 8))

    ctk.CTkLabel(
        frame_excel,
        text="Ruta Archivo Excel (Relación CL.xlsx):",
        font=("Roboto", 12, "bold"),
        text_color="#1E293B",
    ).pack(anchor="w", pady=(0, 4))

    entry_excel = ctk.CTkEntry(
        frame_excel,
        textvariable=self.ruta_excel_val,
        height=38,
        border_color="#CBD5E1",
        border_width=1,
    )
    entry_excel.pack(side="left", fill="x", expand=True, padx=(0, 10))

    ctk.CTkButton(
        frame_excel,
        text="📁 Buscar...",
        width=110,
        height=38,
        fg_color="#0284C7",
        hover_color="#0369A1",
        command=self.buscar_excel,
    ).pack(side="right")

    # Campo 2: Carpeta Raíz PDFs
    frame_folder = ctk.CTkFrame(body, fg_color="transparent")
    frame_folder.pack(fill="x", padx=20, pady=10)

    ctk.CTkLabel(
        frame_folder,
        text="Carpeta Raíz de Checklists (CL SAP ENSAMBLE):",
        font=("Roboto", 12, "bold"),
        text_color="#1E293B",
    ).pack(anchor="w", pady=(0, 4))

    entry_folder = ctk.CTkEntry(
        frame_folder,
        textvariable=self.carpeta_pdfs_val,
        height=38,
        border_color="#CBD5E1",
        border_width=1,
    )
    entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 10))

    ctk.CTkButton(
        frame_folder,
        text="📂 Buscar...",
        width=110,
        height=38,
        fg_color="#0284C7",
        hover_color="#0369A1",
        command=self.buscar_carpeta,
    ).pack(side="right")

    # Campo 3: Interruptor para Habilitar/Bloquear Impresión a Color
    frame_color_opt = ctk.CTkFrame(body, fg_color="transparent")
    frame_color_opt.pack(fill="x", padx=20, pady=12)

    self.switch_color = ctk.CTkSwitch(
        frame_color_opt,
        text=" Habilitar Opción de Impresión a Color para Operadores",
        variable=self.permitir_color_val,
        font=("Roboto", 13, "bold"),
        text_color="#1E293B",
        progress_color="#16A34A",
        onvalue=True,
        offvalue=False,
    )
    self.switch_color.pack(anchor="w")

    # Botón Guardar
    ctk.CTkButton(
        body,
        text="💾 Guardar Configuración y Recargar",
        height=44,
        font=("Roboto", 14, "bold"),
        fg_color="#16A34A",
        hover_color="#15803D",
        command=self.guardar_y_cerrar,
    ).pack(fill="x", padx=20, pady=(12, 18))

  def buscar_excel(self):
    archivo = ctk.filedialog.askopenfilename(
        title="Seleccionar Archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")],
    )
    if archivo:
      self.ruta_excel_val.set(archivo)

  def buscar_carpeta(self):
    carpeta = ctk.filedialog.askdirectory(
        title="Seleccionar Carpeta Raíz de Checklists"
    )
    if carpeta:
      self.carpeta_pdfs_val.set(carpeta)

  def guardar_y_cerrar(self):
    nueva_ruta_excel = self.ruta_excel_val.get().strip()
    nueva_carpeta = self.carpeta_pdfs_val.get().strip()
    nuevo_permitir_color = self.permitir_color_val.get()
    impresora_actual = self.parent.combo_impresora.get()

    guardar_configuracion(
        nueva_ruta_excel,
        nueva_carpeta,
        nuevo_permitir_color,
        impresora_actual,
    )
    self.parent.actualizar_rutas_configuradas(
        nueva_ruta_excel, nueva_carpeta, nuevo_permitir_color
    )
    self.destroy()


# ==============================================================================
# VENTANA MODAL DE VISTA PREVIA Y SELECCIÓN DE PDFS
# ==============================================================================
class VentanaSeleccionPDFs(ctk.CTkToplevel):

  def __init__(self, parent, no_parte, lista_archivos_pdf, cache_preview=None):
    super().__init__(parent)

    self.title(f"Selección de Checklists - No. Parte: {no_parte}")
    self.geometry("1120x700")
    self.minsize(950, 560)
    self.configure(fg_color="#F1F5F9")

    if Path(ICON_PATH).exists():
      try:
        self.iconbitmap(ICON_PATH)
      except Exception:
        pass

    self.transient(parent)
    self.grab_set()

    self.no_parte = no_parte
    self.lista_archivos_pdf = lista_archivos_pdf
    self.variables_check = {}
    self.archivos_seleccionados = []
    self.cache_preview = cache_preview if cache_preview is not None else {}

    self._crear_interfaz()

  def _crear_interfaz(self):
    # Encabezado superior del modal
    header = ctk.CTkFrame(self, fg_color="#003E6B", corner_radius=0, height=54)
    header.pack(fill="x")
    header.pack_propagate(False)

    ctk.CTkLabel(
        header,
        text=f"📋  Checklists Disponibles — No. Parte: {self.no_parte}",
        font=("Roboto", 16, "bold"),
        text_color="#FFFFFF",
    ).pack(side="left", padx=20)

    ctk.CTkLabel(
        header,
        text=f"{len(self.lista_archivos_pdf)} documento(s) encontrado(s)",
        font=("Roboto", 12, "bold"),
        text_color="#BAE6FD",
    ).pack(side="right", padx=20)

    # Contenedor principal
    content = ctk.CTkFrame(self, fg_color="transparent")
    content.pack(fill="both", expand=True, padx=14, pady=14)

    frame_izq = ctk.CTkFrame(
        content,
        corner_radius=10,
        fg_color="#FFFFFF",
        border_width=1,
        border_color="#CBD5E1",
    )
    frame_izq.pack(side="left", fill="both", expand=True, padx=(0, 12))

    lbl_inst = ctk.CTkLabel(
        frame_izq,
        text="Seleccione los documentos que desea enviar a impresión:",
        font=("Roboto", 14, "bold"),
        text_color="#1E293B",
    )
    lbl_inst.pack(anchor="w", padx=14, pady=(14, 8))

    frame_acc_rapidas = ctk.CTkFrame(frame_izq, fg_color="transparent")
    frame_acc_rapidas.pack(fill="x", padx=14, pady=(0, 8))

    ctk.CTkButton(
        frame_acc_rapidas,
        text="✔ Marcar Todos",
        width=125,
        height=32,
        font=("Roboto", 12, "bold"),
        fg_color="#0284C7",
        hover_color="#0369A1",
        command=self.marcar_todos,
    ).pack(side="left", padx=(0, 8))

    ctk.CTkButton(
        frame_acc_rapidas,
        text="✖ Desmarcar Todos",
        width=140,
        height=32,
        font=("Roboto", 12, "bold"),
        fg_color="#64748B",
        hover_color="#475569",
        command=self.desmarcar_todos,
    ).pack(side="left")

    self.scroll_checkboxes = ctk.CTkScrollableFrame(
        frame_izq,
        fg_color="#F8FAFC",
        corner_radius=8,
        border_width=1,
        border_color="#E2E8F0",
    )
    self.scroll_checkboxes.pack(
        fill="both", expand=True, padx=14, pady=(4, 12)
    )

    for nombre_pdf, ruta_pdf in self.lista_archivos_pdf:
      var = ctk.BooleanVar(value=True)
      self.variables_check[ruta_pdf] = var

      frame_item = ctk.CTkFrame(
          self.scroll_checkboxes,
          fg_color="#FFFFFF",
          corner_radius=6,
          border_width=1,
          border_color="#CBD5E1",
      )
      frame_item.pack(fill="x", pady=4, padx=4)

      chk = ctk.CTkCheckBox(
          frame_item,
          text=f"  {nombre_pdf}",
          variable=var,
          font=("Roboto", 13, "bold"),
          text_color="#0F172A",
          fg_color="#16A34A",
          hover_color="#15803D",
          checkbox_width=24,
          checkbox_height=24,
      )
      chk.pack(side="left", fill="x", expand=True, padx=10, pady=8)
      chk.bind("<Enter>", lambda e, r=ruta_pdf: self.mostrar_vista_previa(r))
      frame_item.bind(
          "<Enter>", lambda e, r=ruta_pdf: self.mostrar_vista_previa(r)
      )

    ctk.CTkButton(
        frame_izq,
        text="✔ Confirmar Selección",
        height=46,
        font=("Roboto", 15, "bold"),
        fg_color="#16A34A",
        hover_color="#15803D",
        command=self.confirmar_seleccion,
    ).pack(fill="x", padx=14, pady=(0, 14))

    frame_der = ctk.CTkFrame(
        content,
        corner_radius=10,
        width=470,
        fg_color="#FFFFFF",
        border_width=1,
        border_color="#CBD5E1",
    )
    frame_der.pack(side="right", fill="both", expand=False)
    frame_der.pack_propagate(False)

    header_prev = ctk.CTkFrame(
        frame_der, fg_color="#F1F5F9", corner_radius=8, height=38
    )
    header_prev.pack(fill="x", padx=10, pady=(10, 4))
    header_prev.pack_propagate(False)

    ctk.CTkLabel(
        header_prev,
        text="👁  Vista Previa del Documento",
        font=("Roboto", 13, "bold"),
        text_color="#1E293B",
    ).pack(side="left", padx=12)

    self.lbl_imagen_prev = ctk.CTkLabel(
        frame_der,
        text="Pase el cursor sobre un checklist\npara previsualizarlo.",
        font=("Roboto", 13),
        text_color="#64748B",
    )
    self.lbl_imagen_prev.pack(fill="both", expand=True, padx=12, pady=12)

    if self.lista_archivos_pdf:
      self.mostrar_vista_previa(self.lista_archivos_pdf[0][1])

  def marcar_todos(self):
    for var in self.variables_check.values():
      var.set(True)

  def desmarcar_todos(self):
    for var in self.variables_check.values():
      var.set(False)

  def mostrar_vista_previa(self, ruta_pdf):
    if ruta_pdf in self.cache_preview:
      self.lbl_imagen_prev.configure(
          image=self.cache_preview[ruta_pdf], text=""
      )
      return

    try:
      with fitz.open(ruta_pdf) as doc:
        page = doc[0]
        pix = page.get_pixmap(dpi=130)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

      max_w, max_h = 430, 560
      img.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)

      ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
      self.cache_preview[ruta_pdf] = ctk_img
      self.lbl_imagen_prev.configure(image=ctk_img, text="")
    except Exception as e:
      self.lbl_imagen_prev.configure(
          image=None, text=f"No se pudo generar\nvista previa.\n({e})"
      )

  def confirmar_seleccion(self):
    self.archivos_seleccionados = [
        ruta for ruta, var in self.variables_check.items() if var.get()
    ]
    self.destroy()


# ==============================================================================
# APLICACIÓN PRINCIPAL
# ==============================================================================
class BuscadorChecklistsApp(ctk.CTk):

  def __init__(self):
    super().__init__()

    (
        self.ruta_excel,
        self.carpeta_raiz_pdfs,
        self.permitir_color,
        self.impresora_predeterminada,
    ) = cargar_configuracion()

    self.title("Buscador de Checklists - Piso de Producción")
    self.geometry("1180x730")
    self.minsize(1020, 640)
    self.configure(fg_color="#F1F5F9")

    if Path(ICON_PATH).exists():
      try:
        self.iconbitmap(ICON_PATH)
      except Exception:
        pass

    self.df = self.cargar_excel()
    self.pdfs_seleccionados_lista = []
    self._timer_busqueda = None
    self.cache_preview = {}
    self.indice_carpetas = []
    self._indexando = False

    self.tarjetas_creadas = []
    self.tarjeta_seleccionada = None

    self._crear_interfaz()
    self.iniciar_indexado_carpetas()

    self.bind("<Control-Shift-C>", lambda e: self.solicitar_clave_config())

  def _vincular_scroll_recursivo(self, widget):
    """Propaga el evento de rueda de ratón/trackpad desde widgets hijos al canvas scrollable."""

    def _on_mousewheel(event):
      try:
        canvas = self.scroll_coincidencias._parent_canvas
        delta = (
            int(-1 * (event.delta / 120))
            if event.delta
            else (-1 if event.num == 4 else 1)
        )
        canvas.yview_scroll(delta, "units")
      except Exception:
        pass

    widget.bind("<MouseWheel>", _on_mousewheel, add="+")
    widget.bind("<Button-4>", _on_mousewheel, add="+")
    widget.bind("<Button-5>", _on_mousewheel, add="+")
    for child in widget.winfo_children():
      self._vincular_scroll_recursivo(child)

  def iniciar_indexado_carpetas(self):
    """Indexa las subcarpetas de la unidad de red en segundo plano al iniciar."""
    if self._indexando:
      return
    self._indexando = True
    if hasattr(self, "lbl_estado_indice"):
      self.lbl_estado_indice.configure(
          text="● Indexando carpetas...", text_color="#FACC15"
      )

    def _worker_indexar():
      nuevo_indice = []
      carpeta_raiz_norm = normalizar_ruta_larga(self.carpeta_raiz_pdfs)
      try:
        for root, dirs, _ in os.walk(carpeta_raiz_norm):
          for d in dirs:
            clave = normalizar_texto_busqueda(d)
            nuevo_indice.append((clave, str(Path(root) / d)))
      except Exception as e:
        self.after(0, lambda: self.log(f"⚠️ Error al indexar carpetas: {e}"))
      finally:
        self.indice_carpetas = nuevo_indice
        self._indexando = False
        if hasattr(self, "lbl_estado_indice"):
          total = len(nuevo_indice)
          self.after(
              0,
              lambda: self.lbl_estado_indice.configure(
                  text=f"● Índice listo ({total} carpetas)",
                  text_color="#4ADE80",
              ),
          )

    threading.Thread(target=_worker_indexar, daemon=True).start()

  def cargar_excel(self):
    ruta = Path(self.ruta_excel)
    if not ruta.exists():
      print(f"❌ El archivo no existe: {ruta}")
      return pd.DataFrame()

    cols_necesarias = ["SO", "VCP", "Máquina", "No. Parte"]

    try:
      df = pd.read_excel(ruta, dtype=str, usecols=cols_necesarias)
    except (PermissionError, OSError):
      temp_excel = Path(tempfile.gettempdir()) / "temp_book1_lectura.xlsx"
      try:
        shutil.copyfile(ruta, temp_excel)
        df = pd.read_excel(temp_excel, dtype=str, usecols=cols_necesarias)
      except Exception as e:
        print(f"Error al leer copia temporal de Excel: {e}")
        return pd.DataFrame()
      finally:
        try:
          temp_excel.unlink(missing_ok=True)
        except Exception:
          pass
    except Exception as e:
      print(f"Error al leer el archivo Excel: {e}")
      return pd.DataFrame()

    df.columns = df.columns.str.strip()
    for col in df.columns:
      df[col] = df[col].fillna("").astype(str).str.strip()

    df["_search_col"] = (
        df["SO"] + " " + df["VCP"] + " " + df["Máquina"] + " " + df["No. Parte"]
    ).apply(normalizar_texto_busqueda)

    return df

  def _crear_interfaz(self):
    # ==========================================================================
    # 1. BANNER SUPERIOR CORPORATIVO (ESTILO PIZARRÓN HORA POR HORA)
    # ==========================================================================
    header_bar = ctk.CTkFrame(
        self, fg_color="#003E6B", corner_radius=0, height=64
    )
    header_bar.pack(fill="x")
    header_bar.pack_propagate(False)

    ctk.CTkLabel(
        header_bar,
        text="🏭  BUSCADOR E IMPRESIÓN DE CHECKLISTS",
        font=("Roboto", 19, "bold"),
        text_color="#FFFFFF",
    ).pack(side="left", padx=20)

    # Botón Configuración en el banner
    btn_config = ctk.CTkButton(
        header_bar,
        text="⚙️",
        width=38,
        height=34,
        corner_radius=6,
        fg_color="#002845",
        hover_color="#0284C7",
        text_color="#FFFFFF",
        font=("Roboto", 15),
        command=self.solicitar_clave_config,
    )
    btn_config.pack(side="right", padx=(8, 18))

    # Indicador de estado del índice de red
    self.lbl_estado_indice = ctk.CTkLabel(
        header_bar,
        text="● Iniciando índice...",
        font=("Roboto", 12, "bold"),
        text_color="#FACC15",
    )
    self.lbl_estado_indice.pack(side="right", padx=12)

    # ==========================================================================
    # 2. BARRA DE BÚSQUEDA
    # ==========================================================================
    frame_top = ctk.CTkFrame(
        self,
        corner_radius=10,
        fg_color="#FFFFFF",
        border_width=1,
        border_color="#CBD5E1",
    )
    frame_top.pack(fill="x", padx=16, pady=(14, 10))

    frame_search_inner = ctk.CTkFrame(frame_top, fg_color="transparent")
    frame_search_inner.pack(fill="x", padx=16, pady=12)

    ctk.CTkLabel(
        frame_search_inner,
        text="🔍  Búsqueda Rápida (SO, VCP, Máquina o No. Parte):",
        font=("Roboto", 14, "bold"),
        text_color="#0F172A",
    ).pack(anchor="w", pady=(0, 6))

    frame_entry_row = ctk.CTkFrame(frame_search_inner, fg_color="transparent")
    frame_entry_row.pack(fill="x")

    self.entry_busqueda = ctk.CTkEntry(
        frame_entry_row,
        placeholder_text="Escriba el Número de Parte, SO, VCP o Máquina...",
        height=44,
        font=("Roboto", 15),
        fg_color="#F8FAFC",
        border_color="#94A3B8",
        border_width=1,
        text_color="#0F172A",
    )
    self.entry_busqueda.pack(side="left", fill="x", expand=True, padx=(0, 10))
    self.entry_busqueda.bind("<KeyRelease>", self._on_key_release_debounce)

    ctk.CTkButton(
        frame_entry_row,
        text="Limpiar",
        width=95,
        height=44,
        font=("Roboto", 13, "bold"),
        fg_color="#64748B",
        hover_color="#475569",
        command=self.limpiar_busqueda,
    ).pack(side="right")

    # ==========================================================================
    # 3. ÁREA CENTRAL: TABLA DE RESULTADOS + PANEL DE PROCESO
    # ==========================================================================
    frame_middle = ctk.CTkFrame(self, fg_color="transparent")
    frame_middle.pack(fill="both", expand=True, padx=16, pady=(0, 10))

    # Contenedor izquierdo (Encabezados alineados + Tarjetas)
    panel_resultados = ctk.CTkFrame(
        frame_middle,
        corner_radius=10,
        fg_color="#FFFFFF",
        border_width=1,
        border_color="#CBD5E1",
    )
    panel_resultados.pack(side="left", fill="both", expand=True, padx=(0, 12))

    # Barra de encabezados de columna alineada al píxel con las tarjetas
    header_tabla = ctk.CTkFrame(
        panel_resultados, fg_color="#1F4E79", corner_radius=8, height=38
    )
    header_tabla.pack(fill="x", padx=10, pady=(10, 4))
    header_tabla.pack_propagate(False)

    header_tabla.columnconfigure(0, weight=3, uniform="col_ancho")
    header_tabla.columnconfigure(1, weight=2, uniform="col_ancho")
    header_tabla.columnconfigure(2, weight=2, uniform="col_ancho")
    header_tabla.columnconfigure(3, weight=3, uniform="col_ancho")

    encabezados = ["NO. PARTE", "SO", "VCP", "MÁQUINA"]
    for idx_col, texto_col in enumerate(encabezados):
      ctk.CTkLabel(
          header_tabla,
          text=texto_col,
          font=("Roboto", 16, "bold"),
          text_color="#FFFFFF",
          anchor="center",
      ).grid(row=0, column=idx_col, padx=5, pady=8, sticky="ew")

    self.scroll_coincidencias = ctk.CTkScrollableFrame(
        panel_resultados,
        fg_color="#F8FAFC",
        corner_radius=8,
    )
    self.scroll_coincidencias.pack(
        fill="both", expand=True, padx=10, pady=(0, 10)
    )

    # Panel derecho: Estado y bitácora del proceso
    frame_derecho = ctk.CTkFrame(
        frame_middle,
        corner_radius=10,
        width=340,
        fg_color="#FFFFFF",
        border_width=1,
        border_color="#CBD5E1",
    )
    frame_derecho.pack(side="right", fill="both", expand=False)
    frame_derecho.pack_propagate(False)

    header_log = ctk.CTkFrame(
        frame_derecho, fg_color="#1E293B", corner_radius=8, height=38
    )
    header_log.pack(fill="x", padx=10, pady=(10, 6))
    header_log.pack_propagate(False)

    ctk.CTkLabel(
        header_log,
        text="📋  Estado del Proceso",
        font=("Roboto", 13, "bold"),
        text_color="#FFFFFF",
    ).pack(side="left", padx=12)

    self.txt_log = ctk.CTkTextbox(
        frame_derecho,
        font=("Consolas", 12),
        fg_color="#F8FAFC",
        text_color="#0F172A",
        border_width=1,
        border_color="#E2E8F0",
        state="disabled",
        wrap="word",
    )
    self.txt_log.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    # ==========================================================================
    # 4. BARRA DE CONFIGURACIÓN DE IMPRESIÓN Y BOTÓN PRINCIPAL
    # ==========================================================================
    frame_options = ctk.CTkFrame(
        self,
        corner_radius=10,
        fg_color="#FFFFFF",
        border_width=1,
        border_color="#CBD5E1",
    )
    frame_options.pack(fill="x", padx=16, pady=(0, 10))

    ctk.CTkLabel(
        frame_options,
        text="🖨️ Impresora Destino:",
        font=("Roboto", 13, "bold"),
        text_color="#1E293B",
    ).pack(side="left", padx=(16, 8), pady=12)

    lista_impresoras = obtener_impresoras_sistema()
    self.combo_impresora = ctk.CTkComboBox(
        frame_options,
        values=lista_impresoras,
        width=330,
        height=36,
        font=("Roboto", 13),
        border_color="#94A3B8",
        border_width=1,
        button_color="#0284C7",
        button_hover_color="#0369A1",
        state="readonly",
        command=self.al_cambiar_impresora,
    )
    self.combo_impresora.pack(side="left", padx=5, pady=12)
    self.combo_impresora.bind(
        "<Button-1>", lambda e: self.combo_impresora._open_dropdown_menu()
    )

    if (
        self.impresora_predeterminada
        and self.impresora_predeterminada in lista_impresoras
    ):
      self.combo_impresora.set(self.impresora_predeterminada)

    self.lbl_color = ctk.CTkLabel(
        frame_options,
        text="🎨 Modo:",
        font=("Roboto", 13, "bold"),
        text_color="#1E293B",
    )

    self.combo_color = ctk.CTkComboBox(
        frame_options,
        values=["Blanco y Negro", "Color"],
        width=180,
        height=36,
        font=("Roboto", 13),
        border_color="#94A3B8",
        border_width=1,
        button_color="#0284C7",
        button_hover_color="#0369A1",
        state="readonly",
    )
    self.combo_color.set("Blanco y Negro")
    self.combo_color.bind(
        "<Button-1>", lambda e: self.combo_color._open_dropdown_menu()
    )

    self.actualizar_estado_selector_color()

    frame_bottom = ctk.CTkFrame(self, fg_color="transparent")
    frame_bottom.pack(fill="x", padx=16, pady=(0, 16))

    self.btn_imprimir = ctk.CTkButton(
        frame_bottom,
        text="🖨️  Mandar a Imprimir Checklists Seleccionados",
        height=52,
        corner_radius=10,
        font=("Roboto", 18, "bold"),
        fg_color="#16A34A",
        hover_color="#15803D",
        state="disabled",
        command=self.imprimir_pdfs,
    )
    self.btn_imprimir.pack(fill="x")

    self.ejecutar_filtrado()

  def limpiar_busqueda(self):
    self.entry_busqueda.delete(0, "end")
    self.ejecutar_filtrado()
    self.entry_busqueda.focus()

  def actualizar_estado_selector_color(self):
    if not self.permitir_color:
      self.lbl_color.pack_forget()
      self.combo_color.pack_forget()
      self.combo_color.set("Blanco y Negro")
    else:
      self.lbl_color.pack(side="left", padx=(24, 8), pady=12)
      self.combo_color.pack(side="left", padx=5, pady=12)
      self.combo_color.set("Blanco y Negro")

  def al_cambiar_impresora(self, impresora_seleccionada):
    guardar_configuracion(
        self.ruta_excel,
        self.carpeta_raiz_pdfs,
        self.permitir_color,
        impresora_seleccionada,
    )

  def solicitar_clave_config(self):
    dialog = ctk.CTkInputDialog(
        text="Ingrese la contraseña de administrador:",
        title="Acceso Protegido",
    )
    clave_ingresada = dialog.get_input()

    if clave_ingresada is None:
      return

    hash_ingresado = hashlib.sha256(
        clave_ingresada.encode("utf-8")
    ).hexdigest()
    if hash_ingresado == PASSWORD_HASH:
      VentanaConfiguracionRutas(
          self,
          self.ruta_excel,
          self.carpeta_raiz_pdfs,
          self.permitir_color,
      )
    else:
      self.limpiar_log()
      self.log("❌ Contraseña incorrecta.")

  def actualizar_rutas_configuradas(
      self, nueva_ruta_excel, nueva_carpeta_pdfs, nuevo_permitir_color
  ):
    self.ruta_excel = nueva_ruta_excel
    self.carpeta_raiz_pdfs = nueva_carpeta_pdfs
    self.permitir_color = nuevo_permitir_color
    self.cache_preview.clear()

    self.actualizar_estado_selector_color()

    self.limpiar_log()
    self.log("⚙ Configuración actualizada correctamente.")
    self.log(f" 📄 Excel: {self.ruta_excel}")
    self.log(f" 📁 Raíz PDFs: {self.carpeta_raiz_pdfs}")
    self.log(
        " 🎨 Modo Color:"
        f" {'Habilitado' if self.permitir_color else 'Oculto (Solo B&N)'}\n"
    )

    self.df = self.cargar_excel()
    self.iniciar_indexado_carpetas()
    self.ejecutar_filtrado()

  def _on_key_release_debounce(self, event=None):
    if self._timer_busqueda is not None:
      self.after_cancel(self._timer_busqueda)
    self._timer_busqueda = self.after(300, self.ejecutar_filtrado)

  def log(self, mensaje):
    self.txt_log.configure(state="normal")
    self.txt_log.insert("end", mensaje + "\n")
    self.txt_log.see("end")
    self.txt_log.configure(state="disabled")

  def limpiar_log(self):
    self.txt_log.configure(state="normal")
    self.txt_log.delete("1.0", "end")
    self.txt_log.configure(state="disabled")

  def ejecutar_filtrado(self):
    termino_raw = self.entry_busqueda.get().strip()
    termino_norm = normalizar_texto_busqueda(termino_raw)

    for child in self.scroll_coincidencias.winfo_children():
      child.destroy()

    self.tarjetas_creadas.clear()
    self.tarjeta_seleccionada = None

    try:
      self.scroll_coincidencias._parent_canvas.yview_moveto(0.0)
    except Exception:
      pass

    if self.df.empty:
      ctk.CTkLabel(
          self.scroll_coincidencias,
          text="⚠️ No se pudo cargar el archivo Excel.\nVerifique la ruta en ⚙️.",
          font=("Roboto", 14, "bold"),
          text_color="#DC2626",
      ).pack(pady=40)
      return

    if not termino_norm:
      ctk.CTkLabel(
          self.scroll_coincidencias,
          text="🔎 Ingrese un término en el buscador para consultar registros.",
          font=("Roboto", 14),
          text_color="#64748B",
      ).pack(pady=50)
      return

    mask = self.df["_search_col"].str.contains(
        termino_norm, regex=False, na=False
    )
    df_filtrado = self.df[mask]

    if df_filtrado.empty:
      ctk.CTkLabel(
          self.scroll_coincidencias,
          text="Sin coincidencias para la búsqueda.",
          font=("Roboto", 14),
          text_color="#64748B",
      ).pack(pady=50)
      return

    registros_unicos = df_filtrado.drop_duplicates(subset=["No. Parte"]).head(
        50
    )

    for _, row in registros_unicos.iterrows():
      no_parte = row.get("No. Parte", "N/A")
      so = row.get("SO", "N/A")
      vcp = row.get("VCP", "N/A")
      maquina = row.get("Máquina", "N/A")

      card = ctk.CTkFrame(
          self.scroll_coincidencias,
          corner_radius=8,
          border_width=1,
          border_color="#CBD5E1",
          fg_color="#FFFFFF",
          cursor="hand2",
      )
      card.pack(fill="x", pady=4, padx=4)

      card.columnconfigure(0, weight=3, uniform="col_ancho")
      card.columnconfigure(1, weight=2, uniform="col_ancho")
      card.columnconfigure(2, weight=2, uniform="col_ancho")
      card.columnconfigure(3, weight=3, uniform="col_ancho")

      lbl_np = ctk.CTkLabel(
          card,
          text=f"{no_parte}",
          font=("Roboto", 17, "bold"),
          text_color="#0F172A",
          anchor="center",
      )
      lbl_np.grid(row=0, column=0, padx=5, pady=12, sticky="ew")

      lbl_so = ctk.CTkLabel(
          card,
          text=f"{so}",
          font=("Roboto", 16, "bold"),
          text_color="#475569",
          anchor="center",
      )
      lbl_so.grid(row=0, column=1, padx=5, pady=12, sticky="ew")

      lbl_vcp = ctk.CTkLabel(
          card,
          text=f"{vcp}",
          font=("Roboto", 16, "bold"),
          text_color="#475569",
          anchor="center",
      )
      lbl_vcp.grid(row=0, column=2, padx=5, pady=12, sticky="ew")

      lbl_maq = ctk.CTkLabel(
          card,
          text=f"{maquina}",
          font=("Roboto", 16, "bold"),
          text_color="#0284C7",
          anchor="center",
      )
      lbl_maq.grid(row=0, column=3, padx=5, pady=12, sticky="ew")

      widgets_tarjeta = [card, lbl_np, lbl_so, lbl_vcp, lbl_maq]
      self.tarjetas_creadas.append(card)

      def _on_enter(e, c=card):
        if self.tarjeta_seleccionada != c:
          c.configure(fg_color="#F1F5F9", border_color="#94A3B8")

      def _on_leave(e, c=card):
        if self.tarjeta_seleccionada != c:
          c.configure(fg_color="#FFFFFF", border_color="#CBD5E1")

      for w in widgets_tarjeta:
        w.bind(
            "<Button-1>",
            lambda e, c=card, np=no_parte: self.seleccionar_tarjeta(c, np),
        )
        w.bind("<Enter>", _on_enter)
        w.bind("<Leave>", _on_leave)

      self._vincular_scroll_recursivo(card)

  def seleccionar_tarjeta(self, tarjeta_cliqueada, no_parte):
    for t in self.tarjetas_creadas:
      t.configure(border_color="#CBD5E1", border_width=1, fg_color="#FFFFFF")

    tarjeta_cliqueada.configure(
        border_color="#0284C7", border_width=3, fg_color="#E0F2FE"
    )
    self.tarjeta_seleccionada = tarjeta_cliqueada

    self.procesar_un_solo_no_parte(no_parte)

  def procesar_un_solo_no_parte(self, no_parte_seleccionado):
    self.limpiar_log()
    self.btn_imprimir.configure(state="disabled")

    self.pdfs_seleccionados_lista = []
    self.log(f"🎯 No. Parte: '{no_parte_seleccionado}'\n")

    no_parte_norm = normalizar_texto_busqueda(no_parte_seleccionado)
    carpetas_coincidentes = []

    if self.indice_carpetas:
      rutas_vistas = set()
      for clave_carpeta, ruta_completa in self.indice_carpetas:
        if no_parte_norm in clave_carpeta and ruta_completa not in rutas_vistas:
          rutas_vistas.add(ruta_completa)
          carpetas_coincidentes.append(ruta_completa)
    else:
      carpeta_raiz_norm = normalizar_ruta_larga(self.carpeta_raiz_pdfs)
      for root, dirs, _ in os.walk(carpeta_raiz_norm):
        for d in dirs:
          if no_parte_norm in normalizar_texto_busqueda(d):
            ruta_completa = str(Path(root) / d)
            if ruta_completa not in carpetas_coincidentes:
              carpetas_coincidentes.append(ruta_completa)

    if not carpetas_coincidentes:
      self.log(
          f"❌ No se encontró ninguna carpeta física para el No. Parte"
          f" '{no_parte_seleccionado}'."
      )
      return

    lista_pdfs_disponibles = []
    rutas_procesadas = set()

    for ruta_folder in carpetas_coincidentes:
      try:
        for f in os.listdir(ruta_folder):
          if f.lower().endswith(".pdf") and f.lower() not in rutas_procesadas:
            rutas_procesadas.add(f.lower())
            ruta_pdf = str(Path(ruta_folder) / f)
            quitar_solo_lectura(ruta_pdf)
            lista_pdfs_disponibles.append((f, ruta_pdf))
      except Exception as e:
        self.log(f"⚠️ No se pudo leer carpeta {ruta_folder}: {e}")

    if not lista_pdfs_disponibles:
      self.log("\n❌ Carpeta localizada pero no contiene ningún PDF.")
      return

    ventana_modal = VentanaSeleccionPDFs(
        self,
        no_parte_seleccionado,
        lista_pdfs_disponibles,
        cache_preview=self.cache_preview,
    )
    self.wait_window(ventana_modal)

    pdfs_a_imprimir = ventana_modal.archivos_seleccionados

    if not pdfs_a_imprimir:
      self.log("⚠️ Operación cancelada o no se seleccionó ningún PDF.")
      return

    self.pdfs_seleccionados_lista = pdfs_a_imprimir
    self.log(
        f"🎉 {len(pdfs_a_imprimir)} checklist(s) listos para impresión"
        " individual."
    )
    self.log("✔ Presione el botón verde para enviar a la cola de impresión.")
    self.btn_imprimir.configure(state="normal")

  def imprimir_pdfs(self):
    if not self.pdfs_seleccionados_lista:
      self.log("❌ No hay ningún paquete preparado.")
      return

    impresora_sel = self.combo_impresora.get()
    modo_color = (
        self.combo_color.get() if self.permitir_color else "Blanco y Negro"
    )
    es_color = modo_color == "Color"
    pdfs_a_enviar = list(self.pdfs_seleccionados_lista)

    self.btn_imprimir.configure(state="disabled", text="⏳  Imprimiendo...")
    self.log(f"\n🖨 Enviando a: [{impresora_sel}]")
    self.log(f"   • Documentos independientes: {len(pdfs_a_enviar)}")
    self.log(
        "   • Modo: Dúplex inteligente por documento |"
        f' {"Color" if es_color else "Blanco y Negro"}'
    )

    def _worker_impresion():
      try:
        mandar_a_imprimir_configurado(
            pdfs_a_enviar, impresora_sel, a_color=es_color
        )
        self.after(
            0,
            lambda: self.log(
                "✔ Todos los documentos fueron enviados exitosamente a la cola"
                " de Windows."
            ),
        )
      except Exception as e:
        err_msg = str(e)
        self.after(
            0, lambda: self.log(f"❌ Error al enviar a la impresora: {err_msg}")
        )
      finally:
        self.after(
            0,
            lambda: self.btn_imprimir.configure(
                state="normal",
                text="🖨️  Mandar a Imprimir Checklists Seleccionados",
            ),
        )

    threading.Thread(target=_worker_impresion, daemon=True).start()


if __name__ == "__main__":
  app = BuscadorChecklistsApp()
  app.mainloop()
