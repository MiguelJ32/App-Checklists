import ctypes
import json
import os
import re
import shutil
import stat
import sys
import tempfile
import win32api
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


def obtener_ruta_recurso(nombre_archivo):
  """Obtiene la ruta absoluta del recurso, compatible con desarrollo y PyInstaller (--onedir / --onefile)."""
  if hasattr(sys, "_MEIPASS"):
    return os.path.join(sys._MEIPASS, nombre_archivo)
  return os.path.join(os.path.dirname(os.path.abspath(__file__)), nombre_archivo)


CONFIG_FILE = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "config.json"
)
PASSWORD_CONFIG = "Mike"
ICON_PATH = obtener_ruta_recurso("logo2.ico")

DEFAULTS = {
    "RUTA_EXCEL": r"J:\1 OEE\Miguel Jacinto\Programas\BuscadorCL-Windows\Relación CL.xlsx",
    "CARPETA_CHECKLISTS": r"J:\CL SAP ENSAMBLE",
    "PERMITIR_COLOR": False,  # Por defecto bloqueado a solo Blanco y Negro
    "IMPRESORA_PREDETERMINADA": "",
}


def cargar_configuracion():
  """Carga las configuraciones desde config.json. Si no existe, lo crea con valores por defecto."""
  if not os.path.exists(CONFIG_FILE):
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
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
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
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
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
      # 2 = DMDUP_VERTICAL (Dúplex por el borde largo / ambos lados si el doc tiene 2+ páginas)
      devmode.Duplex = 2
      devmode.Color = 2 if a_color else 1  # 2 = Color, 1 = B&N
      devmode.Fields |= win32con.DM_DUPLEX | win32con.DM_COLOR

    hdc = win32gui.CreateDC("WINSPOOL", nombre_impresora, devmode)
    dc = win32ui.CreateDCFromHandle(hdc)

    width_px = dc.GetDeviceCaps(win32con.PHYSICALWIDTH)
    height_px = dc.GetDeviceCaps(win32con.PHYSICALHEIGHT)
    dpi_x = dc.GetDeviceCaps(win32con.LOGPIXELSX)

    # Margen de seguridad del 2.5% por lado (Escala al 95% centrada)
    factor_escala = 0.95
    print_w = int(width_px * factor_escala)
    print_h = int(height_px * factor_escala)
    offset_x = (width_px - print_w) // 2
    offset_y = (height_px - print_h) // 2

    # Procesar cada archivo PDF de manera individual
    for idx, ruta_pdf in enumerate(lista_pdfs, 1):
      ruta_limpia = ruta_pdf.replace("\\\\?\\", "")
      doc = fitz.open(ruta_limpia)

      # Iniciar un trabajo de impresión independiente en la cola de Windows
      dc.StartDoc(f"Checklist Documento {idx}")

      for page_num in range(len(doc)):
        dc.StartPage()
        page = doc[page_num]

        rect = page.rect
        es_horizontal = rect.width > rect.height

        pix = page.get_pixmap(dpi=max(dpi_x, 150))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

        # Ajuste automático de orientación (Vertical vs Horizontal)
        if es_horizontal and width_px < height_px:
          img = img.rotate(270, expand=True)
        elif not es_horizontal and width_px > height_px:
          img = img.rotate(90, expand=True)

        if not a_color:
          img = ImageOps.grayscale(img).convert("RGB")

        dib_win = ImageWin.Dib(img)
        # Dibujar la imagen ajustada dentro del área con margen
        dib_win.draw(
            hdc, (offset_x, offset_y, offset_x + print_w, offset_y + print_h)
        )

        dc.EndPage()

      doc.close()
      dc.EndDoc()  # Finaliza el trabajo actual antes de pasar al siguiente PDF

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
    self.geometry("750x390")
    self.minsize(650, 350)

    if os.path.exists(ICON_PATH):
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
    lbl_title = ctk.CTkLabel(
        self,
        text="Ajustes Generales del Sistema",
        font=("Helvetica", 14, "bold"),
    )
    lbl_title.pack(anchor="w", padx=20, pady=(15, 10))

    # Campo 1: Ruta Excel
    frame_excel = ctk.CTkFrame(self, fg_color="transparent")
    frame_excel.pack(fill="x", padx=20, pady=5)

    lbl_excel = ctk.CTkLabel(
        frame_excel,
        text="Ruta Archivo Excel (Relación CL.xlsx):",
        font=("Helvetica", 11, "bold"),
    )
    lbl_excel.pack(anchor="w")

    entry_excel = ctk.CTkEntry(
        frame_excel, textvariable=self.ruta_excel_val, height=35
    )
    entry_excel.pack(side="left", fill="x", expand=True, padx=(0, 10))

    btn_explorar_excel = ctk.CTkButton(
        frame_excel,
        text="📁 Buscar...",
        width=100,
        height=35,
        command=self.buscar_excel,
    )
    btn_explorar_excel.pack(side="right")

    # Campo 2: Carpeta Raíz PDFs
    frame_folder = ctk.CTkFrame(self, fg_color="transparent")
    frame_folder.pack(fill="x", padx=20, pady=10)

    lbl_folder = ctk.CTkLabel(
        frame_folder,
        text="Carpeta Raíz de Checklists (CL SAP ENSAMBLE):",
        font=("Helvetica", 11, "bold"),
    )
    lbl_folder.pack(anchor="w")

    entry_folder = ctk.CTkEntry(
        frame_folder, textvariable=self.carpeta_pdfs_val, height=35
    )
    entry_folder.pack(side="left", fill="x", expand=True, padx=(0, 10))

    btn_explorar_folder = ctk.CTkButton(
        frame_folder,
        text="📂 Buscar...",
        width=100,
        height=35,
        command=self.buscar_carpeta,
    )
    btn_explorar_folder.pack(side="right")

    # Campo 3: Interruptor para Habilitar/Bloquear Impresión a Color
    frame_color_opt = ctk.CTkFrame(self, fg_color="transparent")
    frame_color_opt.pack(fill="x", padx=20, pady=10)

    self.switch_color = ctk.CTkSwitch(
        frame_color_opt,
        text=" Habilitar Opción de Impresión a Color para Operadores",
        variable=self.permitir_color_val,
        font=("Helvetica", 12, "bold"),
        onvalue=True,
        offvalue=False,
    )
    self.switch_color.pack(anchor="w")

    # Botón Guardar
    btn_guardar = ctk.CTkButton(
        self,
        text="💾 Guardar Configuración y Recargar",
        height=42,
        font=("Helvetica", 13, "bold"),
        fg_color="#2E7D32",
        hover_color="#1B5E20",
        command=self.guardar_y_cerrar,
    )
    btn_guardar.pack(fill="x", padx=20, pady=(15, 10))

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

  def __init__(self, parent, no_parte, lista_archivos_pdf):
    super().__init__(parent)

    self.title(f"Selección de Checklists - No. Parte: {no_parte}")
    self.geometry("1100x680")
    self.minsize(950, 550)

    if os.path.exists(ICON_PATH):
      try:
        self.iconbitmap(ICON_PATH)
      except Exception:
        pass

    self.transient(parent)
    self.grab_set()

    self.lista_archivos_pdf = lista_archivos_pdf
    self.variables_check = {}
    self.archivos_seleccionados = []

    self._crear_interfaz()

  def _crear_interfaz(self):
    frame_izq = ctk.CTkFrame(self, corner_radius=8)
    frame_izq.pack(side="left", fill="both", expand=True, padx=12, pady=12)

    lbl_inst = ctk.CTkLabel(
        frame_izq,
        text="Marque los checklists que desea incluir:",
        font=("Helvetica", 13, "bold"),
    )
    lbl_inst.pack(anchor="w", padx=10, pady=(10, 5))

    frame_acc_rapidas = ctk.CTkFrame(frame_izq, fg_color="transparent")
    frame_acc_rapidas.pack(fill="x", padx=10, pady=(0, 5))

    btn_todos = ctk.CTkButton(
        frame_acc_rapidas,
        text="Marcar Todos",
        width=110,
        height=28,
        font=("Helvetica", 11),
        command=self.marcar_todos,
    )
    btn_todos.pack(side="left", padx=(0, 5))

    btn_ninguno = ctk.CTkButton(
        frame_acc_rapidas,
        text="Desmarcar Todos",
        width=120,
        height=28,
        font=("Helvetica", 11),
        fg_color="#64748B",
        hover_color="#475569",
        command=self.desmarcar_todos,
    )
    btn_ninguno.pack(side="left")

    self.scroll_checkboxes = ctk.CTkScrollableFrame(frame_izq)
    self.scroll_checkboxes.pack(
        fill="both", expand=True, padx=10, pady=(5, 10)
    )

    for nombre_pdf, ruta_pdf in self.lista_archivos_pdf:
      var = ctk.BooleanVar(value=True)
      self.variables_check[ruta_pdf] = var

      frame_item = ctk.CTkFrame(self.scroll_checkboxes, fg_color="transparent")
      frame_item.pack(fill="x", pady=3)

      chk = ctk.CTkCheckBox(
          frame_item,
          text=nombre_pdf,
          variable=var,
          font=("Helvetica", 12),
          checkbox_width=22,
          checkbox_height=22,
      )
      chk.pack(side="left", fill="x", expand=True)
      chk.bind("<Enter>", lambda e, r=ruta_pdf: self.mostrar_vista_previa(r))

    btn_confirmar = ctk.CTkButton(
        frame_izq,
        text="✔ Confirmar Selección",
        height=42,
        font=("Helvetica", 14, "bold"),
        fg_color="#2E7D32",
        hover_color="#1B5E20",
        command=self.confirmar_seleccion,
    )
    btn_confirmar.pack(fill="x", padx=10, pady=(0, 10))

    frame_der = ctk.CTkFrame(self, corner_radius=8, width=460)
    frame_der.pack(
        side="right", fill="both", expand=False, padx=(0, 12), pady=12
    )
    frame_der.pack_propagate(False)

    lbl_prev_title = ctk.CTkLabel(
        frame_der, text="👁 Vista Previa", font=("Helvetica", 13, "bold")
    )
    lbl_prev_title.pack(anchor="w", padx=10, pady=(10, 5))

    self.lbl_imagen_prev = ctk.CTkLabel(
        frame_der,
        text="Pase el cursor sobre un checklist\npara previsualizarlo.",
        font=("Helvetica", 12),
        text_color="#64748B",
    )
    self.lbl_imagen_prev.pack(fill="both", expand=True, padx=10, pady=10)

    if self.lista_archivos_pdf:
      self.mostrar_vista_previa(self.lista_archivos_pdf[0][1])

  def marcar_todos(self):
    for var in self.variables_check.values():
      var.set(True)

  def desmarcar_todos(self):
    for var in self.variables_check.values():
      var.set(False)

  def mostrar_vista_previa(self, ruta_pdf):
    try:
      doc = fitz.open(ruta_pdf)
      page = doc[0]
      pix = page.get_pixmap(dpi=130)

      img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
      max_w, max_h = 420, 580
      img.thumbnail((max_w, max_h), Image.Resampling.LANCZOS)

      ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=img.size)
      self.lbl_imagen_prev.configure(image=ctk_img, text="")
      doc.close()
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

    self.title("Buscador de Checklists por Excel / No. Parte")
    self.geometry("1150x700")
    self.minsize(1000, 620)

    if os.path.exists(ICON_PATH):
      try:
        self.iconbitmap(ICON_PATH)
      except Exception:
        pass

    self.df = self.cargar_excel()
    self.pdfs_seleccionados_lista = []
    self._timer_busqueda = None

    self.tarjetas_creadas = []
    self.tarjeta_seleccionada = None

    self._crear_interfaz()

    self.bind("<Control-Shift-C>", lambda e: self.solicitar_clave_config())

  def cargar_excel(self):
    try:
      if not os.path.exists(self.ruta_excel):
        print(f"❌ El archivo no existe: {self.ruta_excel}")
        return pd.DataFrame()

      temp_dir = tempfile.gettempdir()
      temp_excel = os.path.join(temp_dir, "temp_book1_lectura.xlsx")
      shutil.copyfile(self.ruta_excel, temp_excel)

      df = pd.read_excel(temp_excel, dtype=str)
      df.columns = df.columns.str.strip()
      for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

      df["_search_col"] = (
          df["SO"].fillna("")
          + " "
          + df["VCP"].fillna("")
          + " "
          + df["Máquina"].fillna("")
          + " "
          + df["No. Parte"].fillna("")
      ).apply(normalizar_texto_busqueda)

      try:
        os.remove(temp_excel)
      except Exception:
        pass

      return df
    except Exception as e:
      print(f"Error al leer el archivo Excel: {e}")
      return pd.DataFrame()

  def _crear_interfaz(self):
    frame_top = ctk.CTkFrame(self, corner_radius=8)
    frame_top.pack(fill="x", padx=15, pady=(15, 10))

    frame_lbl = ctk.CTkFrame(frame_top, fg_color="transparent")
    frame_lbl.pack(fill="x", padx=15, pady=(8, 2))

    lbl_instruccion = ctk.CTkLabel(
        frame_lbl,
        text="🔍 Buscar por SO, VCP, Máquina o No. Parte:",
        font=("Helvetica", 14, "bold"),
    )
    lbl_instruccion.pack(side="left")

    btn_config = ctk.CTkButton(
        frame_lbl,
        text="⚙️",
        width=32,
        height=28,
        fg_color="transparent",
        hover_color="#E2E8F0",
        text_color="#64748B",
        command=self.solicitar_clave_config,
    )
    btn_config.pack(side="right")

    self.entry_busqueda = ctk.CTkEntry(
        frame_top,
        placeholder_text="Escriba para buscar...",
        height=40,
        font=("Helvetica", 14),
    )
    self.entry_busqueda.pack(fill="x", padx=15, pady=(0, 12))
    self.entry_busqueda.bind("<KeyRelease>", self._on_key_release_debounce)

    frame_middle = ctk.CTkFrame(self, fg_color="transparent")
    frame_middle.pack(fill="both", expand=True, padx=15, pady=(0, 10))

    self.scroll_coincidencias = ctk.CTkScrollableFrame(
        frame_middle, label_text="Seleccione un Número de Parte", width=680
    )
    self.scroll_coincidencias.pack(
        side="left", fill="both", expand=True, padx=(0, 10)
    )

    frame_derecho = ctk.CTkFrame(frame_middle, corner_radius=8, width=320)
    frame_derecho.pack(side="right", fill="both", expand=False)
    frame_derecho.pack_propagate(False)

    lbl_log_titulo = ctk.CTkLabel(
        frame_derecho,
        text="📋 Proceso",
        font=("Helvetica", 13, "bold"),
    )
    lbl_log_titulo.pack(anchor="w", padx=12, pady=(10, 5))

    self.txt_log = ctk.CTkTextbox(
        frame_derecho, font=("Consolas", 11), state="disabled", wrap="word"
    )
    self.txt_log.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    frame_options = ctk.CTkFrame(self, corner_radius=8)
    frame_options.pack(fill="x", padx=15, pady=(0, 10))

    lbl_printer = ctk.CTkLabel(
        frame_options, text="🖨️ Impresora:", font=("Helvetica", 12, "bold")
    )
    lbl_printer.pack(side="left", padx=(15, 5), pady=10)

    lista_impresoras = obtener_impresoras_sistema()
    self.combo_impresora = ctk.CTkComboBox(
        frame_options,
        values=lista_impresoras,
        width=300,
        font=("Helvetica", 12),
        state="readonly",
        command=self.al_cambiar_impresora,
    )
    self.combo_impresora.pack(side="left", padx=5, pady=10)
    self.combo_impresora.bind(
        "<Button-1>", lambda e: self.combo_impresora._open_dropdown_menu()
    )

    if (
        self.impresora_predeterminada
        and self.impresora_predeterminada in lista_impresoras
    ):
      self.combo_impresora.set(self.impresora_predeterminada)

    # --- Selector de Color dinámico ---
    self.lbl_color = ctk.CTkLabel(
        frame_options, text="🎨 Color:", font=("Helvetica", 12, "bold")
    )

    self.combo_color = ctk.CTkComboBox(
        frame_options,
        values=["Blanco y Negro", "Color"],
        width=180,
        font=("Helvetica", 12),
        state="readonly",
    )
    self.combo_color.set("Blanco y Negro")
    self.combo_color.bind(
        "<Button-1>", lambda e: self.combo_color._open_dropdown_menu()
    )

    # Mostrar u ocultar el campo de color según config.json
    self.actualizar_estado_selector_color()

    frame_bottom = ctk.CTkFrame(self, fg_color="transparent")
    frame_bottom.pack(fill="x", padx=15, pady=(0, 15))

    self.btn_imprimir = ctk.CTkButton(
        frame_bottom,
        text="🖨️ Mandar a Imprimir Check Lists Seleccionados",
        height=48,
        font=("Helvetica", 18, "bold"),
        fg_color="#2E7D32",
        hover_color="#1B5E20",
        state="disabled",
        command=self.imprimir_pdfs,
    )
    self.btn_imprimir.pack(fill="x")

    self.ejecutar_filtrado()

  def actualizar_estado_selector_color(self):
    """Muestra u oculta dinámicamente el campo de selección de color según config.json."""
    if not self.permitir_color:
      self.lbl_color.pack_forget()
      self.combo_color.pack_forget()
      self.combo_color.set("Blanco y Negro")
    else:
      self.lbl_color.pack(side="left", padx=(20, 5), pady=10)
      self.combo_color.pack(side="left", padx=5, pady=10)
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

    if clave_ingresada == PASSWORD_CONFIG:
      VentanaConfiguracionRutas(
          self,
          self.ruta_excel,
          self.carpeta_raiz_pdfs,
          self.permitir_color,
      )
    elif clave_ingresada is not None:
      self.limpiar_log()
      self.log("❌ Contraseña incorrecta.")

  def actualizar_rutas_configuradas(
      self, nueva_ruta_excel, nueva_carpeta_pdfs, nuevo_permitir_color
  ):
    self.ruta_excel = nueva_ruta_excel
    self.carpeta_raiz_pdfs = nueva_carpeta_pdfs
    self.permitir_color = nuevo_permitir_color

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
          text="No se pudo cargar el archivo Excel.\nVerifique la ruta en ⚙️.",
      ).pack(pady=20)
      return

    if not termino_norm:
      ctk.CTkLabel(
          self.scroll_coincidencias,
          text="Ingrese un término en el buscador para consultar.",
          font=("Helvetica", 12),
          text_color="#64748B",
      ).pack(pady=40)
      return

    df_filtrado = self.df[
        self.df["_search_col"].str.contains(termino_norm, na=False)
    ]

    if df_filtrado.empty:
      ctk.CTkLabel(
          self.scroll_coincidencias,
          text="Sin coincidencias para la búsqueda.",
          font=("Helvetica", 12),
          text_color="#64748B",
      ).pack(pady=40)
      return

    registros_unicos = df_filtrado.drop_duplicates(subset=["No. Parte"])

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
      card.pack(fill="x", pady=5, padx=5)

      card.columnconfigure(0, weight=3, uniform="col_ancho")
      card.columnconfigure(1, weight=2, uniform="col_ancho")
      card.columnconfigure(2, weight=2, uniform="col_ancho")
      card.columnconfigure(3, weight=3, uniform="col_ancho")

      lbl_np = ctk.CTkLabel(
          card,
          text=f"{no_parte}",
          font=("Helvetica", 18, "bold"),
          text_color="#0F172A",
          anchor="center",
      )
      lbl_np.grid(row=0, column=0, padx=5, pady=12, sticky="ew")

      lbl_so = ctk.CTkLabel(
          card,
          text=f"{so}",
          font=("Helvetica", 18, "bold"),
          text_color="#475569",
          anchor="center",
      )
      lbl_so.grid(row=0, column=1, padx=5, pady=12, sticky="ew")

      lbl_vcp = ctk.CTkLabel(
          card,
          text=f"{vcp}",
          font=("Helvetica", 18, "bold"),
          text_color="#475569",
          anchor="center",
      )
      lbl_vcp.grid(row=0, column=2, padx=5, pady=12, sticky="ew")

      lbl_maq = ctk.CTkLabel(
          card,
          text=f"{maquina}",
          font=("Helvetica", 18, "bold"),
          text_color="#1F4E79",
          anchor="center",
      )
      lbl_maq.grid(row=0, column=3, padx=5, pady=12, sticky="ew")

      widgets_tarjeta = [card, lbl_np, lbl_so, lbl_vcp, lbl_maq]
      self.tarjetas_creadas.append(card)

      for w in widgets_tarjeta:
        w.bind(
            "<Button-1>",
            lambda e, c=card, np=no_parte: self.seleccionar_tarjeta(c, np),
        )

  def seleccionar_tarjeta(self, tarjeta_cliqueada, no_parte):
    for t in self.tarjetas_creadas:
      t.configure(border_color="#CBD5E1", border_width=1, fg_color="#FFFFFF")

    tarjeta_cliqueada.configure(
        border_color="#1F4E79", border_width=3, fg_color="#F0F4F8"
    )
    self.tarjeta_seleccionada = tarjeta_cliqueada

    self.procesar_un_solo_no_parte(no_parte)

  def procesar_un_solo_no_parte(self, no_parte_seleccionado):
    self.limpiar_log()
    self.btn_imprimir.configure(state="disabled")

    self.pdfs_seleccionados_lista = []
    self.log(f"🎯 No. Parte: '{no_parte_seleccionado}'\n")

    no_parte_norm = normalizar_texto_busqueda(no_parte_seleccionado)
    carpeta_raiz_norm = normalizar_ruta_larga(self.carpeta_raiz_pdfs)

    carpetas_coincidentes = []
    for root, dirs, _ in os.walk(carpeta_raiz_norm):
      for d in dirs:
        nombre_carpeta_norm = normalizar_texto_busqueda(d)
        if no_parte_norm in nombre_carpeta_norm:
          ruta_carpeta_completa = os.path.join(root, d)
          if ruta_carpeta_completa not in [r[1] for r in carpetas_coincidentes]:
            carpetas_coincidentes.append((d, ruta_carpeta_completa))

    if not carpetas_coincidentes:
      self.log(
          f"❌ No se encontró ninguna carpeta física para el No. Parte"
          f" '{no_parte_seleccionado}'."
      )
      return

    lista_pdfs_disponibles = []
    rutas_procesadas = set()

    for nombre_folder, ruta_folder in carpetas_coincidentes:
      archivos = [
          f for f in os.listdir(ruta_folder) if f.lower().endswith(".pdf")
      ]
      for f in archivos:
        ruta_pdf = os.path.join(ruta_folder, f)
        if f.lower() not in rutas_procesadas:
          rutas_procesadas.add(f.lower())
          quitar_solo_lectura(ruta_pdf)
          lista_pdfs_disponibles.append((f, ruta_pdf))

    if not lista_pdfs_disponibles:
      self.log("\n❌ Carpeta localizada pero no contiene ningún PDF.")
      return

    ventana_modal = VentanaSeleccionPDFs(
        self, no_parte_seleccionado, lista_pdfs_disponibles
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
    self.log("✔ Presione el botón para enviar a la cola de impresión.")
    self.btn_imprimir.configure(state="normal")

  def imprimir_pdfs(self):
    if (
        not hasattr(self, "pdfs_seleccionados_lista")
        or not self.pdfs_seleccionados_lista
    ):
      self.log("❌ No hay ningún paquete preparado.")
      return

    impresora_sel = self.combo_impresora.get()
    modo_color = (
        self.combo_color.get() if self.permitir_color else "Blanco y Negro"
    )
    es_color = modo_color == "Color"

    try:
      self.log(f"\n🖨 Enviando a: [{impresora_sel}]")
      self.log(
          f"   • Documentos independientes: {len(self.pdfs_seleccionados_lista)}"
      )
      self.log(
          "   • Modo: Dúplex inteligente por documento |"
          f' {"Color" if es_color else "Blanco y Negro"}'
      )

      mandar_a_imprimir_configurado(
          self.pdfs_seleccionados_lista, impresora_sel, a_color=es_color
      )
      self.log(
          "✔ Todos los documentos fueron enviados exitosamente a la cola de"
          " Windows."
      )
    except Exception as e:
      self.log(f"❌ Error al enviar a la impresora: {e}")


if __name__ == "__main__":
  app = BuscadorChecklistsApp()
  app.mainloop()
