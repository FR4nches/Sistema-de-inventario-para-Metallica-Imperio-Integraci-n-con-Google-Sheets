import tkinter as tk
from tkinter import messagebox, ttk, simpledialog
from datetime import datetime
from PIL import Image, ImageTk
from tkcalendar import DateEntry

# ==== Google Sheets ====
import gspread
from google.oauth2.service_account import Credentials

# =========================
#   ARCHIVO DE IMAGEN DE FONDO
# =========================
FONDO_IMG = "fondo.webp"   # <<---- IMPORTANTE: TU ARCHIVO

# =========================
#   COLOR PRINCIPAL
# =========================
AZUL_REY = "#1A237E"

# =========================
#   NOMBRE DEL ARCHIVO DE GOOGLE SHEETS
# =========================
SHEETS_FILE_NAME = "Inventario_Pruebas"

# =========================
#   NOMBRES DE HOJAS
# =========================
WS_DIM_ITEMS = "DIM_ITEMS"
WS_DIM_PROVEEDOR = "DIM_PROVEEDOR"
WS_FACTURAS_PEND = "FACTURAS_PENDIENTES"
WS_INVENTARIO = "INVENTARIO"
WS_FACTURAS_RECH = "FACTURAS_RECHAZADAS"

# =========================
#   HEADERS
# =========================
HEADERS_ITEMS = ["sku", "desc_sae", "familia", "proveedor", "stock", "precio"]
HEADERS_PROV = ["id", "nombre", "contacto", "telefono", "email"]
HEADERS_FACT = ["indice", "proveedor", "factura", "fecha", "hora", "cantidad", "creado", "rev"]
HEADERS_RECH = HEADERS_FACT + ["observacion"]

# =========================
#   Datos en memoria
# =========================
usuarios = {
    "superusuario": {"password": "1234", "rol": "superusuario"},
    "supervisor1": {"password": "9999", "rol": "supervisor"},
    "usuario1": {"password": "abcd", "rol": "usuario"},
}

inventario = []
facturas_pendientes = []
facturas_rechazadas = []
proveedores = []
items = []

contador_serie_factura = 1
# =========================
#   Google Sheets helpers
# =========================
def gs_authorize():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
    client = gspread.authorize(creds)
    return client.open(SHEETS_FILE_NAME)

def ws_get_or_create(sh, title, headers):
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=200, cols=max(len(headers), 12))
        ws.update(range_name="1:1", values=[headers])

    first_row = ws.row_values(1)
    if [h.lower() for h in first_row] != [h.lower() for h in headers]:
        ws.update(range_name="1:1", values=[headers])

    return ws

def dicts_from_ws(ws, headers):
    try:
        recs = ws.get_all_records(expected_headers=headers)
    except:
        recs = ws.get_all_records()

    norm = []
    for r in recs:
        d = {}
        for h in headers:
            d[h] = r.get(h, "")
        norm.append(d)

    if not norm:
        all_vals = ws.get_all_values()
        if len(all_vals) > 1:
            hdrs = [h.strip().lower() for h in all_vals[0]]
            for row in all_vals[1:]:
                if not any(row):
                    continue
                tmp = {}
                for j, h in enumerate(hdrs):
                    if j < len(row):
                        tmp[h] = row[j]
                norm.append({h: tmp.get(h, "") for h in headers})

    return norm

def append_dict(ws, headers, data):
    safe = {h: data.get(h, "") for h in headers}
    row = [str(safe.get(h, "")) for h in headers]
    ws.append_row(row, value_input_option="USER_ENTERED")

def find_row_by_key(ws, headers, key_col, key_value):
    all_vals = ws.get_all_values()
    if not all_vals:
        return None
    header_row = all_vals[0]
    header_map = {h.lower(): i+1 for i, h in enumerate(header_row)}
    col_idx = header_map.get(key_col.lower())
    if not col_idx:
        return None
    for i in range(1, len(all_vals)):
        if col_idx-1 < len(all_vals[i]):
            if all_vals[i][col_idx-1] == str(key_value):
                return i+1
    return None

def update_row(ws, headers, row_idx, data):
    values = [[str(data.get(h, "")) for h in headers]]
    start_col = 1
    end_col = len(headers)
    range_notation = (
        gspread.utils.rowcol_to_a1(row_idx, start_col)
        + ":"
        + gspread.utils.rowcol_to_a1(row_idx, end_col)
    )
    ws.update(range_notation, values, value_input_option="USER_ENTERED")

def delete_row(ws, row_idx):
    ws.delete_rows(row_idx)

def find_row_by_criteria(ws, criteria: dict):
    all_vals = ws.get_all_values()
    if not all_vals:
        return None
    headers = [h.strip().lower() for h in all_vals[0]]
    idxs = {k.lower(): headers.index(k.lower()) for k in criteria.keys() if k.lower() in headers}
    if len(idxs) != len(criteria):
        return None
    for i in range(1, len(all_vals)):
        row = all_vals[i]
        ok = True
        for k, v in criteria.items():
            j = idxs[k.lower()]
            val = row[j] if j < len(row) else ""
            if str(val).strip() != str(v).strip():
                ok = False
                break
        if ok:
            return i + 1
    return None

# =========================
#   Inicializar conexión y hojas
# =========================
sheet = gs_authorize()
ws_items = ws_get_or_create(sheet, WS_DIM_ITEMS, HEADERS_ITEMS)
ws_prov  = ws_get_or_create(sheet, WS_DIM_PROVEEDOR, HEADERS_PROV)
ws_pend  = ws_get_or_create(sheet, WS_FACTURAS_PEND, HEADERS_FACT)
ws_invt  = ws_get_or_create(sheet, WS_INVENTARIO, HEADERS_FACT)
ws_rech  = ws_get_or_create(sheet, WS_FACTURAS_RECH, HEADERS_RECH)

# =========================
#   Cargar datos a memoria
# =========================
items = dicts_from_ws(ws_items, HEADERS_ITEMS)
proveedores = dicts_from_ws(ws_prov, HEADERS_PROV)
facturas_pendientes = dicts_from_ws(ws_pend, HEADERS_FACT)
inventario = dicts_from_ws(ws_invt, HEADERS_FACT)
facturas_rechazadas = dicts_from_ws(ws_rech, HEADERS_RECH)
# =========================
#   Utilidades
# =========================
def sku_existe_en_items(sku):
    sku = str(sku).strip()
    for i in items:
        if str(i.get("sku", "")).strip() == sku:
            return True
    return False

def descripcion_por_sku(sku):
    sku = str(sku).strip()
    for i in items:
        if str(i.get("sku", "")).strip() == sku:
            return i.get("desc_sae", "")
    return ""

def serie_proveedor_existe(serie, incluir_listas_extra=None):
    if not serie:
        return False

    listas = [inventario, facturas_pendientes, facturas_rechazadas]
    if incluir_listas_extra:
        listas.extend(incluir_listas_extra)

    for lista in listas:
        for f in lista:
            if f.get("serie_proveedor", "") == serie:
                return True
    return False

def get_proveedor_nombres():
    return [p["nombre"] for p in proveedores]

def pedir_valor(titulo, prompt, valor_inicial=""):
    return simpledialog.askstring(titulo, prompt, initialvalue=valor_inicial)

def pedir_opcion_lista(titulo, prompt, opciones):
    if not opciones:
        return None

    top = tk.Toplevel()
    top.title(titulo)
    top.geometry("350x150")
    top.configure(bg="white")

    tk.Label(top, text=prompt, bg="white", font=("Arial", 11)).pack(padx=10, pady=10)

    sel = tk.StringVar(value=opciones[0])
    combo = ttk.Combobox(top, values=opciones, textvariable=sel, state="readonly", width=30)
    combo.pack(pady=5)

    res = {"value": None}

    def confirmar():
        res["value"] = sel.get()
        top.destroy()

    ttk.Button(top, text="Aceptar", command=confirmar).pack(pady=10)

    top.grab_set()
    top.wait_window()
    return res["value"]

def confirmar_si_no(titulo, texto):
    return messagebox.askyesno(titulo, texto)

def validar_fecha_ddmmyyyy(txt):
    try:
        datetime.strptime(txt, "%d/%m/%Y")
        return True
    except:
        return False


# =========================
#   Funciones de apoyo para encabezados azules en Treeview
# =========================
def aplicar_estilo_treeview():
    """
    Aplica estilo azul rey a encabezados del Treeview.
    Esto se llama cada vez que una ventana nueva se crea.
    """
    style = ttk.Style()
    style.theme_use("clam")

    # Encabezados azul rey
    style.configure(
        "Treeview.Heading",
        background=AZUL_REY,
        foreground="white",
        font=("Arial", 11, "bold")
    )

    # Bordes suaves
    style.configure(
        "Treeview",
        highlightthickness=0,
        bd=0,
        font=("Arial", 10)
    )
# =========================
#   LOGIN MODERNIZADO
# =========================

def crear_panel_redondeado(parent, bg_color="white", radius=25):
    """
    Crea un panel redondeado usando Canvas.
    Devuelve un Frame dentro del panel.
    """
    canvas = tk.Canvas(parent, bg=parent["bg"], highlightthickness=0)
    canvas.pack(expand=True)

    w = 420
    h = 330
    x1 = 10
    y1 = 10
    x2 = x1 + w
    y2 = y1 + h

    # Fondo redondeado
    canvas.create_round_rect = lambda *args, **kwargs: None
    try:
        canvas.create_round_rect(x1, y1, x2, y2, radius=radius, fill=bg_color, outline=bg_color)
    except:
        # Si tkinter no soporta create_round_rect, dibujar manual
        canvas.create_rectangle(x1, y1, x2, y2, fill=bg_color, outline=bg_color)

    frame = tk.Frame(canvas, bg=bg_color)
    frame.place(x=x1+15, y=y1+15, width=w-30, height=h-30)

    return frame


def verificar_login():
    user = entry_usuario.get()
    pwd = entry_contrasena.get()

    if user in usuarios and usuarios[user]["password"] == pwd:
        rol = usuarios[user]["rol"]
        messagebox.showinfo("Login", f"¡Bienvenido {user}! Rol: {rol}")
        ventana_login.withdraw()
        mostrar_menu(user, rol)
    else:
        messagebox.showerror("Error", "Usuario o contraseña incorrecta")


# =========================
#   VENTANA DE LOGIN (FULL RESPONSIVO)
# =========================

def actualizar_imagen_login(event=None):
    """Reescala automáticamente la imagen del login según el tamaño del panel izquierdo."""
    try:
        w = frame_izq.winfo_width()
        h = frame_izq.winfo_height()

        img = Image.open("login.png")
        # La imagen ocupa el 70% del ancho y alto del panel izquierdo
        img = img.resize((int(w * 0.7), int(h * 0.7)))
        img_tk = ImageTk.PhotoImage(img)

        lbl_imagen.config(image=img_tk)
        lbl_imagen.image = img_tk
    except:
        pass


ventana_login = tk.Tk()
ventana_login.title("Login - Sistema de Inventario")
ventana_login.state("zoomed")

try:
    ventana_login.attributes('-zoomed', True)
except:
    pass

# Panel izquierdo azul rey (imagen responsiva)
frame_izq = tk.Frame(ventana_login, bg=AZUL_REY)
frame_izq.pack(side="left", fill="both", expand=True)

lbl_imagen = tk.Label(frame_izq, bg=AZUL_REY)
lbl_imagen.place(relx=0.5, rely=0.5, anchor="center")

# Cada vez que cambia el tamaño, reescala la imagen
frame_izq.bind("<Configure>", actualizar_imagen_login)

# Panel derecho gris claro
frame_der = tk.Frame(ventana_login, bg="#EEEEEE")
frame_der.pack(side="right", fill="both", expand=True)

# Panel blanco del login, proporcional al tamaño de la ventana
panel_login = tk.Frame(frame_der, bg="white", bd=2, relief="ridge")
def ajustar_panel_login(event=None):
    """Ajusta el tamaño del panel de login según la pantalla, con límites min y max."""
    w = frame_der.winfo_width()
    h = frame_der.winfo_height()

    # Porcentajes base
    rel_w = 0.35
    rel_h = 0.45

    # Cálculo en píxeles basado en la ventana
    calc_w = int(w * rel_w)
    calc_h = int(h * rel_h)

    # Límites para no romper la UI
    min_w, min_h = 320, 280
    max_w, max_h = 450, 420

    # Aplicar límites
    final_w = max(min(calc_w, max_w), min_w)
    final_h = max(min(calc_h, max_h), min_h)

    # Colocar el panel centrado con tamaño en pixeles
    panel_login.place(x=w/2 - final_w/2,
                      y=h/2 - final_h/2,
                      width=final_w,
                      height=final_h)

# Activar ajuste dinámico
frame_der.bind("<Configure>", ajustar_panel_login)

# ============ CAMPOS DEL LOGIN ==========
tk.Label(panel_login, text="Inicio de Sesión",
         font=("Arial", 20, "bold"),
         bg="white", fg=AZUL_REY).pack(pady=15)

tk.Label(panel_login, text="Usuario:", bg="white", fg="black",
         font=("Arial", 11)).pack()
entry_usuario = tk.Entry(panel_login, width=30, font=("Arial", 11))
entry_usuario.pack(pady=5)

tk.Label(panel_login, text="Contraseña:", bg="white", fg="black",
         font=("Arial", 11)).pack()
entry_contrasena = tk.Entry(panel_login, width=30, show="*", font=("Arial", 11))
entry_contrasena.pack(pady=5)

tk.Button(panel_login, text="Ingresar", command=verificar_login,
          bg="#4CAF50", fg="white", font=("Arial", 13, "bold"),
          width=15).pack(pady=20)

# =========================
#   FONDO CON IMAGEN Y MENÚ PRINCIPAL
# =========================

def crear_fondo_imagen(ventana):
    """
    Fondo RESPONSIVO estilo COVER:
    - Siempre cubre toda la pantalla
    - No se deforma
    - No deja espacios en blanco
    - Mantiene relación de aspecto
    - Reescala automáticamente al cambiar el tamaño
    """

    try:
        ventana.update_idletasks()

        img_original = Image.open(FONDO_IMG)

        # Label que va a contener el fondo
        fondo_label = tk.Label(ventana)
        fondo_label.place(x=0, y=0, relwidth=1, relheight=1)

        def actualizar_fondo(event=None):
            win_w = ventana.winfo_width()
            win_h = ventana.winfo_height()

            # Aspect ratio de ventana y de la imagen
            win_ratio = win_w / win_h
            img_ratio = img_original.width / img_original.height

            if win_ratio > img_ratio:
                # La ventana es más ancha → expandimos por ancho
                new_w = win_w
                new_h = int(win_w / img_ratio)
            else:
                # La ventana es más alta → expandimos por alto
                new_h = win_h
                new_w = int(win_h * img_ratio)

            # Reescalar con calidad alta (LANCZOS)
            img_resized = img_original.resize((new_w, new_h), Image.LANCZOS)
            img_tk = ImageTk.PhotoImage(img_resized)

            fondo_label.config(image=img_tk)
            fondo_label.image = img_tk

        # Actualizar inmediatamente
        actualizar_fondo()

        # Actualizar cada vez que se redimensiona la ventana
        ventana.bind("<Configure>", actualizar_fondo)

    except Exception as e:
        print("No se pudo cargar el fondo responsivo:", e)
        ventana.configure(bg="#E0E0E0")



def cerrar_sesion(ventana_menu):
    ventana_menu.destroy()
    ventana_login.deiconify()
    try:
        ventana_login.state("zoomed")
        ventana_login.attributes("-zoomed", True)
    except:
        pass
    entry_usuario.delete(0, tk.END)
    entry_contrasena.delete(0, tk.END)


# =========================
#   Menú principal
# =========================
def mostrar_menu(usuario, rol):
    ventana_menu = tk.Toplevel()
    ventana_menu.title("Menú Principal - Sistema de Inventario")
    ventana_menu.state("zoomed")
    try:
        ventana_menu.attributes("-zoomed", True)
    except:
        pass

    # Fondo con imagen
    crear_fondo_imagen(ventana_menu)

    # Contenedor central (tarjeta blanca)
    contenedor = tk.Frame(ventana_menu, bg="white", bd=0, highlightthickness=0)
    contenedor.place(relx=0.5, rely=0.5, anchor="center")

    # "Card" interna con padding
    marco = tk.Frame(contenedor, bg="white", bd=2, relief="ridge")
    marco.pack(padx=20, pady=20)

    # Título usuario/rol
    tk.Label(
        marco,
        text=f"Usuario: {usuario} | Rol: {rol}",
        font=("Arial", 14, "bold"),
        bg="white",
        fg=AZUL_REY
    ).grid(row=0, column=0, columnspan=2, pady=(10, 20), padx=20)

    # Frame para los botones
    cont = tk.Frame(marco, bg="white")
    cont.grid(row=1, column=0, columnspan=2, pady=10, padx=20)

    boton_estilo = {
        "bg": AZUL_REY,
        "fg": "white",
        "font": ("Arial", 11, "bold"),
        "width": 25,
        "height": 1,
        "bd": 0,
        "activebackground": "#283593",
        "activeforeground": "white",
        "cursor": "hand2",
    }

    fila = 0

    # 1) Ver inventario (todos)
    tk.Button(
        cont,
        text="Ver inventario",
        command=ver_inventario,
        **boton_estilo
    ).grid(row=fila, column=0, padx=10, pady=6, sticky="ew")
    fila += 1

    # 2) Revisar facturas (solo supervisor)
    if rol == "supervisor":
        tk.Button(
            cont,
            text="Revisar facturas",
            command=validar_facturas,
            **boton_estilo
        ).grid(row=fila, column=0, padx=10, pady=6, sticky="ew")
        fila += 1

    # 3) Usuarios (solo superusuario)
    if rol == "superusuario":
        tk.Button(
            cont,
            text="Usuarios",
            command=lambda: menu_usuarios(),
            **boton_estilo
        ).grid(row=fila, column=0, padx=10, pady=6, sticky="ew")
        fila += 1

    # 4) Proveedor (solo supervisor)
    if rol == "supervisor":
        tk.Button(
            cont,
            text="Proveedor",
            command=menu_proveedor,
            **boton_estilo
        ).grid(row=fila, column=0, padx=10, pady=6, sticky="ew")
        fila += 1

    # 5) Items (solo supervisor)
    if rol == "supervisor":
        tk.Button(
            cont,
            text="Items",
            command=menu_items,
            **boton_estilo
        ).grid(row=fila, column=0, padx=10, pady=6, sticky="ew")
        fila += 1

        tk.Button(
            cont,
            text="Gestión de Inventario",
            command=lambda: ventana_gestion_inventario(usuario),
            **boton_estilo
        ).grid(row=fila, column=0, padx=10, pady=6, sticky="ew")
        fila += 1

    # Usuario normal: Ingresar facturas + facturas rechazadas
    if rol == "usuario":
        tk.Button(
            cont,
            text="Ingresar facturas",
            command=lambda: ingresar_factura(usuario, rol),
            **boton_estilo
        ).grid(row=fila, column=0, padx=10, pady=6, sticky="ew")
        fila += 1

        tk.Button(
            cont,
            text="Ver facturas rechazadas",
            command=lambda: ver_facturas_rechazadas(usuario),
            **boton_estilo
        ).grid(row=fila, column=0, padx=10, pady=6, sticky="ew")
        fila += 1

    # Botón de cerrar sesión (todos)
    tk.Button(
        marco,
        text="Cerrar sesión",
        command=lambda: cerrar_sesion(ventana_menu),
        bg="#E53935",
        fg="white",
        font=("Arial", 11, "bold"),
        width=20,
        bd=0,
        activebackground="#C62828",
        activeforeground="white",
        cursor="hand2"
    ).grid(row=2, column=0, columnspan=2, pady=(10, 15))
# =========================
#   Facturas (actualizado + estilizado)
# =========================

def ingresar_factura(usuario, rol):
    lista_prov = get_proveedor_nombres()
    if not lista_prov:
        messagebox.showerror("Proveedores", "No hay proveedores cargados. Solicite al supervisor que agregue proveedores antes de ingresar facturas.")
        return

    ventana_factura = tk.Toplevel()
    ventana_factura.title("Ingresar Factura")
    ventana_factura.geometry("1100x750")
    ventana_factura.state("zoomed")

    try:
        ventana_factura.attributes("-zoomed", True)
    except:
        pass

    # Fondo
    crear_fondo_imagen(ventana_factura)

    # Panel central estilo tarjeta
    card = tk.Frame(ventana_factura, bg="white", bd=2, relief="ridge")
    card.place(relx=0.5, rely=0.5, anchor="center", width=1000, height=680)

    tk.Label(card, text="Ingreso de Factura", font=("Arial", 18, "bold"),
             fg=AZUL_REY, bg="white").pack(pady=10)

    # ============ Cabecera ============
    marco = tk.Frame(card, bg="white")
    marco.pack(pady=5)

    fecha_sistema = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    tk.Label(marco, text=f"Fecha del sistema: {fecha_sistema}",
             font=("Arial", 10, "italic"), bg="white").grid(row=0, column=0, columnspan=2, pady=(0,10))

    tk.Label(marco, text="Proveedor:", bg="white").grid(row=1, column=0, sticky="e", padx=5, pady=3)
    var_prov = tk.StringVar(value=lista_prov[0])
    entradas_cabecera = {}
    entradas_cabecera["proveedor"] = ttk.Combobox(marco, values=lista_prov,
                                                  textvariable=var_prov, state="readonly", width=40)
    entradas_cabecera["proveedor"].grid(row=1, column=1, padx=5, pady=3)

    tk.Label(marco, text="Documento:", bg="white").grid(row=2, column=0, sticky="e", padx=5, pady=3)
    entradas_cabecera["documento"] = tk.Entry(marco, width=43)
    entradas_cabecera["documento"].grid(row=2, column=1, padx=5, pady=3)

    hoy = datetime.now()

    tk.Label(marco, text="Fecha del documento:", bg="white").grid(row=3, column=0, sticky="e", padx=5, pady=3)
    entradas_cabecera["fecha_doc"] = DateEntry(
        marco, width=18, background=AZUL_REY, foreground='white',
        borderwidth=2, date_pattern='dd/mm/yyyy',
        year=hoy.year, month=hoy.month, day=hoy.day
    )
    entradas_cabecera["fecha_doc"].grid(row=3, column=1, pady=3, sticky="w")

    tk.Label(marco, text="Fecha de recepción:", bg="white").grid(row=4, column=0, sticky="e", padx=5, pady=3)
    entradas_cabecera["fecha_recepcion"] = DateEntry(
        marco, width=18, background=AZUL_REY, foreground='white',
        borderwidth=2, date_pattern='dd/mm/yyyy',
        year=hoy.year, month=hoy.month, day=hoy.day
    )
    entradas_cabecera["fecha_recepcion"].grid(row=4, column=1, pady=3, sticky="w")

    # =============== Línea de factura ===============
    ttk.Separator(card, orient='horizontal').pack(fill="x", padx=20, pady=15)
    tk.Label(card, text="Líneas de Factura", font=("Arial", 14, "bold"),
             bg="white", fg=AZUL_REY).pack()

    frame_lineas = tk.Frame(card, bg="white")
    frame_lineas.pack(pady=10)

    columnas = ["Serie Interna", "SKU", "Descripción", "Cantidad", "Serie Proveedor"]
    for j, col in enumerate(columnas):
        tk.Label(frame_lineas, text=col, font=("Arial", 10, "bold"),
                 bg="white", fg=AZUL_REY).grid(row=0, column=j, padx=5, pady=3)

    lineas = []

    def crear_linea():
        fila = len(lineas) + 1
        entrada = {
            "serie_interna": tk.Entry(frame_lineas, width=15),
            "sku": tk.Entry(frame_lineas, width=15),
            "descripcion": tk.Entry(frame_lineas, width=30, state="readonly"),
            "cantidad": tk.Entry(frame_lineas, width=10),
            "serie_prov": tk.Entry(frame_lineas, width=15)
        }

        for j, key in enumerate(["serie_interna", "sku", "descripcion", "cantidad", "serie_prov"]):
            entrada[key].grid(row=fila, column=j, padx=5, pady=2)

        def on_sku_change(event=None, e=entrada):
            sku = e["sku"].get().strip()
            desc = descripcion_por_sku(sku)
            e["descripcion"].config(state="normal")
            e["descripcion"].delete(0, tk.END)
            if desc:
                e["descripcion"].insert(0, desc)
            e["descripcion"].config(state="readonly")

        entrada["sku"].bind("<FocusOut>", on_sku_change)
        entrada["sku"].bind("<KeyRelease>", on_sku_change)

        lineas.append(entrada)

    for _ in range(10):
        crear_linea()

    def agregar_linea_extra():
        crear_linea()

    def eliminar_linea():
        if len(lineas) > 1:
            ultima = lineas.pop()
            for w in ultima.values():
                w.destroy()
        else:
            messagebox.showwarning("Aviso", "Debe haber al menos una línea.")

    frame_botones = tk.Frame(card, bg="white")
    frame_botones.pack(pady=8)

    tk.Button(frame_botones, text="Agregar línea ➕", command=agregar_linea_extra,
              bg=AZUL_REY, fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=5)

    tk.Button(frame_botones, text="Eliminar línea ➖", command=eliminar_linea,
              bg="#E53935", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=5)

    # ============ Validación y envío ============
    def validar_cabecera():
        if not entradas_cabecera["documento"].get().strip():
            messagebox.showerror("Error", "El campo 'Documento' es obligatorio.")
            return False
        fd = entradas_cabecera["fecha_doc"].get().strip()
        fr = entradas_cabecera["fecha_recepcion"].get().strip()
        if not validar_fecha_ddmmyyyy(fd) or not validar_fecha_ddmmyyyy(fr):
            messagebox.showerror("Error", "Fechas inválidas. Use formato dd/mm/yyyy.")
            return False
        return True

    def enviar_factura():
        if not validar_cabecera():
            return

        proveedor = entradas_cabecera["proveedor"].get().strip().upper()
        factura = entradas_cabecera["documento"].get().strip().upper()
        codigo_rev = proveedor + factura

        rev_existentes = [str(r.get("rev", "")).strip().upper() for r in dicts_from_ws(ws_invt, HEADERS_FACT)]

        if codigo_rev in rev_existentes:
            messagebox.showwarning("Factura duplicada", "Esa factura ya existe en el Inventario.")
            return

        registros = []
        series_internas_usadas = []

        for e in lineas:
            serie_interna = e["serie_interna"].get().strip()
            sku = e["sku"].get().strip()
            desc = e["descripcion"].get().strip()
            cant = e["cantidad"].get().strip()
            serie_prov = e["serie_prov"].get().strip()

            if not sku and not serie_interna:
                continue

            if not serie_interna:
                return messagebox.showerror("Error", "Cada línea debe tener una serie interna.")

            if serie_interna in series_internas_usadas:
                return messagebox.showerror("Error", f"La serie '{serie_interna}' está repetida.")
            series_internas_usadas.append(serie_interna)

            if not sku_existe_en_items(sku):
                return messagebox.showerror("Error", f"El SKU '{sku}' no existe.")
            if not cant:
                return messagebox.showerror("Error", f"Debe ingresar cantidad en la línea de serie '{serie_interna}'.")
            try:
                float(cant)
            except:
                return messagebox.showerror("Error", f"La cantidad '{cant}' no es numérica.")

            registro = {
                "indice": "",
                "proveedor": entradas_cabecera["proveedor"].get().strip(),
                "factura": entradas_cabecera["documento"].get().strip(),
                "fecha": datetime.now().strftime("%d/%m/%Y"),
                "hora": datetime.now().strftime("%H:%M:%S"),
                "cantidad": cant,
                "creado": usuario,
                "rev": ""
            }
            registros.append(registro)

        if not registros:
            return messagebox.showwarning("Aviso", "No hay líneas válidas.")

        if not confirmar_si_no("Confirmar", f"¿Enviar {len(registros)} líneas?"):
            return

        for r in registros:
            append_dict(ws_pend, HEADERS_FACT, r)
            facturas_pendientes.append({k: r.get(k, "") for k in HEADERS_FACT})

        ventana_factura.destroy()
        messagebox.showinfo("Éxito", f"{len(registros)} líneas enviadas para revisión.")

    tk.Button(card, text="Enviar factura", command=enviar_factura,
              bg="#4CAF50", fg="white", font=("Arial", 13, "bold"),
              width=20).pack(pady=20)
# =========================
#   Ver facturas rechazadas (estilizado)
# =========================

def ver_facturas_rechazadas(usuario):
    v = tk.Toplevel()
    v.title("Facturas Rechazadas")
    v.geometry("1100x700")
    v.state("zoomed")

    try:
        v.attributes("-zoomed", True)
    except:
        pass

    crear_fondo_imagen(v)
    aplicar_estilo_treeview()

    card = tk.Frame(v, bg="white", bd=2, relief="ridge")
    card.place(relx=0.5, rely=0.5, anchor="center", width=1000, height=600)

    tk.Label(card, text="Facturas Rechazadas", font=("Arial", 18, "bold"),
             fg=AZUL_REY, bg="white").pack(pady=15)

    rechazadas_usuario = [
        f for f in dicts_from_ws(ws_rech, HEADERS_RECH)
        if f.get("creado") == usuario
    ]

    if not rechazadas_usuario:
        messagebox.showinfo("Facturas Rechazadas", "No tiene facturas rechazadas.")
        v.destroy()
        return

    cols = ("indice","proveedor","factura","fecha","hora","cantidad","creado","rev","observacion")

    tabla = ttk.Treeview(card, columns=cols, show="headings", height=20)
    tabla.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

    encabezados = {
        "indice":"Índice",
        "proveedor":"Proveedor",
        "factura":"Factura",
        "fecha":"Fecha",
        "hora":"Hora",
        "cantidad":"Cantidad",
        "creado":"Creado por",
        "rev":"Revisión",
        "observacion":"Observación"
    }

    for c in cols:
        tabla.heading(c, text=encabezados[c])
        tabla.column(c, width=120, anchor="center")

    for f in rechazadas_usuario:
        tabla.insert("", tk.END, values=tuple(f.get(c, "") for c in cols))

    tk.Button(
        card, text="Cerrar", command=v.destroy,
        bg="#E53935", fg="white",
        font=("Arial", 11, "bold"), width=15
    ).pack(pady=10)



# =========================
#   Validación de facturas (Supervisor) — estilizada
# =========================

def validar_facturas():
    global facturas_pendientes
    facturas_pendientes = dicts_from_ws(ws_pend, HEADERS_FACT)

    v = tk.Toplevel()
    v.title("Revisar / Validar Facturas")
    v.geometry("1200x750")
    v.state("zoomed")
    try:
        v.attributes('-zoomed', True)
    except:
        pass

    crear_fondo_imagen(v)
    aplicar_estilo_treeview()

    card = tk.Frame(v, bg="white", bd=2, relief="ridge")
    card.place(relx=0.5, rely=0.5, anchor="center")
    card.configure(width=1000, height=650)



    tk.Label(card, text="Revisión de Facturas Pendientes",
             font=("Arial", 18, "bold"), fg=AZUL_REY, bg="white").pack(pady=10)

    cols = ("indice", "proveedor", "factura", "fecha", "hora", "cantidad", "creado", "rev")

    tabla = ttk.Treeview(card, columns=cols, show="headings", height=20)
    tabla.pack(fill="both", expand=True, padx=15, pady=10)

    encabezados = {
        "indice": "Índice",
        "proveedor": "Proveedor",
        "factura": "Factura",
        "fecha": "Fecha",
        "hora": "Hora",
        "cantidad": "Cantidad",
        "creado": "Creado por",
        "rev": "Revisión"
    }

    for c in cols:
        tabla.heading(c, text=encabezados[c])
        tabla.column(c, width=130, anchor="center")

    tabla.delete(*tabla.get_children())
    all_vals = ws_pend.get_all_values()
    if len(all_vals) > 1:
        headers = [h.strip().lower() for h in all_vals[0]]
        for i in range(1, len(all_vals)):
            row = all_vals[i]
            if not any(row):
                continue
            fila_dict = {}
            for j, h in enumerate(headers):
                if j < len(row):
                    fila_dict[h] = row[j]
            tabla.insert("", tk.END, values=tuple(fila_dict.get(c, "") for c in cols))


    # =========================
    #   BOTONES DE APROBAR / RECHAZAR
    # =========================
    btn_frame = tk.Frame(card, bg="white")
    btn_frame.pack(pady=15)

    def aprobar():
        sel = tabla.selection()
        if not sel:
            return messagebox.showwarning("Aviso", "Seleccione una factura.")
        moved = 0

        for s in sel:
            vals = tabla.item(s)["values"]
            f = dict(zip(cols, vals))

            append_dict(ws_invt, HEADERS_FACT, f)

            crit = {
                "proveedor": f.get("proveedor", ""),
                "factura": f.get("factura", ""),
                "creado": f.get("creado", ""),
                "hora": f.get("hora", "")
            }
            ridx = find_row_by_criteria(ws_pend, crit)
            if ridx:
                delete_row(ws_pend, ridx)

            tabla.delete(s)
            moved += 1

        messagebox.showinfo("Aprobación", f"Factura(s) aprobadas: {moved}")


    def rechazar():
        sel = tabla.selection()
        if not sel:
            return messagebox.showwarning("Aviso", "Seleccione una factura.")

        moved = 0

        for s in sel:
            vals = tabla.item(s)["values"]
            f = dict(zip(cols, vals))

            obs = simpledialog.askstring("Observación", "Motivo del rechazo:")
            comp = dict(f)
            comp["observacion"] = obs if obs else "Rechazada sin observación"

            append_dict(ws_rech, HEADERS_RECH, comp)

            crit = {
                "proveedor": f.get("proveedor", ""),
                "factura": f.get("factura", ""),
                "creado": f.get("creado", ""),
                "hora": f.get("hora", "")
            }
            ridx = find_row_by_criteria(ws_pend, crit)
            if ridx:
                delete_row(ws_pend, ridx)

            tabla.delete(s)
            moved += 1

        messagebox.showinfo("Rechazo", f"Factura(s) rechazadas: {moved}")


    tk.Button(
        btn_frame, text="Aprobar", command=aprobar,
        bg="#43A047", fg="white", font=("Arial", 11, "bold"),
        width=15
    ).pack(side="left", padx=10)

    tk.Button(
        btn_frame, text="Rechazar", command=rechazar,
        bg="#E53935", fg="white", font=("Arial", 11, "bold"),
        width=15
    ).pack(side="left", padx=10)
# =========================
#   Vista de Inventario (encabezados azules)
# =========================
def ver_inventario():
    global inventario
    inventario = dicts_from_ws(ws_invt, HEADERS_FACT)

    v = tk.Toplevel()
    v.title("Inventario")
    v.state("zoomed")

    try:
        v.attributes('-zoomed', True)
    except:
        pass

    crear_fondo_imagen(v)
    aplicar_estilo_treeview()

    card = tk.Frame(v, bg="white", bd=2, relief="ridge")
    card.place(relx=0.5, rely=0.5, anchor="center", width=1100, height=650)

    tk.Label(card, text="Inventario", fg=AZUL_REY,
             bg="white", font=("Arial", 18, "bold")).pack(pady=15)

    cols = ("indice","proveedor","factura","fecha","hora","cantidad","creado","rev")

    tabla = ttk.Treeview(card, columns=cols, show="headings", height=22)
    tabla.pack(fill=tk.BOTH, padx=20, pady=10, expand=True)

    encabezados = {
        "indice": "Índice",
        "proveedor": "Proveedor",
        "factura": "Factura",
        "fecha": "Fecha",
        "hora": "Hora",
        "cantidad": "Cantidad",
        "creado": "Creado por",
        "rev": "Revisión"
    }

    for c in cols:
        tabla.heading(c, text=encabezados[c])
        tabla.column(c, width=120, anchor="center")

    for i in inventario:
        tabla.insert("", tk.END, values=tuple(i.get(c, "") for c in cols))


# =========================
#   Gestión de Inventario (Supervisor)
# =========================

def ventana_gestion_inventario(usuario_actual):
    ventana = tk.Toplevel()
    ventana.title("Gestión de Inventario")
    ventana.geometry("1000x600")
    ventana.state("zoomed")

    try:
        ventana.attributes('-zoomed', True)
    except:
        pass

    crear_fondo_imagen(ventana)
    aplicar_estilo_treeview()

    card = tk.Frame(ventana, bg="white", bd=2, relief="ridge")
    card.place(relx=0.5, rely=0.5, anchor="center")
    card.configure(width=1000, height=580)


    tk.Label(card, text="Gestión de Inventario", fg=AZUL_REY,
             bg="white", font=("Arial", 18, "bold")).pack(pady=8)

    columnas = ("indice","proveedor","factura","fecha","hora","cantidad","creado","rev")
    tree = ttk.Treeview(card, columns=columnas, show="headings", height=20)
    tree.pack(fill="both", expand=True, padx=10, pady=10)

    for col in columnas:
        tree.heading(col, text=col.upper())
        tree.column(col, width=120, anchor="center")

    def refrescar_tabla():
        tree.delete(*tree.get_children())
        try:
            data = ws_invt.get_all_records()
            for fila in data:
                valores = (
                    fila.get("indice", ""),
                    fila.get("proveedor", ""),
                    fila.get("factura", ""),
                    fila.get("fecha", ""),
                    fila.get("hora", ""),
                    fila.get("cantidad", ""),
                    fila.get("creado", ""),
                    fila.get("rev", "")
                )
                tree.insert("", tk.END, values=valores)
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los datos: {e}")

    def modificar_cantidad():
        sel = tree.selection()
        if not sel:
            return messagebox.showwarning("Aviso", "Seleccione una factura.")

        item = tree.item(sel[0])
        valores = item["values"]
        factura = valores[2]
        proveedor = valores[1]
        cantidad_actual = valores[5]

        nueva = simpledialog.askstring(
            "Modificar cantidad",
            f"Factura: {factura}\nProveedor: {proveedor}\nCantidad actual: {cantidad_actual}\n\nNueva cantidad:"
        )

        if not nueva:
            return

        try:
            float(nueva)
        except:
            return messagebox.showerror("Error", "Debe ingresar un número válido.")

        crit = {"proveedor": proveedor, "factura": factura}
        row_idx = find_row_by_criteria(ws_invt, crit)
        if not row_idx:
            return messagebox.showerror("Error", "No se encontró la factura.")

        data = dict(zip(columnas, valores))
        data["cantidad"] = nueva

        update_row(ws_invt, HEADERS_FACT, row_idx, data)
        registrar_historial_rebaja(usuario_actual, "MODIFICAR", proveedor, factura, cantidad_actual, nueva, "Ajuste supervisor")

        messagebox.showinfo("Éxito", "Cantidad modificada.")
        refrescar_tabla()

    def eliminar_factura():
        sel = tree.selection()
        if not sel:
            return messagebox.showwarning("Aviso", "Seleccione una factura.")

        item = tree.item(sel[0])
        valores = item["values"]
        proveedor, factura = valores[1], valores[2]

        if not confirmar_si_no("Confirmar", f"¿Eliminar factura '{factura}'?"):
            return

        crit = {"proveedor": proveedor, "factura": factura}
        row_idx = find_row_by_criteria(ws_invt, crit)
        if not row_idx:
            return messagebox.showerror("Error", "No se encontró la factura.")

        delete_row(ws_invt, row_idx)
        registrar_historial_rebaja(usuario_actual, "ELIMINAR", proveedor, factura, valores[5], "0", "Eliminada")

        messagebox.showinfo("Éxito", "Factura eliminada.")
        refrescar_tabla()

    btn_frame = tk.Frame(card, bg="white")
    btn_frame.pack(pady=15)

    tk.Button(
        btn_frame, text="Modificar cantidad", command=modificar_cantidad,
        bg=AZUL_REY, fg="white", font=("Arial", 11, "bold"), width=20
    ).pack(side="left", padx=10)

    tk.Button(
        btn_frame, text="Eliminar factura", command=eliminar_factura,
        bg="#E67E22", fg="white", font=("Arial", 11, "bold"), width=20
    ).pack(side="left", padx=10)

    tk.Button(
        btn_frame, text="Cerrar", command=ventana.destroy,
        bg="#E53935", fg="white", font=("Arial", 11, "bold"), width=20
    ).pack(side="left", padx=10)

    refrescar_tabla()


# =========================
#   Historial Rebajas
# =========================

def registrar_historial_rebaja(usuario, accion, proveedor, factura, cantidad_anterior, cantidad_nueva, motivo=""):
    global sheet
    fecha = datetime.now().strftime("%d/%m/%Y")
    hora = datetime.now().strftime("%H:%M:%S")
    try:
        ws_hist = sheet.worksheet("HISTORIAL_REBAJAS")
    except gspread.WorksheetNotFound:
        ws_hist = sheet.add_worksheet(title="HISTORIAL_REBAJAS", rows=200, cols=9)
        ws_hist.update(
            "A1:I1",
            [["fecha","hora","usuario","accion","proveedor","factura","cantidad_anterior","cantidad_nueva","motivo"]]
        )

    fila = [fecha, hora, usuario, accion, proveedor, factura, cantidad_anterior, cantidad_nueva, motivo]
    ws_hist.append_row(fila, value_input_option="USER_ENTERED")


# =========================
#   CRUD de Usuarios
# =========================

def menu_usuarios():
    w = tk.Toplevel()
    w.title("Usuarios")
    w.geometry("420x320")
    crear_fondo_imagen(w)

    card = tk.Frame(w, bg="white", bd=2, relief="ridge")
    card.place(relx=0.5, rely=0.5, anchor="center")

    tk.Button(card, text="Agregar usuario", width=30, command=agregar_usuario,
              bg=AZUL_REY, fg="white").pack(pady=6)
    tk.Button(card, text="Modificar usuario", width=30, command=modificar_usuario,
              bg=AZUL_REY, fg="white").pack(pady=6)
    tk.Button(card, text="Eliminar usuario", width=30, command=eliminar_usuario,
              bg=AZUL_REY, fg="white").pack(pady=6)
    tk.Button(card, text="Permisos / Rol", width=30, command=cambiar_rol_usuario,
              bg=AZUL_REY, fg="white").pack(pady=6)


def agregar_usuario():
    nombre = pedir_valor("Agregar usuario", "Usuario:")
    if not nombre:
        return
    if nombre in usuarios:
        return messagebox.showerror("Error", "Ese usuario ya existe.")
    pwd = pedir_valor("Agregar usuario", "Password:")
    if not pwd:
        return
    rol = pedir_opcion_lista("Rol", "Seleccione el rol:", ["superusuario","supervisor","usuario"])
    if not rol:
        return
    usuarios[nombre] = {"password": pwd, "rol": rol}
    messagebox.showinfo("OK", f"Usuario '{nombre}' agregado.")


def modificar_usuario():
    lista = list(usuarios.keys())
    nombre = pedir_opcion_lista("Modificar usuario", "Seleccione usuario:", lista)
    if not nombre:
        return
    pwd = pedir_valor("Modificar usuario", "Nuevo password:", usuarios[nombre]["password"])
    if not pwd:
        return
    usuarios[nombre]["password"] = pwd
    messagebox.showinfo("OK", "Actualizado.")


def eliminar_usuario():
    lista = list(usuarios.keys())
    nombre = pedir_opcion_lista("Eliminar usuario", "Seleccione usuario:", lista)
    if not nombre:
        return
    if not confirmar_si_no("Confirmar", f"¿Eliminar '{nombre}'?"):
        return
    del usuarios[nombre]
    messagebox.showinfo("OK", "Eliminado.")


def cambiar_rol_usuario():
    lista = list(usuarios.keys())
    nombre = pedir_opcion_lista("Cambiar Rol", "Seleccione usuario:", lista)
    if not nombre:
        return
    rol = pedir_opcion_lista("Rol", "Seleccione nuevo rol:", ["superusuario","supervisor","usuario"])
    if not rol:
        return
    usuarios[nombre]["rol"] = rol
    messagebox.showinfo("OK", "Actualizado.")


# =========================
#   CRUD Proveedores
# =========================

def menu_proveedor():
    w = tk.Toplevel()
    w.title("Proveedor")
    w.geometry("420x260")
    crear_fondo_imagen(w)

    card = tk.Frame(w, bg="white", bd=2, relief="ridge")
    card.place(relx=0.5, rely=0.5, anchor="center")

    tk.Button(card, text="Agregar proveedor", width=30, command=agregar_proveedor,
              bg=AZUL_REY, fg="white").pack(pady=6)
    tk.Button(card, text="Modificar proveedor", width=30, command=modificar_proveedor,
              bg=AZUL_REY, fg="white").pack(pady=6)
    tk.Button(card, text="Eliminar proveedor", width=30, command=eliminar_proveedor,
              bg=AZUL_REY, fg="white").pack(pady=6)


def agregar_proveedor():
    global proveedores
    nombre = pedir_valor("Nuevo proveedor", "Nombre:")
    if not nombre:
        return
    if any(p["nombre"] == nombre for p in proveedores):
        return messagebox.showerror("Error", "Proveedor ya existe.")
    contacto = pedir_valor("Nuevo proveedor", "Contacto:") or ""
    telefono = pedir_valor("Nuevo proveedor", "Teléfono:") or ""
    email = pedir_valor("Nuevo proveedor", "Email:") or ""

    nuevo = {
        "id": str(len(proveedores) + 1),
        "nombre": nombre,
        "contacto": contacto,
        "telefono": telefono,
        "email": email
    }

    proveedores.append(nuevo)
    append_dict(ws_prov, HEADERS_PROV, nuevo)

    messagebox.showinfo("OK", "Proveedor agregado.")


def seleccionar_proveedor():
    if not proveedores:
        messagebox.showinfo("Proveedor", "No hay proveedores.")
        return None
    nombres = [p["nombre"] for p in proveedores]
    elegido = pedir_opcion_lista("Seleccionar", "Proveedor:", nombres)
    if not elegido:
        return None
    for p in proveedores:
        if p["nombre"] == elegido:
            return p
    return None


def modificar_proveedor():
    global proveedores
    p = seleccionar_proveedor()
    if not p:
        return

    original = p["nombre"]
    nuevo_nombre = pedir_valor("Modificar", "Nombre:", p["nombre"]) or p["nombre"]
    if nuevo_nombre != p["nombre"] and any(x["nombre"] == nuevo_nombre for x in proveedores):
        return messagebox.showerror("Error", "Ya existe.")

    p["nombre"] = nuevo_nombre
    p["contacto"] = pedir_valor("Modificar", "Contacto:", p["contacto"]) or p["contacto"]
    p["telefono"] = pedir_valor("Modificar", "Teléfono:", p["telefono"]) or p["telefono"]
    p["email"] = pedir_valor("Modificar", "Email:", p["email"]) or p["email"]

    row_idx = find_row_by_key(ws_prov, HEADERS_PROV, "nombre", original)
    if not row_idx:
        row_idx = find_row_by_key(ws_prov, HEADERS_PROV, "id", p["id"])
    if row_idx:
        update_row(ws_prov, HEADERS_PROV, row_idx, p)

    proveedores[:] = dicts_from_ws(ws_prov, HEADERS_PROV)
    messagebox.showinfo("OK", "Proveedor actualizado.")


def eliminar_proveedor():
    global proveedores
    p = seleccionar_proveedor()
    if not p:
        return
    if not confirmar_si_no("Confirmar", f"¿Eliminar '{p['nombre']}'?"):
        return

    row_idx = find_row_by_key(ws_prov, HEADERS_PROV, "nombre", p["nombre"])
    if not row_idx:
        row_idx = find_row_by_key(ws_prov, HEADERS_PROV, "id", p["id"])

    if row_idx:
        delete_row(ws_prov, row_idx)

    proveedores.remove(p)
    messagebox.showinfo("OK", "Proveedor eliminado.")


# =========================
#   CRUD Items
# =========================

def menu_items():
    w = tk.Toplevel()
    w.title("Items")
    w.geometry("420x260")
    crear_fondo_imagen(w)

    card = tk.Frame(w, bg="white", bd=2, relief="ridge")
    card.place(relx=0.5, rely=0.5, anchor="center")

    tk.Button(card, text="Agregar item", width=30, command=agregar_item,
              bg=AZUL_REY, fg="white").pack(pady=6)
    tk.Button(card, text="Modificar item", width=30, command=modificar_item,
              bg=AZUL_REY, fg="white").pack(pady=6)
    tk.Button(card, text="Eliminar item", width=30, command=eliminar_item,
              bg=AZUL_REY, fg="white").pack(pady=6)


def seleccionar_item_por_sku():
    if not items:
        messagebox.showinfo("Items", "No hay items.")
        return None
    lista = [i["sku"] for i in items]
    elegido = pedir_opcion_lista("Seleccionar", "SKU:", lista)
    if not elegido:
        return None
    for it in items:
        if it["sku"] == elegido:
            return it
    return None


def agregar_item():
    global items
    desc = pedir_valor("Nuevo item", "Descripción SAE:")
    if not desc:
        return
    sku = pedir_valor("Nuevo item", "SKU:")
    if not sku:
        return
    if any(x["sku"] == sku for x in items):
        return messagebox.showerror("Error", "SKU ya existe.")

    familia = pedir_valor("Nuevo item", "Familia:") or ""
    provs = get_proveedor_nombres()
    proveedor_sel = pedir_opcion_lista("Proveedor", "Seleccione:", provs) if provs else pedir_valor("Proveedor", "Proveedor:")
    stock = pedir_valor("Stock", "Stock inicial:") or "0"
    precio = pedir_valor("Precio", "Precio:") or "0"

    try:
        int(stock)
        float(precio)
    except:
        return messagebox.showerror("Error", "Valores inválidos.")

    nuevo = {
        "sku": sku,
        "desc_sae": desc,
        "familia": familia,
        "proveedor": proveedor_sel,
        "stock": str(stock),
        "precio": str(precio)
    }

    items.append(nuevo)
    append_dict(ws_items, HEADERS_ITEMS, nuevo)
    messagebox.showinfo("OK", "Item agregado.")


def modificar_item():
    global items
    it = seleccionar_item_por_sku()
    if not it:
        return

    original_sku = it["sku"]
    it["desc_sae"] = pedir_valor("Modificar", "Descripción:", it["desc_sae"]) or it["desc_sae"]
    it["familia"] = pedir_valor("Modificar", "Familia:", it["familia"]) or it["familia"]

    provs = get_proveedor_nombres()
    if provs:
        it["proveedor"] = pedir_opcion_lista("Proveedor", "Seleccione:", provs) or it["proveedor"]
    else:
        it["proveedor"] = pedir_valor("Proveedor", "Proveedor:", it["proveedor"]) or it["proveedor"]

    stock_txt = pedir_valor("Modificar", "Stock:", it["stock"]) or it["stock"]
    precio_txt = pedir_valor("Modificar", "Precio:", it["precio"]) or it["precio"]

    try:
        int(stock_txt)
        float(precio_txt)
    except:
        return messagebox.showerror("Error", "Valor inválido.")

    it["stock"] = str(stock_txt)
    it["precio"] = str(precio_txt)

    row_idx = find_row_by_key(ws_items, HEADERS_ITEMS, "sku", original_sku)
    if row_idx:
        update_row(ws_items, HEADERS_ITEMS, row_idx, it)

    items[:] = dicts_from_ws(ws_items, HEADERS_ITEMS)
    messagebox.showinfo("OK", "Item actualizado.")


def eliminar_item():
    global items
    it = seleccionar_item_por_sku()
    if not it:
        return

    if not confirmar_si_no("Confirmar", f"¿Eliminar SKU {it['sku']}?"):
        return

    row_idx = find_row_by_key(ws_items, HEADERS_ITEMS, "sku", it["sku"])
    if row_idx:
        delete_row(ws_items, row_idx)

    items.remove(it)
    messagebox.showinfo("OK", "Item eliminado.")


# =========================
#   MAINLOOP
# =========================
ventana_login.mainloop()
