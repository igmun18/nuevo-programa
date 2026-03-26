import customtkinter as ctk
from tkinter import filedialog, messagebox
from openpyxl import load_workbook
from copy import copy
from dateutil.relativedelta import relativedelta
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# --- 1. INTERFAZ Y FUNCIONES DE APOYO ---

def buscar_archivo():
    archivo = filedialog.askopenfilename(filetypes=[("Archivos de Excel", "*.xlsx")])
    if archivo:
        entry_path.delete(0, ctk.END)
        entry_path.insert(0, archivo)

def ejecutar_macro():
    # Capturamos los datos de la interfaz para usarlos en TU código
    path = entry_path.get()
    porcentaje_input = entry_porcentaje.get()

    if not path:
        messagebox.showwarning("Google", "Selecciona un archivo primero.")
        return

    try:
        # Convertimos tu porcentaje (ej: 15) al factor que usas (1.15)
        factor_multiplicador = 1 + (float(porcentaje_input) / 100)
        
        # --- AQUÍ EMPIEZA TU CÓDIGO ORIGINAL ---
        wb = load_workbook(path)

        for ws in wb.worksheets:
            ultima_columna = ws.max_column

            header_origen = ws.cell(row=1, column=ultima_columna)
            header_destino = ws.cell(row=1, column=ultima_columna + 1)

            valor_header = header_origen.value
            fecha = None

            if isinstance(valor_header, str):
                try:
                    fecha = datetime.strptime(valor_header, "%d/%m/%Y")
                except:
                    fecha = datetime.now()
            elif isinstance(valor_header, datetime):
                fecha = valor_header

            if fecha:
                nueva_fecha = fecha + relativedelta(months=1)
                header_destino.value = nueva_fecha
                header_destino.number_format = "DD/MM/YYYY"

            header_destino.font = copy(header_origen.font)
            header_destino.border = copy(header_origen.border)
            header_destino.fill = copy(header_origen.fill)

            header_destino.alignment = Alignment(
                wrap_text=True, 
                horizontal=header_origen.alignment.horizontal, 
                vertical=header_origen.alignment.vertical
            ) 

            for i in range(2, ws.max_row + 1):
                celda_origen = ws.cell(row=i, column=ultima_columna)
                celda_destino = ws.cell(row=i, column=ultima_columna + 1)       

                valor = celda_origen.value

                try:
                    if isinstance(valor, (int, float)):
                        # Cambiamos el 1.15 fijo por tu factor dinámico
                        celda_destino.value = valor * factor_multiplicador
                    else:
                        celda_destino.value = valor
                except:
                    celda_destino.value = valor
            
                if celda_origen.has_style:
                    celda_destino.font = copy(celda_origen.font)
                    celda_destino.border = copy(celda_origen.border)
                    celda_destino.fill = copy(celda_origen.fill)
                    celda_destino.number_format = copy(celda_origen.number_format)

                    nueva_ali = copy(celda_origen.alignment)
                    nueva_ali.wrap_text = True
                    celda_destino.alignment = nueva_ali
            
            letra_orig = get_column_letter(ultima_columna)
            letra_dest = get_column_letter(ultima_columna + 1)
            ws.column_dimensions[letra_dest].width = ws.column_dimensions[letra_orig].width

        # Guardamos con un nombre que indique que ya está listo
        wb.save(path)
        messagebox.showinfo("Google", f"¡Listo! Archivo guardado en:\n{path}")
        # --- AQUÍ TERMINA TU CÓDIGO ORIGINAL ---

    except Exception as e:
        messagebox.showerror("Error", f"Algo falló: {e}")

# --- 2. DISEÑO DE LA VENTANA (CustomTkinter) ---
ctk.set_appearance_mode("dark")  # Opciones: "dark", "light", "system"
ctk.set_default_color_theme("blue") # Opciones: "blue", "green", "dark-blue"
app = ctk.CTk()
app.title("Automatizador de incrementos")
app.geometry("500x350")

# Buscador de archivo
ctk.CTkLabel(app, text="Selecciona tu archivo Excel:", font=("Arial", 13, "bold")).pack(pady=(20, 0))
frame_path = ctk.CTkFrame(app)
frame_path.pack(pady=10, padx=20, fill="x")

entry_path = ctk.CTkEntry(frame_path, placeholder_text="Ruta del archivo...")
entry_path.pack(side="left", padx=10, expand=True, fill="x")

btn_buscar = ctk.CTkButton(frame_path, text="Buscar", width=100, command=buscar_archivo)
btn_buscar.pack(side="right", padx=10)

# Porcentaje
ctk.CTkLabel(app, text="Porcentaje de aumento (%):", font=("Arial", 13, "bold")).pack(pady=(20, 0))
entry_porcentaje = ctk.CTkEntry(app, width=120, placeholder_text="Ej: 15")
entry_porcentaje.insert(0, "15") # Valor inicial por defecto
entry_porcentaje.pack(pady=10)

# Botón de ejecución
btn_ejecutar = ctk.CTkButton(app, text="APLICAR MACRO", fg_color="#0B1575", hover_color="#2808B4", 
                             height=45, font=("Arial", 14, "bold"), command=ejecutar_macro)
btn_ejecutar.pack(pady=30)

app.mainloop()