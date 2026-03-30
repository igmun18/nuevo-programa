import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import filedialog, messagebox
from PIL import Image
import os
import openpyxl
import re
from openpyxl import load_workbook
from copy import copy
from dateutil.relativedelta import relativedelta
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

class App(ctk.CTk, TkinterDnD.Tk):
    def __init__(self):

        ctk.CTk.__init__(self)
        TkinterDnD.Tk.__init__(self)

        ctk.set_appearance_mode("dark")
        self.title("Mi primer programa c:")
        self.geometry("700x750")
        self.configure(fg_color="#242424")

        # ---------------- ICONO ----------------
        try:
            self.icono_excel = ctk.CTkImage(
                light_image=Image.open("excel_icon.png"),
                dark_image=Image.open("excel_icon.png"),
                size=(40, 40)
            )
        except:
            self.icono_excel = None

        self.rutas_archivos = []

        # ---------------- CONTENEDOR PRINCIPAL ----------------
        self.container = ctk.CTkFrame(self, fg_color="transparent")
        self.container.pack(fill="both", expand=True, padx=20, pady=20)

        # ---------------- SECCIÓN BUSCAR ----------------
        frame_path = ctk.CTkFrame(self.container)
        frame_path.pack(fill="x", pady=10)

        btn_buscar = ctk.CTkButton(frame_path, text="Buscar", width=100, command=self.buscar_archivo)
        btn_buscar.pack(side="left", padx=10)

        self.entry_path = ctk.CTkEntry(frame_path, placeholder_text="Ruta del archivo...")
        self.entry_path.pack(side="left", padx=10, expand=True, fill="x")

        # ---------------- SECCIÓN PORCENTAJE ----------------
        frame_porcentaje = ctk.CTkFrame(self.container)
        frame_porcentaje.pack(fill="x", pady=10)

        label = ctk.CTkLabel(frame_porcentaje, text="Porcentaje de aumento (%):")
        label.pack(side="left", padx=10)

        self.entry_porcentaje = ctk.CTkEntry(frame_porcentaje, width=100)
        self.entry_porcentaje.insert(0, "15")
        self.entry_porcentaje.pack(side="left", padx=10)

        # ---------------- DRAG & DROP ----------------
        self.frame_drop = ctk.CTkFrame(
            self.container,
            border_width=2,
            border_color="#3B3B3B"
        )
        self.frame_drop.pack(fill="both", expand=True, pady=10)

        self.label_drop = ctk.CTkLabel(
            self.frame_drop,
            text="Arrastra archivos Excel aquí o usa 'Buscar'"
        )
        self.label_drop.pack(pady=10)

        # SCROLLABLE FRAME
        self.lista_iconos_frame = ctk.CTkScrollableFrame(
            self.frame_drop,
            label_text="Archivos seleccionados"
        )
        self.lista_iconos_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # ---------------- BOTÓN ----------------
        self.btn_ejecutar = ctk.CTkButton(
            self.container,
            text="APLICAR INCREMENTO",
            fg_color="#0B1575",
            hover_color="#2808B4",
            height=45,
            command=self.ejecutar_macro
        )
        self.btn_ejecutar.pack(fill="x", pady=15)

        # DRAG & DROP
        self.frame_drop.drop_target_register(DND_FILES)
        self.frame_drop.dnd_bind('<<Drop>>', self.al_soltar_archivo)

    # ---------------- FUNCIONES ----------------

    def buscar_archivo(self):
        archivos = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if archivos:
            self.agregar_a_lista(archivos)

    def al_soltar_archivo(self, event):
        rutas = self.tk.splitlist(event.data)
        self.agregar_a_lista(rutas)

    def agregar_a_lista(self, rutas):
        for ruta in rutas:
            ruta = ruta.strip('{}')
            if ruta not in self.rutas_archivos and ruta.lower().endswith(('.xlsx', '.xls')):
                self.rutas_archivos.append(ruta)

        self.renderizar_iconos()

    # ---------------- GRILLA DE ICONOS ----------------

    def renderizar_iconos(self):
        # limpiar todo
        for widget in self.lista_iconos_frame.winfo_children():
            widget.destroy()

        columnas = 4

        for index, ruta in enumerate(self.rutas_archivos):
            nombre = os.path.basename(ruta)

            fila = index // columnas
            col = index % columnas

            item = ctk.CTkFrame(self.lista_iconos_frame, width=120, height=120)
            item.grid(row=fila, column=col, padx=10, pady=10)
            item.grid_propagate(False)

            # icono
            label_img = ctk.CTkLabel(
                item,
                image=self.icono_excel,
                text="" if self.icono_excel else "📄"
            )
            label_img.pack(pady=(10, 5))

            # nombre corto
            nombre_corto = nombre[:12] + "..." if len(nombre) > 12 else nombre
            label_nombre = ctk.CTkLabel(item, text=nombre_corto)
            label_nombre.pack()

            # botón eliminar
            btn_eliminar = ctk.CTkButton(
                item,
                text="✕",
                width=20,
                height=20,
                fg_color="#A12121",
                hover_color="#E63946",
                command=lambda r=ruta: self.eliminar_archivo(r)
            )
            btn_eliminar.place(relx=1, rely=0, anchor="ne")

    def eliminar_archivo(self, ruta):
        if ruta in self.rutas_archivos:
            self.rutas_archivos.remove(ruta)

        self.renderizar_iconos()

    # ---------------- MACRO ----------------

    def ejecutar_macro(self):

        try:
            factor = 1 + (float(self.entry_porcentaje.get()) / 100)
        except ValueError:
            messagebox.showerror("Error", "Ingresa un porcentaje válido.")
            return  # ✔ ahora sí está bien

        if not self.rutas_archivos:
            messagebox.showwarning("Atención", "No hay archivos seleccionados.")
            return  # ✔ bien indentado

        exitos = 0

        for ruta in self.rutas_archivos:
            try:
                wb = openpyxl.load_workbook(ruta,data_only=True)

                for ws in wb.worksheets:

                    fila_header = None
                    col_fecha = None

                    # 🔍 BUSCAR HEADER
                    for fila in range(1, 15):
                        col_valores = ws.max_column

                        valor = ws.cell(row=fila, column=col_valores).value

                        if isinstance(valor, datetime):
                            fila_header = fila
                            col_fecha = col_valores
                            break

                        elif isinstance(valor, str):
                            try:
                                datetime.strptime(valor, "%d/%m/%Y")
                                fila_header = fila
                                col_fecha = col_valores
                                break
                            except:
                                pass

                        if fila_header:
                            break

                    if not fila_header:
                        continue

                    ultima_columna = ws.max_column
                    nueva_columna = ultima_columna + 1
                    col_valores2 = ultima_columna

                    header_origen = ws.cell(row=fila_header, column=col_valores2)
                    header_destino = ws.cell(row=fila_header, column=nueva_columna)

                    fecha = header_origen.value

                    if isinstance(fecha, datetime):
                        pass

                    elif isinstance(fecha, str):
                        try:
                            fecha = datetime.strptime(fecha, "%d/%m/%Y")
                        except:
                            fecha = None

                    elif isinstance(fecha, (int, float)):
                        from openpyxl.utils.datetime import from_excel
                        fecha = from_excel(fecha)

                    if not fecha:
                        continue
                    
                    nueva_fecha = fecha + relativedelta(months=1)
                    header_destino.value = nueva_fecha
                    header_destino.number_format = "DD/MM/YYYY"

                    # copiar estilo
                    header_destino.font = copy(header_origen.font)
                    header_destino.border = copy(header_origen.border)
                    header_destino.fill = copy(header_origen.fill)
                    header_destino.alignment = copy(header_origen.alignment)

                    # 🔢 DATOS
                    for fila in range(fila_header + 1, ws.max_row + 1):

                        celda_origen = ws.cell(row=fila, column=col_fecha)
                        celda_destino = ws.cell(row=fila, column=nueva_columna)

                        valor = celda_origen.value

                        if isinstance(valor, (int, float)):
                            celda_destino.value = valor * factor

                        elif isinstance(valor, str):
                            match = re.search(r'(NN\s*[xX]\s*)(\d+)', valor)

                            if match:
                                prefijo = match.group(1)   # "NN X "
                                numero = match.group(2)    # "1500"

                                nuevo_valor = float(numero) * factor

                                # si querés entero:
                                nuevo_valor = float(nuevo_valor)

                                celda_destino.value = f"{prefijo}{nuevo_valor}"
                            else:
                                celda_destino.value = valor       
                        else:
                            celda_destino.value = valor

                        if celda_origen.has_style:
                            celda_destino.font = copy(celda_origen.font)
                            celda_destino.border = copy(celda_origen.border)
                            celda_destino.fill = copy(celda_origen.fill)
                            celda_destino.number_format = copy(celda_origen.number_format)
                            celda_destino.alignment = copy(celda_origen.alignment)

                    # ancho columna
                    letra_orig = get_column_letter(col_fecha)
                    letra_dest = get_column_letter(nueva_columna)
                    ws.column_dimensions[letra_dest].width = ws.column_dimensions[letra_orig].width
                print(nueva_fecha)       
                nuevo_nombre = ruta.replace(".xlsx", f"_{nueva_fecha}.xlsx")
                wb.save(nuevo_nombre)

                exitos += 1

            except Exception as e:
                print(f"Error en {ruta}: {e}")

        messagebox.showinfo("Hecho", f"Se procesaron {exitos} archivos correctamente.")

if __name__ == "__main__": 
    app = App() 
    app.mainloop()