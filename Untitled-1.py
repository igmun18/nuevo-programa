import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import filedialog, messagebox
from PIL import Image
import os
import openpyxl

class App(ctk.CTk, TkinterDnD.Tk):
    def __init__(self):
        ctk.CTk.__init__(self)
        TkinterDnD.Tk.__init__(self)
        
        ctk.set_appearance_mode("dark")
        self.title("Macro Excel Pro")
        self.geometry("600x700")

        # --- 1. CARGA DE ICONO ---
        try:
            self.icono_excel = ctk.CTkImage(light_image=Image.open("excel_icon.png"),
                                            dark_image=Image.open("excel_icon.png"),
                                            size=(25, 25))
        except:
            self.icono_excel = None 

        self.rutas_archivos = []

        # --- 2. SECCIÓN DE RUTA Y BUSCAR ---
        frame_path = ctk.CTkFrame(self)
        frame_path.pack(pady=20, padx=20, fill="x")

        self.entry_path = ctk.CTkEntry(frame_path, placeholder_text="Ruta del archivo...")
        self.entry_path.pack(side="left", padx=10, expand=True, fill="x")

        btn_buscar = ctk.CTkButton(frame_path, text="Buscar", width=100, command=self.buscar_archivo)
        btn_buscar.pack(side="right", padx=10)

        # --- 3. SECCIÓN DE PORCENTAJE ---
        ctk.CTkLabel(self, text="Porcentaje de aumento (%):", font=("Arial", 13, "bold")).pack(pady=(10, 0))
        self.entry_porcentaje = ctk.CTkEntry(self, width=120, placeholder_text="Ej: 15")
        self.entry_porcentaje.insert(0, "15") 
        self.entry_porcentaje.pack(pady=10)

        # --- 4. ÁREA DE DRAG & DROP ---
        self.frame_drop = ctk.CTkFrame(self, height=200, border_width=2, border_color="#3B3B3B")
        self.frame_drop.pack(pady=10, padx=20, fill="x")
        
        self.label_drop = ctk.CTkLabel(self.frame_drop, text="O arrastra tus archivos aquí ↓")
        self.label_drop.pack(pady=10)

        # Frame con scroll para ver los iconos
        self.lista_iconos_frame = ctk.CTkScrollableFrame(self, label_text="Archivos seleccionados", height=250)
        self.lista_iconos_frame.pack(pady=10, padx=20, fill="both", expand=True)

        # --- 5. BOTÓN EJECUTAR --- 
        self.btn_ejecutar = ctk.CTkButton(self, text="APLICAR MACRO", fg_color="#0B1575", hover_color="#2808B4", 
                                     height=45, font=("Arial", 14, "bold"), command=self.ejecutar_macro)
        self.btn_ejecutar.pack(pady=20)

        self.frame_drop.drop_target_register(DND_FILES)
        self.frame_drop.dnd_bind('<<Drop>>', self.al_soltar_archivo)

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
                self.crear_item_lista(ruta)

    def crear_item_lista(self, ruta):
        nombre = os.path.basename(ruta)
        
        # Contenedor de la fila
        fila = ctk.CTkFrame(self.lista_iconos_frame, fg_color="transparent")
        fila.pack(fill="x", pady=2)
        
        # Icono de Excel
        label_img = ctk.CTkLabel(fila, image=self.icono_excel, text="") if self.icono_excel else ctk.CTkLabel(fila, text="📄")
        label_img.pack(side="left", padx=5)
        
        # Nombre del archivo
        label_nombre = ctk.CTkLabel(fila, text=nombre, anchor="w")
        label_nombre.pack(side="left", padx=5, expand=True, fill="x")

        # Botón eliminar (Basurero)
        btn_eliminar = ctk.CTkButton(fila, text="🗑", width=30, height=30, 
                                     fg_color="#A12121", hover_color="#E63946",
                                     command=lambda r=ruta, f=fila: self.eliminar_archivo(r, f))
        btn_eliminar.pack(side="right", padx=5)

    def eliminar_archivo(self, ruta, frame_fila):
        # Eliminar de la lista lógica
        if ruta in self.rutas_archivos:
            self.rutas_archivos.remove(ruta)
        # Eliminar de la interfaz visual
        frame_fila.destroy()

    def ejecutar_macro(self):
        try:
            factor = 1 + (float(self.entry_porcentaje.get()) / 100)
        except ValueError:
            messagebox.showerror("Error", "Ingresa un porcentaje válido.")
            return

        if not self.rutas_archivos:
            messagebox.showwarning("Atención", "No hay archivos seleccionados.")
            return

        exitos = 0
        for ruta in self.rutas_archivos:
            try:
                wb = openpyxl.load_workbook(ruta)
                for hoja in wb.worksheets:
                    for fila in hoja.iter_rows():
                        for celda in fila:
                            if isinstance(celda.value, (int, float)):
                                celda.value *= factor
                wb.save(ruta)
                exitos += 1
            except Exception as e:
                print(f"Error en {ruta}: {e}")

        messagebox.showinfo("Hecho", f"Se procesaron {exitos} archivos correctamente.")

if __name__ == "__main__":
    app = App()
    app.mainloop()