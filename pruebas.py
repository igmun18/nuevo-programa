import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD
from PIL import Image
import os

class App(ctk.CTk, TkinterDnD.Tk):
    def __init__(self):
        ctk.CTk.__init__(self)
        TkinterDnD.Tk.__init__(self)
        
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("Arrastra tu Excel")
        self.geometry("500x400")

        # 1. Cargar el icono de Excel (asegúrate de que el archivo esté en la misma carpeta)
        # Redimensionamos a 32x32 para que se vea bien en la lista
        self.icono_excel = ctk.CTkImage(light_image=Image.open("excel_icon.png"),
                                        dark_image=Image.open("excel_icon.png"),
                                        size=(30, 30))

        # Contenedor principal de arrastre
        self.frame_drop = ctk.CTkFrame(self)
        self.frame_drop.pack(pady=10, padx=20, fill="x")

        self.label_instruccion = ctk.CTkLabel(self.frame_drop, text="Suelte sus archivos Excel aquí", font=("Arial", 14, "bold"))
        self.label_instruccion.pack(pady=20)

        # 2. Frame con scroll para mostrar los iconos de forma ordenada
        self.lista_iconos_frame = ctk.CTkScrollableFrame(self, label_text="Archivos cargados")
        self.lista_iconos_frame.pack(pady=10, padx=20, fill="both", expand=True)

        # Configurar el Drag & Drop
        self.frame_drop.drop_target_register(DND_FILES)
        self.frame_drop.dnd_bind('<<Drop>>', self.al_soltar_archivo)

    def al_soltar_archivo(self, event):
        # Limpiar la lista previa si deseas que solo se vean los nuevos, 
        # o quitar esto si quieres ir acumulándolos.
        # for widget in self.lista_iconos_frame.winfo_children():
        #     widget.destroy()

        rutas_completas = self.tk.splitlist(event.data)
        
        for ruta in rutas_completas:
            nombre_archivo = os.path.basename(ruta) # Extrae el nombre más limpiamente
            
            # Crear un pequeño "item" para la lista
            item_frame = ctk.CTkFrame(self.lista_iconos_frame, fg_color="transparent")
            item_frame.pack(fill="x", pady=2, padx=5)

            # Insertar el Icono
            img_label = ctk.CTkLabel(item_frame, image=self.icono_excel, text="")
            img_label.pack(side="left", padx=5)

            # Insertar el Nombre
            texto_label = ctk.CTkLabel(item_frame, text=nombre_archivo, font=("Segoe UI", 12))
            texto_label.pack(side="left", padx=5)
        
        print(f"Se han añadido {len(rutas_completas)} archivos.")

if __name__ == "__main__":
    app = App()
    app.mainloop()