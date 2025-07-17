import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import io
import os

# Manejo de compatibilidad Pillow >=10 y <10
try:
    RESAMPLE_METHOD = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE_METHOD = Image.LANCZOS

class PDFEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Insertar Timbre en PDF")
        self.root.geometry("460x500")
        self.pdf_path = ""
        self.img_path = ""
        self.preview_coords = None
        self.img_widget = None
        self.drag_data = {"x": 0, "y": 0}
        self.insert_mode = tk.StringVar(value="Una página")
        self.crear_widgets()

    def crear_widgets(self):
        tk.Label(self.root, text="Insertar Imagen en PDF", font=("Arial", 16, "bold")).pack(pady=10)
        tk.Button(self.root, text="Seleccionar PDF", command=self.seleccionar_pdf).pack(pady=5)
        self.label_pdf = tk.Label(self.root, text="Ningún PDF seleccionado")
        self.label_pdf.pack()
        tk.Button(self.root, text="Seleccionar Imagen", command=self.seleccionar_imagen).pack(pady=5)
        self.label_img = tk.Label(self.root, text="Ninguna imagen seleccionada")
        self.label_img.pack()
        self.ancho = self.crear_entrada("Ancho imagen (cm):", "7")
        self.alto = self.crear_entrada("Alto imagen (cm):", "4")
        self.pagina = self.crear_entrada("Página base (desde 1):", "1")

        # Modo de inserción
        modos = ["Una página", "Varias páginas", "Todas las páginas"]
        tk.Label(self.root, text="Insertar imagen en:").pack(pady=5)
        tk.OptionMenu(self.root, self.insert_mode, *modos).pack()

        self.otras_paginas = self.crear_entrada("Si aplica: Páginas (1,3,5)", "")

        tk.Button(self.root, text="Vista previa (mover imagen)", command=self.vista_previa).pack(pady=5)
        tk.Button(self.root, text="Insertar imagen", command=self.insertar_imagen).pack(pady=15)

    def crear_entrada(self, texto, valor_defecto):
        frame = tk.Frame(self.root)
        frame.pack(pady=2)
        tk.Label(frame, text=texto, width=25, anchor="w").pack(side="left")
        entrada = tk.Entry(frame)
        entrada.insert(0, valor_defecto)
        entrada.pack(side="left")
        return entrada

    def seleccionar_pdf(self):
        self.pdf_path = filedialog.askopenfilename(title="Seleccionar PDF", filetypes=[("PDF Files", "*.pdf")])
        self.label_pdf.config(text=os.path.basename(self.pdf_path) if self.pdf_path else "Ningún PDF seleccionado")

    def seleccionar_imagen(self):
        self.img_path = filedialog.askopenfilename(title="Seleccionar Imagen", filetypes=[("Imágenes", "*.png *.jpg *.jpeg")])
        self.label_img.config(text=os.path.basename(self.img_path) if self.img_path else "Ninguna imagen seleccionada")

    def cm_a_puntos(self, cm):
        return cm * 72 / 2.54

    def puntos_a_cm(self, puntos):
        return puntos * 2.54 / 72

    def vista_previa(self):
        try:
            if not self.pdf_path or not self.img_path:
                messagebox.showerror("Error", "Debe seleccionar un PDF y una imagen.")
                return

            ancho_cm = float(self.ancho.get())
            alto_cm = float(self.alto.get())
            pagina = int(self.pagina.get()) - 1

            doc = fitz.open(self.pdf_path)
            if pagina < 0 or pagina >= len(doc):
                raise ValueError("Número de página fuera de rango.")

            pix = doc[pagina].get_pixmap(dpi=100)
            img_bytes = pix.tobytes("ppm")
            fondo = Image.open(io.BytesIO(img_bytes))

            imagen = Image.open(self.img_path)
            escala = 100 / 2.54
            ancho_px = int(ancho_cm * escala)
            alto_px = int(alto_cm * escala)
            imagen = imagen.resize((ancho_px, alto_px), RESAMPLE_METHOD)

            self.fondo_tk = ImageTk.PhotoImage(fondo)
            self.imagen_tk = ImageTk.PhotoImage(imagen)

            self.preview_window = tk.Toplevel(self.root)
            self.preview_window.title("Vista previa (arrastra la imagen)")
            self.canvas = tk.Canvas(self.preview_window, width=fondo.width, height=fondo.height)
            self.canvas.pack()
            self.canvas.create_image(0, 0, anchor="nw", image=self.fondo_tk)

            x_inicial = (fondo.width - ancho_px) // 2
            y_inicial = (fondo.height - alto_px) // 2
            self.preview_coords = (x_inicial, y_inicial)

            self.img_widget = self.canvas.create_image(x_inicial, y_inicial, anchor="nw", image=self.imagen_tk)
            self.canvas.tag_bind(self.img_widget, "<Button-1>", self.iniciar_arrastre)
            self.canvas.tag_bind(self.img_widget, "<B1-Motion>", self.mover_imagen)

        except Exception as e:
            messagebox.showerror("Error en vista previa", str(e))

    def iniciar_arrastre(self, event):
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y

    def mover_imagen(self, event):
        dx = event.x - self.drag_data["x"]
        dy = event.y - self.drag_data["y"]
        self.canvas.move(self.img_widget, dx, dy)
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        coords = self.canvas.coords(self.img_widget)
        self.preview_coords = (int(coords[0]), int(coords[1]))

    def insertar_imagen(self):
        try:
            if not self.pdf_path or not self.img_path or not self.preview_coords:
                messagebox.showerror("Error", "Debe realizar la vista previa y mover la imagen antes de insertar.")
                return

            ancho_cm = float(self.ancho.get())
            alto_cm = float(self.alto.get())
            x_px, y_px = self.preview_coords
            escala = 100 / 2.54
            x_cm = x_px / escala
            y_cm = y_px / escala

            x = self.cm_a_puntos(x_cm)
            y = self.cm_a_puntos(y_cm)
            ancho = self.cm_a_puntos(ancho_cm)
            alto = self.cm_a_puntos(alto_cm)
            rect = fitz.Rect(x, y, x + ancho, y + alto)

            doc = fitz.open(self.pdf_path)
            total_paginas = len(doc)

            modo = self.insert_mode.get()
            paginas_a_insertar = []

            if modo == "Una página":
                pagina = int(self.pagina.get()) - 1
                if pagina < 0 or pagina >= total_paginas:
                    raise ValueError("Página fuera de rango.")
                paginas_a_insertar = [pagina]

            elif modo == "Varias páginas":
                entradas = self.otras_paginas.get()
                paginas_a_insertar = [int(p.strip()) - 1 for p in entradas.split(",") if p.strip().isdigit()]
                for p in paginas_a_insertar:
                    if p < 0 or p >= total_paginas:
                        raise ValueError(f"Página {p + 1} fuera de rango.")

            elif modo == "Todas las páginas":
                paginas_a_insertar = list(range(total_paginas))

            for p in paginas_a_insertar:
                doc[p].insert_image(rect, filename=self.img_path)

            nuevo_nombre = os.path.splitext(self.pdf_path)[0] + "_con_imagen.pdf"
            doc.save(nuevo_nombre)
            doc.close()

            messagebox.showinfo("Éxito", f"Imagen insertada correctamente.\nGuardado como:\n{nuevo_nombre}")

        except Exception as e:
            messagebox.showerror("Error al insertar imagen", str(e))


# Ejecutar aplicación
if __name__ == "__main__":
    print("Aplicación iniciada...")
    root = tk.Tk()
    app = PDFEditorApp(root)
    root.mainloop()

