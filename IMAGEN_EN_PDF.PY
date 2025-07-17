import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, messagebox
import os

class PDFEditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Insertar Timbre en PDF")
        self.root.geometry("420x480")
        self.pdf_path = ""
        self.img_path = ""

        self.crear_widgets()

    def crear_widgets(self):
        tk.Label(self.root, text="Insertar Imagen en PDF", font=("Arial", 16, "bold")).pack(pady=10)

        tk.Button(self.root, text="Seleccionar PDF", command=self.seleccionar_pdf).pack(pady=5)
        self.label_pdf = tk.Label(self.root, text="Ningún PDF seleccionado")
        self.label_pdf.pack()

        tk.Button(self.root, text="Seleccionar Imagen", command=self.seleccionar_imagen).pack(pady=5)
        self.label_img = tk.Label(self.root, text="Ninguna imagen seleccionada")
        self.label_img.pack()

        # Entradas en centímetros
        self.x = self.crear_entrada("Posición X (cm):", "3")
        self.y = self.crear_entrada("Posición Y (cm):", "3")
        self.ancho = self.crear_entrada("Ancho imagen (cm):", "7")
        self.alto = self.crear_entrada("Alto imagen (cm):", "4")
        self.pagina = self.crear_entrada("Número de página (desde 1):", "1")

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
        return cm * 72 / 2.54  # ≈ 28.3465

    def insertar_imagen(self):
        try:
            if not self.pdf_path or not self.img_path:
                messagebox.showerror("Error", "Debe seleccionar un PDF y una imagen.")
                return

            x_cm = float(self.x.get())
            y_cm = float(self.y.get())
            ancho_cm = float(self.ancho.get())
            alto_cm = float(self.alto.get())
            pagina = int(self.pagina.get()) - 1  # páginas comienzan en 0

            # Convertir a puntos
            x = self.cm_a_puntos(x_cm)
            y = self.cm_a_puntos(y_cm)
            ancho = self.cm_a_puntos(ancho_cm)
            alto = self.cm_a_puntos(alto_cm)

            doc = fitz.open(self.pdf_path)
            if pagina >= len(doc):
                messagebox.showerror("Error", "El número de página excede las páginas del PDF.")
                return

            page = doc[pagina]
            rect = fitz.Rect(x, y, x + ancho, y + alto)
            page.insert_image(rect, filename=self.img_path)

            nuevo_nombre = os.path.splitext(self.pdf_path)[0] + "_con_imagen.pdf"
            doc.save(nuevo_nombre)
            doc.close()

            messagebox.showinfo("Éxito", f"Imagen insertada exitosamente.\nArchivo guardado como:\n{nuevo_nombre}")
        except Exception as e:
            messagebox.showerror("Error", str(e))

# Ejecutar
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFEditorApp(root)
    root.mainloop()
