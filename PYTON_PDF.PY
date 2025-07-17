import fitz  # PyMuPDF
import tkinter as tk
from tkinter import filedialog, simpledialog
import os

# Función para seleccionar archivos
def seleccionar_archivo(titulo, tipos):
    root = tk.Tk()
    root.withdraw()
    archivo = filedialog.askopenfilename(title=titulo, filetypes=tipos)
    return archivo

# Seleccionar archivo PDF
pdf_path = seleccionar_archivo("Selecciona el archivo PDF", [("Archivos PDF", "*.pdf")])
if not pdf_path:
    print("No se seleccionó PDF")
    exit()

# Seleccionar imagen
img_path = seleccionar_archivo("Selecciona la imagen", [("Imágenes", "*.jpg *.jpeg *.png")])
if not img_path:
    print("No se seleccionó imagen")
    exit()

# Pedir datos de ubicación y tamaño
def pedir_entero(mensaje, valor_defecto):
    return simpledialog.askinteger("Posición", mensaje, initialvalue=valor_defecto)

x = pedir_entero("Posición X en puntos (1 punto = 1/72 pulgada)", 100)
y = pedir_entero("Posición Y", 100)
ancho = pedir_entero("Ancho de la imagen", 200)
alto = pedir_entero("Alto de la imagen", 200)
pagina = pedir_entero("Número de página (comienza en 1)", 1) - 1

# Abrir el PDF
doc = fitz.open(pdf_path)

if pagina >= len(doc):
    print("Número de página fuera de rango.")
    exit()

# Insertar imagen
page = doc[pagina]
rect = fitz.Rect(x, y, x + ancho, y + alto)
page.insert_image(rect, filename=img_path)

# Guardar como nuevo archivo
base, ext = os.path.splitext(pdf_path)
output_pdf = base + "_con_imagen.pdf"
doc.save(output_pdf)
doc.close()

print("✅ PDF guardado como:", output_pdf)
