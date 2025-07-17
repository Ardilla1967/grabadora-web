import win32com.client
import pandas as pd

# Iniciar Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Identificar el nombre de la lista (ajustar si necesario)
address_list = namespace.AddressLists.Item("Lista global de direcciones")

# Crear una lista para guardar los datos
datos = []

# Recorrer las entradas del directorio
for i in range(1, address_list.AddressEntries.Count + 1):
    entry = address_list.AddressEntries.Item(i)
    datos.append({
        'Nombre': entry.Name,
        'Correo': entry.Address
    })

# Convertir a DataFrame de pandas
df = pd.DataFrame(datos)

# Guardar en un archivo Excel
ruta_salida = "C:/Users/mtalal/OneDrive - Chileatiende/Escritorio/contactos_outlook.xlsx"
df.to_excel(ruta_salida, index=False)

print(f"Archivo Excel creado correctamente en:\n{ruta_salida}")
