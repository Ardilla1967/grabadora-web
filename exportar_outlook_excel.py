import win32com.client
import pandas as pd
import os

print("[INFO] Conectando a Outlook y cargando 'Offline Global Address List'...")

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Usamos el nombre exacto detectado
try:
    gal = outlook.AddressLists["Offline Global Address List"]
except:
    print("[ERROR] No se encontró la libreta con ese nombre exacto.")
    exit()

print(f"[OK] Libreta encontrada: {gal.Name}")
entries = gal.AddressEntries
datos = []

# Recorremos todos los contactos de la GAL
for i in range(entries.Count):
    try:
        entry = entries.Item(i + 1)
        nombre = entry.Name
        correo = ""

        if entry.AddressEntryUserType == 0:
            try:
                exchange_user = entry.GetExchangeUser()
                if exchange_user:
                    correo = exchange_user.PrimarySmtpAddress
            except:
                correo = entry.Address
        else:
            correo = getattr(entry, "Address", "")

        if isinstance(correo, str) and "@chileatiende.cl" in correo.lower():
            print(f"{nombre}: {correo}")
            datos.append({"Nombre": nombre, "Correo": correo})

    except Exception as e:
        print(f"[ERROR] Contacto {i+1}: {e}")
        continue

# Guardar resultados en Excel
if datos:
    df = pd.DataFrame(datos)
    ruta = os.path.expandvars(r"%USERPROFILE%\OneDrive - Chileatiende\Escritorio\contactos_gal.xlsx")
    df.to_excel(ruta, index=False)
    print(f"\n[OK] Exportado exitosamente a:\n{ruta}")
else:
    print("\nNo se encontraron correos con @chileatiende.cl en la GAL.")
