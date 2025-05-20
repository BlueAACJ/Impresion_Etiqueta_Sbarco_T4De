import win32print
from datetime import datetime

def dividir_linea(texto, largo_max=30):
    palabras = texto.split()
    lineas = []
    linea_actual = ""
    for palabra in palabras:
        if len(linea_actual) + len(palabra) + 1 <= largo_max:
            if linea_actual:
                linea_actual += " "
            linea_actual += palabra
        else:
            lineas.append(linea_actual)
            linea_actual = palabra
    if linea_actual:
        lineas.append(linea_actual)
    return lineas

def centrar_texto(texto, ancho_total=609, ancho_char=11):
    ancho_texto = len(texto) * ancho_char
    return (ancho_total - ancho_texto) // 2

def formatear_fecha(fecha_str):
    try:
        fecha = datetime.strptime(fecha_str, "%Y-%m-%d %H:%M:%S")
        return fecha.strftime("%d/%m/%y")
    except:
        return fecha_str

def generar_ticket_1(id_barra, usuario, fecha, codigo_lingote, peso, cod_barra):
    epl = """N
q609
Q1076,24
"""
    epl += 'A30,30,0,4,1,1,N,"Informacion de Pesaje"\n'
    y = 90
    campos = [
        ("ID de Cono/Barra", id_barra),
        ("Usuario", usuario),
        ("Fecha de pesaje", formatear_fecha(fecha)),
        ("Codigo de lingote", codigo_lingote),
        ("Peso", peso),
    ]
    for etiqueta, valor in campos:
        for linea in dividir_linea(f"- {etiqueta}: {valor}"):
            epl += f'A30,{y},0,3,1,1,N,"{linea}"\n'
            y += 40
    epl += f'A30,{y},0,3,1,1,N,"- Codigo de Barra:"\n'
    y += 40
    epl += f'B30,{y},0,1,3,3,80,N,"{cod_barra}"\n'
    y += 90
    epl += f'A100,{y},0,3,1,1,N,"{cod_barra}"\n'
    y += 50
    x_centro = centrar_texto("HEMCO - AZOCAR")
    epl += f'A{x_centro},{y},0,3,1,1,N,"HEMCO - AZOCAR"\n'
    epl += "P1\n"
    return epl

def generar_ticket_2(id_barra, usuario, fecha1, peso1, cod_lingote1, cod_barra1,
                     fecha2, peso2, cod_lingote2, cod_barra2):
    epl = """N
q609
Q1076,24
"""
    epl += 'A30,30,0,4,1,1,N,"Informacion de Pesaje"\n'
    y = 90
    campos = [
        ("ID de Cono/Barra", id_barra),
        ("Usuario", usuario),
        ("Fecha de pesaje", formatear_fecha(fecha1)),
        ("Peso", peso1),
        ("Codigo de lingote", cod_lingote1),
    ]
    for etiqueta, valor in campos:
        for linea in dividir_linea(f"- {etiqueta}: {valor}"):
            epl += f'A30,{y},0,3,1,1,N,"{linea}"\n'
            y += 40
    epl += f'A30,{y},0,3,1,1,N,"- Codigo de Barra:"\n'
    y += 40
    epl += f'B30,{y},0,1,3,3,80,N,"{cod_barra1}"\n'
    y += 90
    epl += f'A100,{y},0,3,1,1,N,"{cod_barra1}"\n'
    y += 50

    # Segunda parte
    campos2 = [
        ("Fecha de pesaje Neto", formatear_fecha(fecha2)),
        ("peso 2", peso2),
        ("Codigo de lingote2", cod_lingote2),
    ]
    for etiqueta, valor in campos2:
        for linea in dividir_linea(f"- {etiqueta}: {valor}"):
            epl += f'A30,{y},0,3,1,1,N,"{linea}"\n'
            y += 40
    epl += f'A30,{y},0,3,1,1,N,"- Codigo de Barra 2:"\n'
    y += 40
    epl += f'B30,{y},0,1,3,3,80,N,"{cod_barra2}"\n'
    y += 90
    epl += f'A100,{y},0,3,1,1,N,"{cod_barra2}"\n'
    y += 50
    x_centro = centrar_texto("HEMCO - AZOCAR")
    epl += f'A{x_centro},{y},0,3,1,1,N,"HEMCO - AZOCAR"\n'
    epl += "P1\n"
    return epl

def generar_ticket_3(id_barra, usuario,
                     fecha3, peso3, cod_lingote3, cod_barra3,
                     cod_seguridad1, cod_seguridad2):
    epl = """N
q609
Q1076,24
"""
    epl += 'A30,30,0,4,1,1,N,"Informacion de Pesaje"\n'
    y = 90
    campos = [
        ("ID de Cono/Barra", id_barra),
        ("Usuario", usuario),
        ("Fecha de pesaje Caja mas cono/barra", formatear_fecha(fecha3)),
        ("peso 3", peso3),
        ("Codigo de lingote3", cod_lingote3),
    ]
    for etiqueta, valor in campos:
        for linea in dividir_linea(f"- {etiqueta}: {valor}"):
            epl += f'A30,{y},0,3,1,1,N,"{linea}"\n'
            y += 40
            
    epl += f'A30,{y},0,3,1,1,N,"- Codigo de Barra 3:"\n'
    y += 40
    epl += f'B30,{y},0,1,3,3,80,N,"{cod_barra3}"\n'
    y += 90
    epl += f'A100,{y},0,3,1,1,N,"{cod_barra3}"\n'
    y += 50

    epl += f'A30,{y},0,3,1,1,N,"- Codigo de seguridad 1:"\n'
    y += 40
    epl += f'B30,{y},0,1,3,3,80,N,"{cod_seguridad1}"\n'
    y += 90
    epl += f'A100,{y},0,3,1,1,N,"{cod_seguridad1}"\n'
    y += 50

    epl += f'A30,{y},0,3,1,1,N,"- Codigo de seguridad 2:"\n'
    y += 40
    epl += f'B30,{y},0,1,3,3,80,N,"{cod_seguridad2}"\n'
    y += 90
    epl += f'A100,{y},0,3,1,1,N,"{cod_seguridad2}"\n'
    y += 50

    x_centro = centrar_texto("HEMCO - AZOCAR")
    epl += f'A{x_centro},{y},0,3,1,1,N,"HEMCO - AZOCAR"\n'
    epl += "P1\n"
    return epl

def imprimir_epl(epl, printer_name):
    raw_data = epl.encode("ascii")
    hprinter = win32print.OpenPrinter(printer_name)
    try:
        hjob = win32print.StartDocPrinter(hprinter, 1, ("Etiqueta", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, raw_data)
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
    finally:
        win32print.ClosePrinter(hprinter)

# ------------------------
# EJEMPLO DE USO DE LOS 3 TICKETS:
# ------------------------

if __name__ == "__main__":
    printer_name = "Sbarco T4De 203 dpi (PEPL) (Copiar 1)"

    # Ticket 1
    epl1 = generar_ticket_1(
        id_barra="11",
        usuario="admin",
        fecha="2025-03-28 15:58:05",
        codigo_lingote="2025032836541111",
        peso="36541 kg",
        cod_barra="2025032836541111"
    )
    imprimir_epl(epl1, printer_name)

    try:
        # Ticket 2
        epl2 = generar_ticket_2(
            id_barra="11",
            usuario="admin",
            fecha1="2025-03-28 15:58:05",
            peso1="36541 kg",
            cod_lingote1="2025032836541111",
            cod_barra1="2025032836541111",
            fecha2="2025-03-28 15:58:05",
            peso2="36541 kg",
            cod_lingote2="20250328365411112",
            cod_barra2="20250328365411112"
        )
        imprimir_epl(epl2, printer_name)
    except Exception as e:
        print("Ocurrió un error al generar o imprimir el ticket 2:", e)


    # Ticket 3 (resumido)
    epl3 = generar_ticket_3(
        id_barra="11",
        usuario="admin",
        fecha3="2025-03-28 15:58:05",
        peso3="36541 kg",
        cod_lingote3="20250328365411112",
        cod_barra3="20250328365411112",
        cod_seguridad1="SEGURIDAD123",
        cod_seguridad2="SEGURIDAD456"
    )
    
    imprimir_epl(epl3, printer_name)
