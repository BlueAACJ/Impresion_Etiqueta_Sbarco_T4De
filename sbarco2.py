import win32print

# Datos
id_barra = "2"
usuario = "MINEROS DE NICARAGUA"
fecha = "30/07/2003"
codigo = "HEMCO"
peso = "25 Kg"
codigos_barra = ["12345678975", "98765432109", "11223344556", "99887766554"]

# Divide el texto largo sin cortar palabras
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

# Inicia EPL
epl = """
N
q609
Q1076,24
"""

# Título
epl += f'A30,30,0,4,1,1,N,"Informacion de Registro"\n'
y_actual = 90

# Campos en orden
campos = {
    "ID cono/barra": id_barra,
    "Usuario": usuario,
    "Fecha": fecha,
    "Codigo": codigo,
    "Peso": peso,
}

for etiqueta, valor in campos.items():
    lineas = dividir_linea(f"- {etiqueta}: {valor}")
    for linea in lineas:
        epl += f'A30,{y_actual},0,3,1,1,N,"{linea}"\n'
        y_actual += 40

# Encabezado códigos de barra
epl += f'A30,{y_actual},0,3,1,1,N,"- Codigos de barra:"\n'
y_actual += 40

# Códigos de barra
for cod in codigos_barra:
    epl += f'B30,{y_actual},0,1,3,3,80,N,"{cod}"\n'
    y_actual += 90
    epl += f'A100,{y_actual},0,3,1,1,N,"{cod}"\n'
    y_actual += 50

# Texto final centrado
texto_final = "HEMCO - AZOCAR"
ancho_caracter = 11
ancho_texto = len(texto_final) * ancho_caracter
x_centro = (609 - ancho_texto) // 2
epl += f'A{x_centro},{y_actual},0,3,1,1,N,"{texto_final}"\n'

epl += "P1\n"

# Envío a la impresora
raw_data = epl.encode("ascii")
printer_name = "Sbarco T4De 203 dpi (PEPL) (Copiar 1)"

hprinter = win32print.OpenPrinter(printer_name)
try:
    hjob = win32print.StartDocPrinter(hprinter, 1, ("Etiqueta", None, "RAW"))
    win32print.StartPagePrinter(hprinter)
    win32print.WritePrinter(hprinter, raw_data)
    win32print.EndPagePrinter(hprinter)
    win32print.EndDocPrinter(hprinter)
finally:
    win32print.ClosePrinter(hprinter)
