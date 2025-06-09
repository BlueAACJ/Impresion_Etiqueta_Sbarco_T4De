import win32print

nombre_impresora = "Sbarco T4De 203 dpi (PEPL)"

comandos = """
N
q609
Q1076,24

A100,30,0,3,1,1,N,"Comparación de grosor"

A100,100,0,3,1,1,N,"Grosor 2,3"
B130,130,0,1,2,3,60,N,"100.01"

A100,220,0,3,1,1,N,"Grosor 3,4"
B130,250,0,1,3,4,70,N,"100.01"

A100,340,0,3,1,1,N,"Grosor 4,5"
B130,370,0,1,4,5,80,N,"172"

A100,470,0,3,1,1,N,"Grosor 5,6"
B130,500,0,1,5,6,90,N,"100.01"

P1
"""

handle = win32print.OpenPrinter(nombre_impresora)
try:
    job_id = win32print.StartDocPrinter(handle, 1, ("TestGrosorBarras", None, "RAW"))
    win32print.StartPagePrinter(handle)
    win32print.WritePrinter(handle, comandos.encode('cp1252'))
    win32print.EndPagePrinter(handle)
    win32print.EndDocPrinter(handle)
    print("Prueba de grosor de código de barras enviada correctamente.")
finally:
    win32print.ClosePrinter(handle)
