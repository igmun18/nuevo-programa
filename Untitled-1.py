import pandas as pd
from openpyxl import load_workbook
from copy import copy
from dateutil.relativedelta import relativedelta
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

path = "PRUEBA.xlsx"
wb = load_workbook(path)

# Nota: Si no usas el df para procesar datos, podrías omitirlo para ahorrar memoria,
# ya que aquí estamos operando directamente con el objeto 'wb'.
for ws in wb.worksheets:
    ultima_columna = ws.max_column

    header_origen = ws.cell(row=1, column=ultima_columna)
    header_destino = ws.cell(row=1, column=ultima_columna + 1)

    valor_header = header_origen.value
    fecha = None

    if isinstance(valor_header, str):
        try:
            fecha = datetime.strptime(valor_header, "%d/%m/%Y")
        except:
            fecha = datetime.now()
    elif isinstance(valor_header, datetime):
        fecha = valor_header

    if fecha:
        # Usamos 'months' para sumar, 'month' cambiaría la fecha a Enero
        nueva_fecha = fecha + relativedelta(months=1)
        header_destino.value = nueva_fecha
        header_destino.number_format = "DD/MM/YYYY"

    # Copiar formatos básicos
    header_destino.font = copy(header_origen.font)
    header_destino.border = copy(header_origen.border)
    header_destino.fill = copy(header_origen.fill)

    # CORRECCIÓN: alignment (vertical con c) y wrap_text
    header_destino.alignment = Alignment(
        wrap_text=True, 
        horizontal=header_origen.alignment.horizontal, 
        vertical=header_origen.alignment.vertical
    ) 

    for i in range(2, ws.max_row + 1):
        celda_origen = ws.cell(row=i, column=ultima_columna)
        celda_destino = ws.cell(row=i, column=ultima_columna + 1)       

        valor = celda_origen.value

        # CORRECCIÓN: Asignación del valor con '='
        try:
            if isinstance(valor, (int, float)):
                celda_destino.value = valor * 1.15
            else:
                celda_destino.value = valor
        except:
            celda_destino.value = valor
    
        if celda_origen.has_style:
            celda_destino.font = copy(celda_origen.font)
            celda_destino.border = copy(celda_origen.border)
            celda_destino.fill = copy(celda_origen.fill)
            celda_destino.number_format = copy(celda_origen.number_format)

            # Aseguramos el ajuste de texto en cada celda
            nueva_ali = copy(celda_origen.alignment)
            nueva_ali.wrap_text = True
            celda_destino.alignment = nueva_ali
    
    # CORRECCIÓN: column_dimensions (en singular)
    letra_orig = get_column_letter(ultima_columna)
    letra_dest = get_column_letter(ultima_columna + 1)
    ws.column_dimensions[letra_dest].width = ws.column_dimensions[letra_orig].width

wb.save("Nueva.xlsx")