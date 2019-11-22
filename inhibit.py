# -*- coding: latin-1 -*-
import csv
from openpyxl import Workbook
from openpyxl.drawing.image import Image
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from datetime import datetime
from inhibit_lib.mylib import *

wb = Workbook()
ws = wb.active
ws.append(['DATA', 'HORA', 'AREA', 'TAG', 'BYPASS'])        # Nome das colunas
file_csv = 'Inhibit.csv'                                    # ex.: 'C:/Users/as0x/Desktop/Inhibit.csv'
is_file(file_csv)                                           # verifica se existe o arquivo

for row in csv.reader(open(file_csv)):
    ws.append([cell.replace('/HMI', '') for cell in row])   # limpa o TAG para ficar mais apresentável


ws.delete_rows(2)  # Deleta índice de coluna original
ws.insert_rows(1)  # Insere linha do título

"""         FORMATAÇÃO  DO  ÁREA  DOS  DADOS         """

area = "A3:" + ws.dimensions.split(':')[1]
font = Font(name='Times New Roman', sz=13)
alignm = Alignment(indent=1)
thin_border = Border(left=Side(style='thin', color='FF66AACC'), right=Side(style='thin', color='FF66AACC'),
                     top=Side(style='thin', color='FF66AACC'), bottom=Side(style='thin', color='FF66AACC'))
range_format(ws, cell_range=area, font=font, border=thin_border, alignment=alignm)
col_width(ws, factor=1.35)  # Largura das colunas (fator de acordo com o tamanho da fonte)

"""         FORMATAÇÃO  DO  ÍNDICE  DAS  COLUNAS         """

font = Font(name='Arial', sz=12, color='FFFFFFFF', b=True)
fill = PatternFill(patternType='solid', fgColor="FF6699BB")
range_format(ws, cell_range='A2:E2', font=font, fill=fill, border=thin_border, alignment=alignm)

"""         FORMATAÇÃO  DA  ÁREA  DO  TÍTULO           """

ws.merge_cells('A1:E1')
ws['A1'].value = ' LISTA  DE  INIBIÇÕES  ICSS  -  P55'
font = Font(name='Arial', sz=16, color="00333399", b=True)
alignm = Alignment(horizontal='center', vertical='center')
fill = PatternFill(patternType='solid', fgColor="FFE0E0E0")
range_format(ws, cell_range='A1:E1', font=font, alignment=alignm, fill=fill)

ws.row_dimensions[1].height = 27                        # Altura da linha do título

img = Image('logo.png')                                 # Gerar imagem logo BR
ws.add_image(img, 'A1')

"""         FINALIZAÇÃO           """

now = datetime.now()                                    # Levanta as infomraões de data e hora
print_settings(ws)                                      # Ajuste da impressão
view_settings(ws, now)                                  # Ajuste da visualização

outfile = "Inhibit_"                                    # ex.: 'C:/Users/as0x/Desktop/Inhibit_'
outfile += now.strftime("%Y%m%d%H%M") + ".xlsx"         # Arquivo com data e hora no nome

write_file(wb, outfile)                                 # Grava o arquivo excel xlsx
