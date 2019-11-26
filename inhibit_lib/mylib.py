# coding: latin1
import os.path
import sys


def is_file(file_csv):
    """ M�dulo para checar se arquivo CSV existe no caminho indicado"""
    if os.path.exists(file_csv):
        print(f'\nArquivo CSV: {file_csv}\n')
    else:
        print(
            f'\nOops, algo deu ruim!\nArquivo CSV não encontrado: {file_csv}')
        sys.exit()


def col_width(ws, factor=1.1):
    """ Define e formata a largura das colunas, encontrando a maior c�lula da coluna"""
    for col in ws.columns:
        max_len = 0
        for cell in col:
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col[0].column_letter].width = (
            max_len + 3) * factor


def range_format(ws, cell_range, font=None, alignment=None, border=None, fill=None):
    """ Formata range de c�lulas
        Os tipos de formata��es s�o opcionais, podendos escolher todos ou apenas os desej�veis
     """
    if ':' not in cell_range:
        cell_range += ':' + cell_range

    for row in ws[cell_range]:
        for cell in row:
            if font:
                cell.font = font
            if alignment:
                cell.alignment = alignment
            if border:
                cell.border = border
            if fill:
                cell.fill = fill


def print_settings(ws):
    ws.print_area = ws.dimensions  # configura��es de visualiza��o e impress�o
    ws.title = 'Inhibit List'  # Planilha
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE  # Paisagem
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.footer = 0.2
    ws.print_options.horizontalCentered = True  # Centraliza �rea de impress�o
    # ws.print_options.verticalCentered = True
    ws.page_setup.fitToWidth = True  # Ajusta � Largura


def view_settings(ws, now):
    ws.freeze_panes = "A3"  # Congela pain�is
    ws.sheet_view.showGridLines = False  # Linhas de grade
    ws.print_title_rows = '1:2'  # Linhas como cabe�alho
    # Rodap� �direita: P�gina
    ws.oddFooter.right.text = "P�gina &[Page] de &N"
    # Rodap� �esquerda: data e hora
    ws.oddFooter.left.text = f'Impress�o de {now.strftime("%d-%m-%Y %H:%M")}'


def write_file(wb, outfile, execute=False):
    try:
        wb.save(outfile)  # Grava o arquivo excel
    except PermissionError:
        print(f'Oops, arquivo {outfile} n�o foi gerado.')
        print('Feche o arquivo excel e tente novamente.')
    except Exception:
        print(f'Oops, deu ruim:')
        raise Exception
    else:
        if execute:
            try:
                print(f'Abrindo arquivo Excel:\n{outfile}')
                os.system(outfile)
            except Exception:
                print(f'N�o foi poss�vel abrir o arquivo: {outfile}')
        else:
            print(f'Arquivo Excel gerado com sucesso:\n\t{outfile}')
