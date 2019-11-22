import os.path
import sys


def is_file(file_csv):
    """ Módulo para checar se arquivo CSV existe no caminho indicado"""
    if os.path.exists(file_csv):
        print(f'\nArquivo CSV: {file_csv}\n')
    else:
        print(f'\nOops, algo deu ruim!\nArquivo CSV não encontrado: {file_csv}')
        sys.exit()


def col_width(ws, factor=1.1):
    """ Define e formata a largura das colunas, encontrando a maior célula da coluna"""
    for col in ws.columns:
        max_len = 0
        for cell in col:
            max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col[0].column_letter].width = (max_len + 3) * factor


def range_format(ws, cell_range, font=None, alignment=None, border=None, fill=None):
    """ Formata range de células
        Os tipos de formatação são opcionais, podendos escolher todos ou apenas os desejáveis
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
    ws.print_area = ws.dimensions  # configurações de visualização e impressão
    ws.title = 'Inhibit List'  # Planilha
    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE  # Paisagem
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.left = 0.5
    ws.page_margins.right = 0.5
    ws.page_margins.footer = 0.2
    ws.print_options.horizontalCentered = True  # Centraliza área de impressão
    # ws.print_options.verticalCentered = True
    ws.page_setup.fitToWidth = True  # Ajusta à Largura


def view_settings(ws, now):
    ws.freeze_panes = "A3"  # Congela painéis
    ws.sheet_view.showGridLines = False  # Linhas de grade
    ws.print_title_rows = '1:2'  # Linhas como cabeçalho
    ws.oddFooter.right.text = "Página &[Page] de &N"  # Rodapé à direita: Página
    ws.oddFooter.left.text = f'Impressão de {now.strftime("%d-%m-%Y %H:%M")}'  # Rodapé à esquerda data e hora


def write_file(wb, outfile):
    try:
        wb.save(outfile)  # Grava o arquivo excel
    except PermissionError:
        print(f'Oops, arquivo {outfile} não foi gerado.')
        print('Feche o arquivo excel e tente novamente.')
    except Exception:
        print(f'Oops, deu ruim:')
        raise Exception
    else:
        print(f'Arquivo Excel gerado com sucesso:\n\t{outfile}')
