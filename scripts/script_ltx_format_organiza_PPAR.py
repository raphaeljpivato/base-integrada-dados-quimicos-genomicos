import xlsxwriter
import pandas as pd

def write_header(spreadsheet, worksheet, indexes):
    # Instancia as colunas da planilha parametrizada
    columns = []

    for index in range( len(indexes) ):
        columns.append( spreadsheet[indexes[index]] )

    # Grava os cabecalhos superiores (A1:I4)
    for row in range(4):
        for column in range(9):
            worksheet.write( row, column, columns[column][row] )

    # Grava os cabecalhos restantes em linha (J4:V4)
    row = 3  # Todos os seguintes na quarta linha
    for column in range(9, 22):
        worksheet.write( row, column, columns[column][row] )


def write_body(spreadsheet, worksheet, indexes):
    # Instancia as colunas da planilha parametrizada
    columns = []

    for index in range( len(indexes) ):
        columns.append( spreadsheet[indexes[index]] )

    # Primeiro formata as strings das linhagens celulares 
    #   (coluna A = indice 0), removendo separadores.
    # O indice inicia em 4 dado que e a partir da quinta linha
    #   da planilha que os dados de detalhe se encontram.
    # Adicionalmente padroniza as strings para upper case
    for i in range( 4, len(columns[0]) ):
        columns[i][0] = str(columns[i][0]).replace(' ', '')
        columns[i][0] = str(columns[i][0]).replace('-', '')
        columns[i][0] = str(columns[i][0]).replace('/', '')
        columns[i][0] = str(columns[i][0]).replace(';', '')
        columns[i][0] = str(columns[i][0]).upper()

    # Nomes das linhagenes celulares (atual e anterior para 
    #   identificar repeticoes).
    # row_write_index identifica o indice da linha em que a 
    #   escrita deve ocorrer (inicia em tres dado que a 
    #   atualizacao de seu valor ++ ocorre no inicio do loop).
    previous_cell_line = ''
    current_cell_line = ''
    row_write_index = 3
    for row in range( 4, len(columns[0]) ):
        current_cell_line = columns[0][row]  # Atualiza a 
                                             # linhagem celular

        # Caso atual seja diferente da anterior, trata uma 
        #   linhagem celular utilizada apenas por um dos conjuntos
        #   de dados (isolada); caso sejam iguais, trata a 
        #   repeticao para gravar os registros em uma unica
        #   linha na planilha de saida.
        if current_cell_line != previous_cell_line:
            row_write_index += 1  # Com a quebra o indice de 
                                  # escrita deve ser incrementado

            # Grava de Arow:Vrow
            for i in range(len(columns)):
                if columns[i][row] != '' \
                    and columns[i][row] is not None:
                    worksheet.write(
                        row_write_index, i, columns[i][row]
                    )
        else:
            # Grava de Brow:Vrow
            for i in range( 1, len(columns) ):
                if columns[i][row] != '' \
                    and columns[i][row] is not None: 
                    worksheet.write(
                        row_write_index, i, columns[i][row]
                    )
            
        previous_cell_line = current_cell_line


sheet1 = 'PPARa'
sheet2 = 'PPARd'
sheet3 = 'PPARgC1A'
sheet4 = 'PPARgC1B'
sheet5 = 'PPARg'

columnNames = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
                'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 
                'U', 'V' ]

spreadsheet1 = pd.read_excel(
    'Dados_genomicos_e_atividade_biologica_ordenados.xlsx', 
    sheet_name = sheet1, 
    header = None, 
    names = columnNames
    ).fillna('')
spreadsheet2 = pd.read_excel(
    'Dados_genomicos_e_atividade_biologica_ordenados.xlsx',
    sheet_name = sheet2,
    header = None,
    names = columnNames
    ).fillna('')
spreadsheet3 = pd.read_excel(
    'Dados_genomicos_e_atividade_biologica_ordenados.xlsx', 
    sheet_name = sheet3, 
    header = None, 
    names = columnNames 
    ).fillna('')
spreadsheet4 = pd.read_excel(
    'Dados_genomicos_e_atividade_biologica_ordenados.xlsx', 
    sheet_name = sheet4, 
    header = None, 
    names = columnNames 
    ).fillna('')
spreadsheet5 = pd.read_excel(
    'Dados_genomicos_e_atividade_biologica_ordenados.xlsx', 
    sheet_name = sheet5, 
    header = None, 
    names = columnNames 
    ).fillna('')

workbook = xlsxwriter.Workbook( 
    'Dados_genomicos_e_atividade_biologica_organizados.xlsx', 
    {'nan_inf_to_errors': True} 
    )

worksheet1 = workbook.add_worksheet(sheet1)
worksheet2 = workbook.add_worksheet(sheet2)
worksheet3 = workbook.add_worksheet(sheet3)
worksheet4 = workbook.add_worksheet(sheet4)
worksheet5 = workbook.add_worksheet(sheet5)

write_header(spreadsheet1, worksheet1, columnNames)
write_header(spreadsheet2, worksheet2, columnNames)
write_header(spreadsheet3, worksheet3, columnNames)
write_header(spreadsheet4, worksheet4, columnNames)
write_header(spreadsheet5, worksheet5, columnNames)

write_body(spreadsheet1, worksheet1, columnNames)
write_body(spreadsheet2, worksheet2, columnNames)
write_body(spreadsheet3, worksheet3, columnNames)
write_body(spreadsheet4, worksheet4, columnNames)
write_body(spreadsheet5, worksheet5, columnNames)

workbook.close()
