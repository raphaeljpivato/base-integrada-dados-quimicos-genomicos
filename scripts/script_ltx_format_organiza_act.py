import xlsxwriter
import pandas as pd


def generate_names_array():
    letters = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
               'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 
               'U', 'V', 'W', 'X', 'Y', 'Z' ]
    response_array = []

    # Adiciona letras unitarias (A:Z)
    for i in range( len(letters) ):
        response_array.append(letters[i])

    # Adiciona letras em pares (AA:AZ, BA:BZ, ..., ZA:ZZ)
    for first_letter in range( len(letters) ):
        for last_letter in range( len(letters) ):
            response_array.append(
                letters[first_letter] + letters[last_letter] 
                )

    # Adiciona letras em trios (AAA:AAZ, ABA:ABZ, ..., CAA:CZZ)
    for first_letter in range(3): 
        for middle_letter in range( len(letters) ):
            for last_letter in range( len(letters) ):
                response_array.append( 
                    letters[first_letter] + 
                    letters[middle_letter] + 
                    letters[last_letter] 
                    )

    # Adiciona letras em trios de DAA:DOZ
    for middle_letter in range(15):
        for last_letter in range( len(letters) ):
            response_array.append( 
                letters[3] + 
                letters[middle_letter] + 
                letters[last_letter] 
                )

    # Adiciona letras em trios de DPA:DPM
    for last_letter in range(13):
        response_array.append( 
            letters[3] + 
            letters[15] + 
            letters[last_letter] 
            )

    return response_array


def write_header(spreadsheet, worksheet, indexes):
    # Instancia as colunas da planilha parametrizada
    columns = []

    for index in range( len(indexes) ):
        columns.append( spreadsheet[indexes[index]] )

    # Grava os cabecalhos (A1:D30449)
    for column in range(4):
        for row in range( len(columns[column]) ):
            worksheet.write( row, column, columns[column][row] )


def write_body(spreadsheet, worksheet, indexes):
    # Instancia as colunas da planilha parametrizada
    columns = []

    for index in range( len(indexes) ):
        columns.append( spreadsheet[indexes[index]] )

    # Primeiro formata as strings das linhagens celulares (primeira
    #   linha, colunas de E1:DPM1), removendo espacos, tracos,
    #   barras, ponto e virgulas.
    # Indice inicia em 4 = coluna E
    # Adicionalmente padroniza as strings para upper case
    for i in range( 4, len(columns) ):
        columns[i][0] = str(columns[i][0]).replace(' ', '')
        columns[i][0] = str(columns[i][0]).replace('-', '')
        columns[i][0] = str(columns[i][0]).replace('/', '')
        columns[i][0] = str(columns[i][0]).replace(';', '')
        columns[i][0] = str(columns[i][0]).upper()

    # Nomes das linhagenes celulares (atual e anterior para 
    #   identificar repeticoes).
    # column_write_index identifica o indice da coluna em que a 
    #   escrita deve ocorrer (inicia em tres dado que a 
    #   atualizacao de seu valor ++ ocorre no inicio do loop).
    previous_cell_line = ''
    current_cell_line = ''
    column_write_index = 3
    for column in range( 4, len(columns) ):
        current_cell_line = columns[column][0]  # Atualiza 
                                                # a linhagem celular

        # Caso atual seja diferente da anterior, trata uma 
        #   linhagem celular utilizada apenas por um dos conjuntos
        #   de dados (isolada); caso sejam iguais, trata a 
        #   repeticao para gravar os registros em uma unica linha 
        #   na planilha de saida.
        if current_cell_line != previous_cell_line:
            column_write_index += 1  # Com a quebra o indice de 
                                     # escrita deve ser incrementado

            # Grava de E1:DPMrow
            for i in range(len(columns[0])): # colunas de mesmo 
                                             # comprimento, utilizando
                                             # a A como referencia
                if columns[column][i] != '' \
                    and columns[column][i] is not None:
                    worksheet.write( 
                        i, 
                        column_write_index, 
                        columns[column][i] 
                        )
        else:
            # Grava de E2:DPMrow
            for i in range( 1, len(columns[0]) ): # colunas de mesmo 
                                                  # comprimento, 
                                                  # utilizando a A 
                                                  # como referencia
                if columns[column][i] != '' \
                    and columns[column][i] is not None:
                    worksheet.write( 
                        i, 
                        column_write_index, 
                        columns[column][i] 
                        )
            
        previous_cell_line = current_cell_line


sheet1 = 'act'
names1 = generate_names_array()

spreadsheet1 = pd.read_excel(
    'Dados_genomicos_e_atividade_biologica_ordenados.xlsx', 
    sheet_name = sheet1, 
    header = None, 
    names = names1
    ).fillna('')

workbook = xlsxwriter.Workbook(
    'Dados_genomicos_e_atividade_biologica_organizados.xlsx', 
    {'nan_inf_to_errors': True})

worksheet1 = workbook.add_worksheet(sheet1)

write_header(spreadsheet1, worksheet1, names1)

write_body(spreadsheet1, worksheet1, names1)

workbook.close()
