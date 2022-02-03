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
                letters[first_letter] +
                letters[last_letter] 
                )

    # Adiciona letras em trios (AAA:AAZ, ABA:ABZ, ..., AZA:AZZ)
    for first_letter in range(1): 
        for middle_letter in range( len(letters) ):
            for last_letter in range( len(letters) ):
                response_array.append( 
                    letters[first_letter] + 
                    letters[middle_letter] + 
                    letters[last_letter] 
                    )

    # Adiciona letras em trios de BAA:BGZ
    for middle_letter in range(7):
        for last_letter in range( len(letters) ):
            response_array.append( 
                letters[1] + 
                letters[middle_letter] + 
                letters[last_letter] 
                )

    # Adiciona letras em trios de BHA:BHE
    for last_letter in range(5):
        response_array.append( 
            letters[1] + 
            letters[7] + 
            letters[last_letter] 
            )

    return response_array


def write_header(spreadsheet, worksheet, indexes):
    # Instancia as colunas da planilha parametrizada
    columns = []

    for index in range( len(indexes) ):
        columns.append( spreadsheet[indexes[index]] )

    # Grava os cabecalhos superiores (A1:D30449)
    for row in range(30449):
        for column in range(4):
            worksheet.write( row, column, columns[column][row] )

    # Grava os cabecalhos em linha (E1:BHE1)
    row = 0  # Todos os seguintes na primeira linha
    for column in range(4, 1565):
        worksheet.write( row, column, columns[column][row] )

    # Grava as contagens e medias (BHD2:BHE30449)
    for row in range(1, 30449):
        for column in range(1563, 1565):
            worksheet.write( row, column, columns[column][row] )


def write_body(spreadsheet, worksheet, indexes):
    # Instancia as colunas da planilha parametrizada
    columns = []

    for index in range( len(indexes) ):
        columns.append( spreadsheet[indexes[index]] )

    # Grava dados de detalhe no intervalo E2:BHB30449 obedecendo 
    #   o seguinte criterio:
    #   Caso haja dados na celula estes sao mantidos
    #   Caso contrario, preenche-se as celulas vazias com a 
    #       media da linha - coluna 1564 (BHE)
    for column in range( 4, len(columns) - 3 ): # De E a BHB
        for row in range( 1, len(columns[column])): # De 2 a 30449
            if columns[column][row] != '' \
                and columns[column][row] is not None: 
                    worksheet.write( 
                        row, 
                        column, 
                        columns[column][row] 
                        )
            else:
                worksheet.write( 
                    row, 
                    column, 
                    columns[1564][row] 
                    )


sheet1 = 'act'
names1 = generate_names_array()

spreadsheet1 = pd.read_excel('Dados_genomicos_e_atividade_biologica_tratados.xlsx', 
                             sheet_name = sheet1, 
                             header = None, 
                             names = names1
                             ).fillna('')

workbook = xlsxwriter.Workbook('Dados_genomicos_e_atividade_biologica_final.xlsx', 
                              {'nan_inf_to_errors': True})
workbook.use_zip64()

worksheet1 = workbook.add_worksheet(sheet1)

write_header(spreadsheet1, worksheet1, names1)

write_body(spreadsheet1, worksheet1, names1)

workbook.close()
