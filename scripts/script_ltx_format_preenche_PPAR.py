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

    # Grava dados estaticos (coluna A + row 1683) - 
    #   rotulos das linhagens celulares e as medias de cada coluna 
    #   antes do preenchimento das lacunas vazias.
    # coluna A = columns[0]
    # row 1683 = columns[i][1682]
    # Iteracao iniciando na 4 linha dado que o cabecalho fora 
    #   previamente gravado.
    for row in range( 4, len(columns[0]) ):
        worksheet.write( row, 0, columns[0][row] )

    for column in range( 1, len(columns) ):
        worksheet.write( 1682, column, columns[column][1682] )

    # Tratamento dinamico dos DNA Copy Number (cop) - 
    #   colunas de B a E (1-4).
    # Laco itera por cada linha partindo do indice 4 (quinta linha)
    #   ate o total de registros da coluna B, desprezando as 
    #   ultimas tres linhas com as medias (utilizando a coluna B
    #   de exemplo, mas sao todas de mesmo tamanho).
    # Para cada linha e verificada a quantidade de registros vazios:
    #   4: linha vazia, preencher todas as colunas com as medias 
    #       das colunas calculadas na linha 1683;
    #   3: faltam os demais registros, copia-se o existente para 
    #       os vazios;
    #   2: dois registros presentes, copia-se a media dos dois 
    #       para os vazios;
    #   1: tres registros presentes, copia-se a media dos tres 
    #       para o vazio;
    #   0: linha cheia, copiar item a item para a planilha de saida.
    for row in range( 4, len(columns[1]) - 3 ):
        data_line = [ 
            columns[1][row], 
            columns[2][row], 
            columns[3][row], 
            columns[4][row] 
            ] 
        empty_count = data_line.count('')

        if empty_count == 4:
            worksheet.write( row, 1, columns[1][1682] )
            worksheet.write( row, 2, columns[2][1682] )
            worksheet.write( row, 3, columns[3][1682] )
            worksheet.write( row, 4, columns[4][1682] )

        elif empty_count == 3:
            filled_data = 0

            for data in data_line:
                if data != '':
                    filled_data = data

            worksheet.write( row, 1, filled_data )
            worksheet.write( row, 2, filled_data )
            worksheet.write( row, 3, filled_data )
            worksheet.write( row, 4, filled_data )

        elif empty_count == 2:
            average = 0

            for data in data_line:
                if data != '':
                    average += data

            average /= 2

            for column in range( len(data_line) ):
                if data_line[column] != '':
                    worksheet.write( 
                        row, 
                        column + 1, 
                        data_line[column] 
                        )
                else:
                    worksheet.write( row, column + 1, average )

        elif empty_count == 1:
            average = 0

            for data in data_line:
                if data != '':
                    average += data

            average /= 3

            for column in range( len(data_line) ):
                if data_line[column] != '':
                    worksheet.write( 
                        row, 
                        column + 1, 
                        data_line[column] 
                        )
                else:
                    worksheet.write( row, column + 1, average )

        else:
            worksheet.write( row, 1, columns[1][row] )
            worksheet.write( row, 2, columns[2][row] )
            worksheet.write( row, 3, columns[3][row] )
            worksheet.write( row, 4, columns[4][row] )

    # Tratamento dinamico do Crispr (cri) - coluna F (5)
    # Laco itera por cada linha partindo do indice 4 (quinta linha)
    #   ate o total de registros da coluna F, desprezando as 
    #   ultimas tres linhas com as medias. 
    # Para cada linha e verificada se esta preenchida ou vazia
    #   Se vazia, preencher com a media da coluna calculada na 
    #       linha 1683;
    #   Se preenchida, copiar o item para a planilha de saida
    for row in range( 4, len(columns[5]) - 3 ):
        if columns[5][row] == '':
            worksheet.write( row, 5, columns[5][1682] )

        else:
            worksheet.write( row, 5, columns[5][row] )

    # Tratamento dinamico das Microarray RNA Expression (exp) - 
    #   colunas de G a J (6-9).
    # Laco itera por cada linha partindo do indice 4 (quinta linha) 
    #   ate o total de registros da coluna G, desprezando as 
    #   ultimas tres linhas com as medias.
    # Para cada linha e verificada a quantidade de registros vazios:
    #   4: linha vazia, preencher todas as colunas com as medias 
    #       das colunas calculadas na linha 1683;
    #   3: faltam os demais registros, copia-se o existente para 
    #       os vazios;
    #   2: dois registros presentes, copia-se a media dos dois 
    #       para os vazios;
    #   1: tres registros presentes, copia-se a media dos tres 
    #       para os vazios;
    #   0: linha cheia, copiar item a item para a planilha de saida.
    for row in range( 4, len(columns[6]) - 3 ):
        data_line = [ 
            columns[6][row], 
            columns[7][row], 
            columns[8][row], 
            columns[9][row] 
            ] 
        empty_count = data_line.count('')

        if empty_count == 4:
            worksheet.write( row, 6, columns[6][1682] )
            worksheet.write( row, 7, columns[7][1682] )
            worksheet.write( row, 8, columns[8][1682] )
            worksheet.write( row, 9, columns[9][1682] )

        elif empty_count == 3:
            filled_data = 0

            for data in data_line:
                if data != '':
                    filled_data = data

            worksheet.write( row, 6, filled_data )
            worksheet.write( row, 7, filled_data )
            worksheet.write( row, 8, filled_data )
            worksheet.write( row, 9, filled_data )

        elif empty_count == 2:
            average = 0

            for data in data_line:
                if data != '':
                    average += data

            average /= 2

            for column in range( len(data_line) ):
                if data_line[column] != '':
                    worksheet.write( 
                        row, 
                        column + 6, 
                        data_line[column] 
                        )
                else:
                    worksheet.write( row, column + 6, average )

        elif empty_count == 1:
            average = 0

            for data in data_line:
                if data != '':
                    average += data

            average /= 3

            for column in range( len(data_line) ):
                if data_line[column] != '':
                    worksheet.write( 
                        row, 
                        column + 6, 
                        data_line[column] 
                        )
                else:
                    worksheet.write( row, column + 6, average )

        else:
            worksheet.write( row, 6, columns[6][row] )
            worksheet.write( row, 7, columns[7][row] )
            worksheet.write( row, 8, columns[8][row] )
            worksheet.write( row, 9, columns[9][row] )

    # Tratamento dinamico das DNA Methylation 450K (met) - 
    #   colunas K e L (10-11).
    # Laco itera por cada linha partindo do indice 4 (quinta linha) 
    #   ate o total de registros da coluna K, desprezando as 
    #   ultimas tres linhas com as medias.
    # Para cada linha e verificada a quantidade de registros vazios:
    #   2: linha vazia, preencher todas as colunas com as medias das 
    #       colunas calculadas na linha 1683;
    #   1: copia-se o existente para o vazio;
    #   0: linha cheia, copiar item a item para a planilha de saida.
    for row in range( 4, len(columns[10]) - 3 ):
        data_line = [ columns[10][row], columns[11][row] ]
        empty_count = data_line.count('')

        if empty_count == 2:
            worksheet.write( row, 10, columns[10][1682] )
            worksheet.write( row, 11, columns[11][1682] )

        elif empty_count == 1:
            filled_data = 0

            for data in data_line:
                if data != '':
                    filled_data = data

            worksheet.write( row, 10, filled_data )
            worksheet.write( row, 11, filled_data )

        else:
            worksheet.write( row, 10, columns[10][row] )
            worksheet.write( row, 11, columns[11][row] )

    # Tratamento dinamico do DNA Methylation RRBS (rrb) - 
    #   coluna M (12).
    # Laco itera por cada linha partindo do indice 4 (quinta linha) 
    #   ate o total de registros da coluna M, desprezando as 
    #   ultimas tres linhas com as medias.
    # Para cada linha e verificada se esta preenchida ou vazia:
    #   Se vazia, preencher com a media da coluna calculada 
    #       na linha 1683;
    #   Se preenchida, copiar o item para a planilha de saida.
    for row in range( 4, len(columns[12]) - 3 ):
        if columns[12][row] == '':
            worksheet.write( row, 12, columns[12][1682] )

        else:
            worksheet.write( row, 12, columns[12][row] )

    # Tratamento dinamico das RNA-seq Expression using log2.FPKM+1 
    #   (xsq) - colunas N e O (13-14).
    # Laco itera por cada linha partindo do indice 4 (quinta linha) 
    #   ate o total de registros da coluna N, desprezando as 
    #   ultimas tres linhas com as medias.
    # Para cada linha e verificada a quantidade de registros vazios:
    #   2: linha vazia, preencher todas as colunas com as medias das
    #       colunas calculadas na linha 1683;
    #   1: copia-se o existente para o vazio;
    #   0: linha cheia, copiar item a item para a planilha de saida.
    for row in range( 4, len(columns[13]) - 3 ):
        data_line = [ columns[13][row], columns[14][row] ]
        empty_count = data_line.count('')

        if empty_count == 2:
            worksheet.write( row, 13, columns[13][1682] )
            worksheet.write( row, 14, columns[14][1682] )

        elif empty_count == 1:
            filled_data = 0

            for data in data_line:
                if data != '':
                    filled_data = data

            worksheet.write( row, 13, filled_data )
            worksheet.write( row, 14, filled_data )

        else:
            worksheet.write( row, 13, columns[13][row] )
            worksheet.write( row, 14, columns[14][row] )


sheet1 = 'PPARa'
sheet2 = 'PPARd'
sheet3 = 'PPARgC1A'
sheet4 = 'PPARgC1B'
sheet5 = 'PPARg'

columnNames = [ 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 
                'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 
                'U', 'V' ]

spreadsheet1 = pd.read_excel( 
    'Dados_genomicos_e_atividade_biologica_tratados.xlsx', 
    sheet_name = sheet1, 
    header = None, 
    names = columnNames
    ).fillna('')
spreadsheet2 = pd.read_excel( 
    'Dados_genomicos_e_atividade_biologica_tratados.xlsx', 
    sheet_name = sheet2, 
    header = None, 
    names = columnNames 
    ).fillna('')
spreadsheet3 = pd.read_excel( 
    'Dados_genomicos_e_atividade_biologica_tratados.xlsx', 
    sheet_name = sheet3, 
    header = None, 
    names = columnNames 
    ).fillna('')
spreadsheet4 = pd.read_excel( 
    'Dados_genomicos_e_atividade_biologica_tratados.xlsx', 
    sheet_name = sheet4, 
    header = None, 
    names = columnNames 
    ).fillna('')
spreadsheet5 = pd.read_excel( 
    'Dados_genomicos_e_atividade_biologica_tratados.xlsx', 
    sheet_name = sheet5, 
    header = None, 
    names = columnNames 
    ).fillna('')

workbook = xlsxwriter.Workbook( 
    'Dados_genomicos_e_atividade_biologica_final.xlsx', 
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
