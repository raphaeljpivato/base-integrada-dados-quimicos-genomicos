import requests
import random
import pandas as pd
import time

from os.path import exists


SMILES_FILE_DIRECTORY = "C:/Users/Raphael/SMILES"

GET_CID_URL = \
    "https://pubchem.ncbi.nlm.nih.gov/rest/pug/" + \
        "substance/sourceid/DTP.NCI/%s/cids/JSON"
""""
    Response example:
{
  "InformationList": {
    "Information": [
      {
        "SID": 540771,
        "CID": [
          54605111
        ]
      }
    ]
  }
}
"""

GET_SMILES_BY_CID_URL = \
    "https://pubchem.ncbi.nlm.nih.gov/rest/pug/" + \
        "compound/cid/%s/property/CanonicalSMILES/JSON"
"""
    Response example:
{
  "PropertyTable": {
    "Properties": [
      {
        "CID": 54605111,
"CanonicalSMILES": "COC1=CC(=CC2=C1OC(=O)C(=C2)C(=O)O)CN3CCCCC3.Cl"
      }
    ]
  }
}
"""

GET_SMILES_BY_NAME_URL = \
    "https://pubchem.ncbi.nlm.nih.gov/rest/pug/" + \
        "compound/name/%s/property/CanonicalSMILES/JSON"
"""
    Response example:
{
  "PropertyTable": {
    "Properties": [
      {
        "CID": 6505803,
"CanonicalSMILES": "CC1CC(C(C(C=C(C(C(C=CC=C(C(=O)NC2=CC(=O)..."
      },
      {
        "CID": 6440175,
"CanonicalSMILES": "CC1CC(C(C(C=C(C(C(C=CC=C(C(=O)NC2=CC(=O)..."
      },
      {
        "CID": 24870925,
"CanonicalSMILES": "CC1CC(C(C(C=C(C(C(C=CC=C(C(=O)NC2=CC(=O)..."
      }
    ]
  }
}
"""

# Recebe como parametros o ID e o NAME do composto conforme o 
#   banco de dados, bem como o SMILES a ser gravado.
# Formata o nome final do arquivo e checa sua existencia no 
#   diretorio de destino.
# Caso nao exista o nome do arquivo sera o ID recebido com 
#   a extensao SMI (id.smi).
# Caso ja exista um arquivo com esse nome no diretorio serao 
#   tambem concatenados os primeiros 15 chars do NAME + um 
#   identificador unico.
def write_smiles_file(id, name, smiles):
    file_name = SMILES_FILE_DIRECTORY + "/" + id + ".smi"

    if exists(file_name):
       file_name = generate_unique_file_name(id, name)

    with open( file_name, 'w' ) as file:
        file.write(smiles)
    
    file.close()

# Recebe como parametros o ID e o NAME do composto conforme o 
#   banco de dados.
# Inicia um laco que vai de 0 a 20 (pois por uma heuristica deve 
#   cubrir todos os casos de repeticoes de ID).
# Forma um novo nome concatenando o caminho relativo do diretorio 
#   de arquivos SMILES, o ID e ate os primeiros 15 chars do NAME, 
#   acrescidos de um numero identificador de unicidade e a 
#   extensao SMI, resultando em (id_namenamename_n.smi).
# Em ultima instancia, caso a heuristica falhe, retorna um 
#   identificador aleatorio entre 1000 e 1000000.
def generate_unique_file_name(id, name):
    names_first_15_chars = name

    if len(name) > 15:
        names_first_15_chars = name[0:15]

    for i in range(20):
        new_file_name = \
            SMILES_FILE_DIRECTORY + "/" + \
            id + "_" + \
            names_first_15_chars + \
            "_" + \
            str(i) + \
            ".smi"

        if not exists(new_file_name):
            return new_file_name

    return \
        SMILES_FILE_DIRECTORY + "/" + \
        id + "_" + \
        name + "_" + \
        str( random.randint(1000, 1000000) ) + \
        ".smi"


# Excecoes a logica do algoritmo: IDs numericos que nao sao 
#   provenientes da NCI
#   681640 - GDSC:
#https://pubchem.ncbi.nlm.nih.gov/compound/16760707#section=InChI-Key
#   968 - CTRP:
#https://pubchem.ncbi.nlm.nih.gov/compound/3372016

#-----------------------------
#   LEITURA DO EXCEL
#-----------------------------
sheet_name = 'drug_ids_names'
column_names = [ 'A', 'B' ]

spreadsheet = pd.read_excel(
    'Lista_Id_nome_drogas.xlsx', 
    sheet_name = sheet_name, 
    header = None, 
    names = column_names
    ).fillna('')

columns = [ spreadsheet['A'], spreadsheet['B'] ]

#-----------------------------
#   LOOP DO ALGORITMO / LOGICA PARA OBTENCAO DOS SMILES
#-----------------------------
for i in range( 1, len(columns[0]) ):
    smiles = ""
    url = ""
    id = columns[0][i]
    name = columns[1][i]

    try:
        if type(id) is int:
            url = ( GET_CID_URL % id )

            response = requests.get(url)

            if response.status_code == 200:
                response_body = response.json()
                cid = response_body \
                    ['InformationList'] \
                    ['Information'] \
                    [0] \
                    ['CID'] \
                    [0]
                
                url = ( GET_SMILES_BY_CID_URL % cid )

                response = requests.get(url)

                if response.status_code == 200:
                    response_body = response.json()
                    smiles = response_body \
                        ['PropertyTable'] \
                        ['Properties'] \
                        [0] \
                        ['CanonicalSMILES']
        else:
            url = ( GET_SMILES_BY_NAME_URL % id )

            response = requests.get(url)

            if response.status_code == 200:
                smiles = response_body \
                    ['PropertyTable'] \
                    ['Properties'] \
                    [0] \
                    ['CanonicalSMILES']
            else:
                if type(name) is int:
                    parameter = "NSC-" + name
                else:
                    parameter = name
                
                url = ( GET_SMILES_BY_NAME_URL % parameter )

                response = requests.get(url)

                if response.status_code == 200:
                    response_body = response.json()
                    smiles = response_body \
                        ['PropertyTable'] \
                        ['Properties'] \
                        [0] \
                        ['CanonicalSMILES']

        write_smiles_file( str(id), str(name), smiles )
        time.sleep(1)
    except requests.exceptions.Timeout:
        print("Timeout na requisicao:\n")
        print("\tid: %s\n" % id)
        print("\tname: %s\n" % name)
    except Exception:
        print("Excecao na requisicao:\n")
        print("\tid: %s\n" % id)
        print("\tname: %s\n" % name)
