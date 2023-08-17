import pandas as pd
import xmltodict
import os


def take_file_out (file):
    with open (f'saidas\{file}', 'rb') as xml_file:
        dict_xml = xmltodict.parse(xml_file)
        serie = dict_xml['nfeProc']['NFe']['infNFe']['ide']['serie']
        num = dict_xml['nfeProc']['NFe']['infNFe']['ide']['nNF']
        data = dict_xml['nfeProc']['NFe']['infNFe']['ide']['dhEmi']

        if 'CPF' in dict_xml['nfeProc']['NFe']['infNFe']['dest']:
            cod = dict_xml['nfeProc']['NFe']['infNFe']['dest']['CPF']
        else:
            cod = dict_xml['nfeProc']['NFe']['infNFe']['dest']['CNPJ']

        nome =  dict_xml['nfeProc']['NFe']['infNFe']['dest']['xNome']
        if 'prod' in dict_xml['nfeProc']['NFe']['infNFe']['det']:
            cfop = dict_xml['nfeProc']['NFe']['infNFe']['det']['prod']['CFOP']
        else:
            cfop = dict_xml['nfeProc']['NFe']['infNFe']['det'][1]['prod']['CFOP']
        valor = dict_xml['nfeProc']['NFe']['infNFe']['total']['ICMSTot']['vNF']

        dsvalue_out.append([serie, num, data, cod, nome, cfop, valor])
        

def take_file_in (file):
    with open (f'entradas\{file}', 'rb') as xml_file:
        dict_xml = xmltodict.parse(xml_file)

        num = dict_xml['nfeProc']['NFe']['infNFe']['ide']['nNF']
        data = dict_xml['nfeProc']['NFe']['infNFe']['ide']['dhEmi']

        if 'CPF' in dict_xml['nfeProc']['NFe']['infNFe']['dest']:
            cod = dict_xml['nfeProc']['NFe']['infNFe']['dest']['CPF']
        else:
            cod = dict_xml['nfeProc']['NFe']['infNFe']['dest']['CNPJ']

        nome =  dict_xml['nfeProc']['NFe']['infNFe']['dest']['xNome']
        if 'prod' in dict_xml['nfeProc']['NFe']['infNFe']['det']:
            cfop = dict_xml['nfeProc']['NFe']['infNFe']['det']['prod']['CFOP']
        else:
            cfop = dict_xml['nfeProc']['NFe']['infNFe']['det'][1]['prod']['CFOP']
        valor = dict_xml['nfeProc']['NFe']['infNFe']['total']['ICMSTot']['vNF']

        dsvalue_in.append([ num, data, cod, nome, cfop, valor])


def format_file(file):
    if 'SERIE' in file:
        file['SERIE'] = file['SERIE'].astype(float)
    file['Nº'] = file['Nº'].astype(int)
    file['DATA'] = pd.to_datetime(file['DATA'])
    file['DATA'] = file['DATA'].dt.date
    file['VALOR_TOTAL'] = file['VALOR_TOTAL'].astype(float)
    return file


dscolumns_out = ['SERIE', 'Nº', 'DATA', 'CNPJ/CPF', 'NOME', 'CFOP', 'VALOR_TOTAL']
dscolumns_in = ['Nº', 'DATA', 'CNPJ/CPF', 'NOME', 'CFOP', 'VALOR_TOTAL']
dsvalue_out = []
dsvalue_in = []

list_file_out = os.listdir("saidas")
list_file_in = os.listdir("entradas")

for file in list_file_out:
    take_file_out(file)

for file in list_file_in:
    take_file_in(file)
    
dataset_out = pd.DataFrame(columns=dscolumns_out, data=dsvalue_out)
dataset_in = pd.DataFrame(columns=dscolumns_in, data=dsvalue_in)

dataset_out = format_file(dataset_out)
dataset_in = format_file(dataset_in)

dataset_out.to_excel('NotasFiscais_Saidas.xlsx', index=False, sheet_name='saidas')
dataset_in.to_excel('NotasFiscais_Entradas.xlsx', index=False, sheet_name='entradas')
