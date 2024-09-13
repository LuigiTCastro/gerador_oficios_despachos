import os
import pandas as pd
from dotenv import load_dotenv



def get_order_worksheet():
    load_dotenv()
    FILE_PATH = os.getenv('FILE_PATH')
    SHEET_NAME = os.getenv('SHEET_NAME')
    order_worksheet = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, skiprows=5)
    filtered_worksheet = order_worksheet.loc[order_worksheet['OFÍCIO'].str.upper() == 'ELABORAR']
    filtered_worksheet = filtered_worksheet[['Tipo Movimento', 'PROCESSO', 'DESPACHO', 'OFÍCIO', 'PESSOA AFETADA', 'Data Início', 'Função', 'Empresa', 'Lotação', 'Status Pedido']]
    return filtered_worksheet


# VERIFICAR
def get_output_type():
    worksheet = get_order_worksheet()
    if "AVISO PREVIO" in worksheet['Data Início']:
        output_type = f"O aviso prévio poderá ser cumprido na unidade de lotação."
        return output_type
    else:
        output_type = f"Não deverá haver cumprimento de aviso prévio na unidade de lotação."
        return output_type
 

# DESENVOLVER
def get_date():
    day = '10'
    month = 'setembro'
    year = '2024'
    return day, month, year


# DESENVOLVER
def get_oficio_number():
    oficio_number = 207
    return oficio_number


def get_enterprise_worksheet():
    load_dotenv
    FILE_PATH = os.getenv('FILE_PATH2')
    SHEET_NAME = os.getenv('SHEET_NAME2')
    worksheet = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)
    return worksheet


def get_enterprise_elements(enterprise):
    enterprises = get_enterprise_worksheet()
    
    if isinstance(enterprise, list):
        enterprise_name = enterprise[0].upper()
    else:
        enterprise_name = enterprise.upper()
        
    filtered_data = enterprises[enterprises['Empresa'].str.upper() == enterprise_name]
    if filtered_data.empty:
        return None
    
    destinatary = filtered_data.iloc[0]['Representantes']
    detailing = filtered_data.iloc[0]['Detalhamento Empresa']
    # contract_number = filtered_data.iloc[0]['Nº Contrato']
    
    if destinatary and detailing:
        return destinatary, detailing
    else:
        return None


def get_contract(function):
    load_dotenv
    FILE_PATH = os.getenv('FILE_PATH')
    SHEET_NAME = os.getenv('SHEET_NAME3')
    worksheet = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME)
    filtered_data = worksheet[worksheet['Função'].str.upper() == function[0].upper()]
    contract_number = filtered_data.iloc[0]['Contrato'].split()[0]
    contract = filtered_data.iloc[0]['Contrato'].split()[2]
    if contract and contract_number:
        return contract_number, contract
    else:
        return None
    

# print(get_enterprise_elements("CONTEC"))
# get_output_type()
# get_contract("Auxiliar de Serviços Gerais I")