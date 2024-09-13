import os
import pandas as pd
from dotenv import load_dotenv
from Textos_Despacho import create_simple_file_pdf


load_dotenv()


def get_order_worksheet():
    FILE_PATH = os.getenv('FILE_PATH')
    SHEET_NAME = os.getenv('SHEET_NAME')
    order_worksheet = pd.read_excel(FILE_PATH, sheet_name=SHEET_NAME, skiprows=6)
    filtered_worksheet = order_worksheet.loc[order_worksheet['DESPACHO'].str.upper() == 'ELABORAR']
    filtered_worksheet = filtered_worksheet[['Tipo Movimento', 'PROCESSO', 'DESPACHO', 'OFÍCIO', 'PESSOA AFETADA', 'Data Início', 'Função', 'Lotação', 'Status Pedido']]
    return filtered_worksheet


def convert_name(name):
    prepositions = ['e', 'a', 'de', 'em', 'para', 'com', 'por', 'da', 'do', 'das', 'dos']
    words = name.lower().split()
    convertted_words = [word.capitalize() if word not in prepositions else word for word in words]
    convertted_name = ' '.join(convertted_words)
    return convertted_name


def create_order():
    orders = get_order_worksheet()
    process_elements = {}
    
    for _, row in orders.iterrows():
        movement_type = row['Tipo Movimento']
        process = row['PROCESSO']
        employee = row['PESSOA AFETADA']
        function = row['Função']
        workplace = row['Lotação']
        
        if process not in process_elements:
            process_elements[process] = {
                'movements': [],
                'employee': [],
                'function': [],
                'workplace':[] 
            }
        
        if movement_type == "Devolução" or movement_type == "Contratação" or movement_type == "Substituição":
            process_elements[process]['movements'].append(movement_type)
            process_elements[process]['employee'].append(convert_name(employee))
            process_elements[process]['function'].append(function)
            process_elements[process]['workplace'].append(workplace)
                
    for process, elements in process_elements.items():
        movements = elements['movements']
        employee = elements['employee']
        function = elements['function']
        workplace = elements['workplace']

        if len(movements) == 1 and movements[0] == 'Contratação':
            orders_qtt = len(movements)
            movement_type = movements[0]
            create_simple_file_pdf(orders_qtt, movement_type, process, employee, function, workplace)
            print(f'CONTRATAÇÃO [1]: {process, elements}\n')
            
        elif len(movements) == 1 and movements[0] == 'Devolução':
            orders_qtt = len(movements)
            movement_type = movements[0]
            create_simple_file_pdf(orders_qtt, movement_type, process, employee, function, workplace)
            print(f'DEVOLUÇÃO [1]: {process, elements}\n')
            
        elif len(movements) == 2 and 'Devolução' in movements and 'Contratação' in movements:
            orders_qtt = len(movements)
            movement_type = 'Substituição'
            create_simple_file_pdf(orders_qtt, movement_type, process, employee, function, workplace)
            print(f'SUBSTITUIÇÃO [1]: {process, elements}\n')
            
        elif len(movements) > 1 and all(m == 'Contratação' for m in movements):
            orders_qtt = len(movements)
            movement_type = movements[0]
            create_simple_file_pdf(orders_qtt, movement_type, process, employee, function, workplace)
            print(f'CONTRATAÇÃO [2/+]: {process, elements}\n')


create_order()




