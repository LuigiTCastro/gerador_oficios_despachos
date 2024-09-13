from Dados_Ofícios import get_oficio_number, get_order_worksheet
from Textos_Ofícios import create_simple_file_pdf



def convert_name(name):
    prepositions = ['e', 'a', 'de', 'em', 'para', 'com', 'por', 'da', 'do', 'das', 'dos']
    words = name.lower().split()
    convertted_words = [word.capitalize() if word not in prepositions else word for word in words]
    convertted_name = ' '.join(convertted_words)
    return convertted_name


def generate_order():
    orders = get_order_worksheet()
    oficio_number = get_oficio_number()
    process_elements = {}
    
    for _, row in orders.iterrows():
        movement_type = row['Tipo Movimento']
        process = row['PROCESSO']
        employee = row['PESSOA AFETADA']
        function = row['Função']
        enterprise = row['Empresa']
        workplace = row['Lotação']
        office = row['OFÍCIO']
        
        if process not in process_elements:
            process_elements[process] = {
                'movements': [],
                'employee': [],
                'function': [],
                'enterprise': [],
                'workplace':[] 
            }
        
        if movement_type == "Devolução" or movement_type == "Contratação" or movement_type == "Substituição":
            process_elements[process]['movements'].append(movement_type)
            process_elements[process]['employee'].append(convert_name(employee))
            process_elements[process]['function'].append(function)
            process_elements[process]['enterprise'].append(enterprise.upper())
            process_elements[process]['workplace'].append(workplace)
        
    for process, elements in process_elements.items():
        movements = elements['movements']
        employee = elements['employee']
        function = elements['function']
        enterprise = elements['enterprise']
        workplace = elements['workplace']        

        if len(movements) == 1 and movements[0] == 'Contratação':
            orders_qtt = len(movements)
            movement_type = movements[0]
            create_simple_file_pdf(orders_qtt, oficio_number, enterprise, movement_type, employee, function, workplace)
            print(f'CONTRATAÇÃO [1]: {process, elements}\n') 
        
        elif len(movements) == 1 and movements[0] == 'Devolução':
            orders_qtt = len(movements)
            movement_type = movements[0]
            create_simple_file_pdf(orders_qtt, oficio_number, enterprise, movement_type, employee, function, workplace)
            print(f'DEVOLUÇÃO [1]: {process, elements}\n')
            
        elif len(movements) == 2 and 'Devolução' in movements and 'Contratação' in movements:
            orders_qtt = 1
            movement_type = movements[0]
            create_simple_file_pdf(orders_qtt, oficio_number, enterprise, movement_type, employee, function, workplace)
            oficio_number += 1
            movement_type = movements[1]
            create_simple_file_pdf(orders_qtt, oficio_number, enterprise, movement_type, employee, function, workplace)
            print(f'SUBSTITUIÇÃO [1]: {process, elements}\n')

        elif len(movements) > 1 and all(m == 'Contratação' for m in movements):
            orders_qtt = len(movements)
            movement_type = movements[0]
            create_simple_file_pdf(orders_qtt, oficio_number, enterprise, movement_type, employee, function, workplace)
            print(f'CONTRATAÇÃO [2/+]: {process, elements}\n')
            
        oficio_number += 1


generate_order()
