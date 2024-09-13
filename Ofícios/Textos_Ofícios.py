import os
import pandas as pd
from Dicts import number_in_full
from dotenv import load_dotenv
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Indenter, Image
from Dados_Ofícios import get_contract, get_output_type, get_date, get_enterprise_elements



def format_name(name):
    symbols = ['-','(',')']
    numbers = ['I','II','III','IV','V']
    prepositions = ['e', 'a', 'de', 'em', 'para', 'com', 'por', 'da', 'do', 'das', 'dos']
    splited_name = name.split()
    acronym = []
    
    if len(splited_name) == 1:
        result = name[:6].upper()
    elif len(splited_name) > 1:
        for n in splited_name:
            if n not in prepositions and n not in numbers:
                letter = n[0].upper()
                acronym.append(letter)
            if n in numbers:
                acronym.append(n)
        result = ''.join(acronym)
    return result


def movement_type_txt(orders_qtt, movement_type, employee, function, workplace, contract_number, new_function=None, new_workplace=None): 
    output_type = get_output_type()
    day, month, year = get_date()
    day_from = int(day) + 3
    
    DEVOLUÇÃO_TXT = f"""De ordem do Secretário de Gestão de Pessoas e com amparo no contrato n. {contract_number}, 
    solicito a {movement_type.lower()} do(a) colaborador(a) terceirizado(a) {employee[0]}, ocupante da função de {function[0]}, 
    lotado(a) no(a) {workplace[0]}, a partir de {day_from:02}/07/{year}."""
    
    CONTRATAÇÃO_TXT = f"""De ordem do Secretário de Gestão de Pessoas e com amparo no contrato n. {contract_number}, 
    formalizo o pedido de {movement_type.lower()} de 01 (uma) vaga de {function[0]}, com lotação no(a) {workplace[0]}."""
    
    MUDANÇA_FUNÇÃO_TXT = f"""De ordem do Secretário de Gestão de Pessoas e com amparo no contrato n. {contract_number},
    solicito a {movement_type.lower()} do(a) colaborador(a) terceirizado(a) {employee[0]}, de {function[0]} para {new_function}, 
    com lotação no(a) {workplace[0]}."""
    
    MUDANÇA_LOTAÇÃO_TXT = f"""De ordem do Secretário de Gestão de Pessoas e com amparo no contrato n. {contract_number},
    solicito a {movement_type.lower()} do(a) colaborador(a) terceirizado(a) {employee[0]}, ocupante da função de 
    {function[0]}, do(a) {workplace[0]} para o(a) {new_workplace}."""
    
    
    CONTRATAÇÕES_TXT1 = f"""De ordem do Secretário de Gestão de Pessoas e com amparo no contrato n. {contract_number},
    formalizo o pedido de {movement_type.lower()} de {orders_qtt:02} ({number_in_full[orders_qtt]}) vagas de {function[0]}, 
    com lotação no(a) {workplace[0]}."""
    
    
    if movement_type.upper() == "DEVOLUÇÃO":
        return DEVOLUÇÃO_TXT, output_type
    elif movement_type.upper() == "CONTRATAÇÃO":
        if orders_qtt == 1:
            return CONTRATAÇÃO_TXT
        elif orders_qtt > 1:
            return CONTRATAÇÕES_TXT1
    elif movement_type.upper() == "MUDANÇA DE FUNÇÃO":
        return MUDANÇA_FUNÇÃO_TXT
    elif movement_type.upper() == "MUDANÇA DE LOTAÇÃO":
        return MUDANÇA_LOTAÇÃO_TXT


def create_oficio_name(oficio_number, enterprise, movement_type, employee, function, workplace, new_function=None, new_workplace=None):
    load_dotenv()
    OFICIO_FOLDER = os.getenv('OFICIO_FOLDER')
    if not os.path.exists(OFICIO_FOLDER):
        os.makedirs(OFICIO_FOLDER)  
    
    day, month, year = get_date()
    reduced_enterprise = enterprise[0]
    reduced_function = format_name(function[0])
    reduced_workplace = format_name(workplace[0])
    reduced_employee = ' '.join(employee[0].split()[:1])
    # reduced_function = format_name(function)
    # reduced_workplace = format_name(workplace)
    # reduced_employee = ' '.join(employee.split()[:3])
    
    if movement_type == "Devolução":
        file_name = f"{oficio_number} {year} {reduced_enterprise} - {movement_type} - {reduced_function} - {reduced_workplace} - {reduced_employee}.pdf"
    elif movement_type == "Contratação":
        file_name = f"{oficio_number} {year} {reduced_enterprise} - {movement_type} - {reduced_function} - {reduced_workplace}.pdf"
    elif movement_type == "Mudança de Função":
        file_name = f"{oficio_number} {year} {reduced_enterprise} - {movement_type} - {reduced_function}.pdf"
    elif movement_type == "Mudança de Lotação":
        file_name = f"{oficio_number} {year} {reduced_enterprise} - {movement_type} - {reduced_function}.pdf"
        
    file_path = os.path.join(OFICIO_FOLDER, file_name)
    return file_path


def create_simple_file_pdf(orders_qtt, oficio_number, enterprise, movement_type, employee, function, workplace, new_function=None, new_workplace=None):
    load_dotenv()
    day, month, year = get_date()
    file_path = create_oficio_name(oficio_number, enterprise, movement_type, employee, function, workplace)
    
    # Configure the PDF
    doc = SimpleDocTemplate(file_path, pagesize=A4)
    styles = getSampleStyleSheet()
    
    # Custom styles
    styles.add(ParagraphStyle(name='Times12', fontName='Times-Roman', fontSize=12, leading=18))
    styles.add(ParagraphStyle(name='Times12Bold', fontName='Times-Bold', fontSize=12, leading=18))
    styles.add(ParagraphStyle(name='Times12Center', fontName='Times-Roman', fontSize=12, leading=18, alignment=1))
    styles.add(ParagraphStyle(name='Times12CenterBold', fontName='Times-Bold', fontSize=12, leading=18, alignment=1))
    styles.add(ParagraphStyle(name='JustifyIndented', fontName='Times-Roman', fontSize=12, leading=18, alignment=4, firstLineIndent=2*cm))
    styles.add(ParagraphStyle(name='JustifyIndentedBold', fontName='Times-Bold', fontSize=12, leading=18, alignment=4, firstLineIndent=2*cm))
    styles.add(ParagraphStyle(name='Times12Right', fontName='Times-Roman', fontSize=12, leading=18, alignment=2))

    # Define the elements to add to the PDF
    elements = []
    
    # Add image if provided
    IMAGE_PATH = os.getenv('IMAGE_PATH')
    if IMAGE_PATH:
        logo = Image(IMAGE_PATH)
        logo.drawHeight = 2*cm  # Adjust the height
        logo.drawWidth = 2*cm   # Adjust the width
        elements.append(logo)
        elements.append(Spacer(1, 0.5*cm))  # Add some space below the image
        
    # Add header
    header = "ESTADO DO CEARÁ<br/>PODER JUDICIÁRIO<br/>COORDENADORIA DE ACOMPANHAMENTO DE CONTRATOS"
    elements.append(Paragraph(header, styles['Times12CenterBold']))
    elements.append(Spacer(1, 1*cm))
    
    # Add oficio title
    elements.append(Paragraph(f"Ofício nº {oficio_number}/{year}/CAC", styles['Times12']))
    elements.append(Spacer(1, 0.5*cm))
    
    # Add date
    date = f"{day} de {month} de {year}"
    date_content = f"Fortaleza, {date}."
    elements.append(Paragraph(date_content, styles['Times12Right']))
    elements.append(Spacer(1, 0.5*cm))
    
    # Add destination
    treatment = 'Ao(À) Senhor(a)'
    destinatary, detailing = get_enterprise_elements(enterprise)
    contract_number, contract = get_contract(function)
    city = 'Fortaleza - Ceará'
    elements.append(Paragraph(f"{treatment}<br/>{destinatary}<br/>{detailing}<br/>{city}<br/>", styles['Times12']))
    elements.append(Spacer(0.5, 0.5*cm))
    
    # Add subject
    subject = f"Assunto: {movement_type} de colaborador terceirizado."
    elements.append(Paragraph(f"{subject}", styles['Times12']))
    elements.append(Spacer(1, 1*cm))
    
    # Add introduction
    introduction = "Prezado(a) senhor(a),"
    elements.append(Paragraph(f"{introduction}", styles['JustifyIndented']))
    elements.append(Spacer(1, 0.5*cm))
    
    # Add main content
    if movement_type == "Devolução":
        main_content, output_type = movement_type_txt(orders_qtt, movement_type, employee, function, workplace, contract_number, new_function, new_workplace)
        elements.append(Paragraph(main_content, styles['JustifyIndented']))
        elements.append(Spacer(1, 0.5*cm))
        elements.append(Paragraph(output_type, styles['JustifyIndentedBold']))
        elements.append(Spacer(1, 1*cm))
    else:
        main_content = movement_type_txt(orders_qtt, movement_type, employee, function, workplace, contract_number, new_function, new_workplace)
        elements.append(Paragraph(main_content, styles['JustifyIndented']))
        elements.append(Spacer(1, 1*cm))
    
    # Add closure
    closure = "Atenciosamente,"
    elements.append(Paragraph(closure, styles['JustifyIndented']))
    elements.append(Spacer(1, 3*cm))
    
    # Add signature
    subscriber = "Fransilvia Oliveira Paiva<br/>"
    department = "Coordenadora de Acompanhamento de Contratos"
    elements.append(Paragraph(subscriber, styles['Times12CenterBold']))
    elements.append(Paragraph(department, styles['Times12Center']))
    
    # Build the PDF
    doc.build(elements)
    
    
    