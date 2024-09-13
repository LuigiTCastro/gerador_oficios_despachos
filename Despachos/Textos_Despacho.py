import os
import re
from dotenv import load_dotenv
from Dicts import number_in_full
from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Indenter, Image


def movement_type_txt(orders_qtt, movement_type, employee, function, workplace, new_function=None, new_workplace=None):    
    if function is None:
        function = ['']
    if workplace is None:
        workplace = ['']
        
    DEVOLUÇÃO_TXT = f"""Considerando as informações contidas no processo em epígrafe, 
    autorizo a {movement_type.lower()} do(a) colaborador(a) {employee[0]}, ocupante da função de 
    {function[0]}, com lotação no(a) {workplace[0]}."""
    
    CONTRATAÇÃO_TXT = f"""Considerando as informações contidas no processo em epígrafe, 
    autorizo a {movement_type.lower()} de 01 (uma) vaga de {function[0]}, com lotação no(a) {workplace[0]}."""
    
    SUBSTITUIÇÃO_TXT = f"""Considerando as informações contidas no processo em epígrafe, 
    autorizo a {movement_type.lower()} do(a) colaborador(a) {employee[0]}, ocupante da função de 
    {function[0]}, com lotação no(a) {workplace[0]}."""
    
    MUDANÇA_FUNÇÃO_TXT = f"""Considerando as informações contidas no processo em epígrafe, 
    autorizo a {movement_type.lower()} do(a) colaborador(a) {employee[0]}, de {function[0]} para {new_function}, 
    com lotação no(a) {workplace[0]}."""
    
    MUDANÇA_LOTAÇÃO_TXT = f"""Considerando as informações contidas no processo em epígrafe, 
    autorizo a {movement_type.lower()} do(a) colaborador(a) {employee[0]}, ocupante da função de 
    {function[0]}, do(a) {workplace[0]} para o(a) {new_workplace}."""
    
    
    CONTRATAÇÕES_TXT1 = f"""Considerando as informações contidas no processo em epígrafe, autorizo a {movement_type.lower()} 
    de {orders_qtt:02} ({number_in_full[orders_qtt]}) vagas de {function[0]}, com lotação no(a) {workplace[0]}."""
    
    
    if movement_type.upper() == "DEVOLUÇÃO":
        return DEVOLUÇÃO_TXT
    elif movement_type.upper() == "CONTRATAÇÃO":
        if orders_qtt == 1:
            return CONTRATAÇÃO_TXT
        elif orders_qtt > 1:
            return CONTRATAÇÕES_TXT1
    elif movement_type.upper() == "SUBSTITUIÇÃO":
        return SUBSTITUIÇÃO_TXT
    elif movement_type.upper() == "MUDANÇA DE FUNÇÃO":
        return MUDANÇA_FUNÇÃO_TXT
    elif movement_type.upper() == "MUDANÇA DE LOTAÇÃO":
        return MUDANÇA_LOTAÇÃO_TXT


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


# DESENVOLVER
# def get_date():
#     month = 'agosto'
#     year = '2024'
#     return month, year


def create_file_name(movement_type, process, employee, function, workplace, new_function=None, new_workplace=None):
    load_dotenv()
    DESPACHO_FOLDER = os.getenv('DESPACHO_FOLDER')
    if not os.path.exists(DESPACHO_FOLDER):
        os.makedirs(DESPACHO_FOLDER)  
    
    reduced_function = format_name(function[0])
    reduced_workplace = format_name(workplace[0])
    reduced_employee = ' '.join(employee[0].split()[:1])
    
    
    if movement_type == "Devolução":
        file_name = f"SGP-CAC-{process} - {movement_type} - {reduced_function} - {reduced_workplace} - {reduced_employee}.pdf"
    elif movement_type == "Contratação":
        file_name = f"SGP-CAC-{process} - {movement_type} - {reduced_function} - {reduced_workplace}.pdf"
    elif movement_type == "Substituição":
        file_name = f"SGP-CAC-{process} - {movement_type} - {reduced_function} - {reduced_workplace} - {reduced_employee}.pdf"
    elif movement_type == "Mudança de Função":
        file_name = f"SGP-CAC-{process} - {movement_type} - {reduced_function}.pdf"
    elif movement_type == "Mudança de Lotação":
        file_name = f"SGP-CAC-{process} - {movement_type} - {reduced_function}.pdf"
        
    # safe_file_name = "".join([c for c in file_name if c.isalpha() or c.isdigit() or c in (' ', '-', '_')]).rstrip()
    safe_file_name = re.sub(r'[<>:"/\\|?*]', '', file_name)
    if len(safe_file_name) > 100:
        safe_file_name = safe_file_name[:100]
        
    file_path = os.path.join(DESPACHO_FOLDER, safe_file_name)
    return file_path


def create_simple_file_pdf(orders_qtt, movement_type, process, employee, function, workplace, new_function=None, new_workplace=None):
    file_path = create_file_name(movement_type, process, employee, function, workplace)
    # month, year = get_date()
    # date = f" __ de {month} de {year}"
    
    # Configure the PDF
    doc = SimpleDocTemplate(file_path, pagesize=A4)
    styles = getSampleStyleSheet()
    
    # Custom styles
    styles.add(ParagraphStyle(name='Times12', fontName='Times-Roman', fontSize=12, leading=18))
    styles.add(ParagraphStyle(name='Times12Bold', fontName='Times-Bold', fontSize=12, leading=18))
    styles.add(ParagraphStyle(name='Times12Center', fontName='Times-Roman', fontSize=12, leading=18, alignment=1))
    styles.add(ParagraphStyle(name='Times12CenterBold', fontName='Times-Bold', fontSize=12, leading=18, alignment=1))
    styles.add(ParagraphStyle(name='JustifyIndented', fontName='Times-Roman', fontSize=12, leading=18, alignment=4, firstLineIndent=2*cm))

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
    header = "ESTADO DO CEARÁ<br/>PODER JUDICIÁRIO<br/>SECRETARIA DE GESTÃO DE PESSOAS"
    elements.append(Paragraph(header, styles['Times12CenterBold']))
    elements.append(Spacer(1, 2*cm))
    
    # Add despacho title
    elements.append(Paragraph("DESPACHO", styles['Times12CenterBold']))
    elements.append(Spacer(1, 2*cm))
    
    # Add assunto
    elements.append(Paragraph(f"Assunto: {movement_type} de colaborador terceirizado.", styles['Times12']))
    
    # Add process number
    elements.append(Paragraph(f"Nº do Processo: {process}", styles['Times12']))
    elements.append(Spacer(1, 2*cm))
    
    # Add main content with indentation
    main_content = movement_type_txt(orders_qtt, movement_type, employee, function, workplace, new_function, new_workplace)
    elements.append(Paragraph(main_content, styles['JustifyIndented']))
    elements.append(Spacer(1, 1*cm))
    
    # Add the text before the date
    destination_content = "À Coordenadoria de Acompanhamento de Contratos."
    elements.append(Indenter(left=2*cm))
    elements.append(Paragraph(destination_content, styles['Times12']))
    elements.append(Indenter(left=-2*cm))
    
    # Add footer with date and signatory
    footer_content = f"Fortaleza/CE, data registrada pelo sistema."
    elements.append(Indenter(left=2*cm))
    elements.append(Paragraph(footer_content, styles['Times12']))
    elements.append(Indenter(left=-2*cm))
    elements.append(Spacer(1, 3*cm))  # Espaçamento de 6 linhas
    
    # Add signature
    # signature_content = "Felipe de Albuquerque Mourão<br/>"
    signature_content = "Victor Alves Dias<br/>"
    secretary = "Secretário de Gestão de Pessoas, em substituição"
    elements.append(Paragraph(signature_content, styles['Times12CenterBold']))
    elements.append(Paragraph(secretary, styles['Times12Center']))
    
    # Build the PDF
    doc.build(elements)


# DESENVOLVER
def create_file_pdf_with_table(orders_qtt, movement_type, process, employee, function, workplace, new_function=None, new_workplace=None):
    file_path = create_file_name(movement_type, process, employee, function, workplace)
    # month, year = get_date()
    # date = f" __ de {month} de {year}"
    
    # Configure the PDF
    doc = SimpleDocTemplate(file_path, pagesize=A4)
    styles = getSampleStyleSheet()
    
    # Custom styles
    styles.add(ParagraphStyle(name='Times12', fontName='Times-Roman', fontSize=12, leading=18))
    styles.add(ParagraphStyle(name='Times12Bold', fontName='Times-Bold', fontSize=12, leading=18))
    styles.add(ParagraphStyle(name='Times12Center', fontName='Times-Roman', fontSize=12, leading=18, alignment=1))
    styles.add(ParagraphStyle(name='Times12CenterBold', fontName='Times-Bold', fontSize=12, leading=18, alignment=1))
    styles.add(ParagraphStyle(name='JustifyIndented', fontName='Times-Roman', fontSize=12, leading=18, alignment=4, firstLineIndent=2*cm))

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
    header = "ESTADO DO CEARÁ<br/>PODER JUDICIÁRIO<br/>SECRETARIA DE GESTÃO DE PESSOAS"
    elements.append(Paragraph(header, styles['Times12CenterBold']))
    elements.append(Spacer(1, 2*cm))
    
    # Add despacho title
    elements.append(Paragraph("DESPACHO", styles['Times12CenterBold']))
    elements.append(Spacer(1, 2*cm))
    
    # Add assunto
    elements.append(Paragraph(f"Assunto: {movement_type} de colaborador terceirizado.", styles['Times12']))
    
    # Add process number
    elements.append(Paragraph(f"Nº do Processo: {process}", styles['Times12']))
    elements.append(Spacer(1, 2*cm))
    
    # Add main content with indentation
    main_content = movement_type_txt(movement_type, employee, function, workplace, new_function, new_workplace)
    elements.append(Paragraph(main_content, styles['JustifyIndented']))
    elements.append(Spacer(1, 1*cm))
    
    # Add the text before the date
    destination_content = "À Coordenadoria de Acompanhamento de Contratos."
    elements.append(Indenter(left=2*cm))
    elements.append(Paragraph(destination_content, styles['Times12']))
    elements.append(Indenter(left=-2*cm))
    
    # Add footer with date and signatory
    footer_content = f"Fortaleza/CE, data registrada pelo sistema."
    elements.append(Indenter(left=2*cm))
    elements.append(Paragraph(footer_content, styles['Times12']))
    elements.append(Indenter(left=-2*cm))
    elements.append(Spacer(1, 3*cm))  # Espaçamento de 6 linhas
    
    # Add signature
    signature_content = "Felipe de Albuquerque Mourão<br/>"
    secretary = "Secretário de Gestão de Pessoas"
    elements.append(Paragraph(signature_content, styles['Times12CenterBold']))
    elements.append(Paragraph(secretary, styles['Times12Center']))
    
    # Build the PDF
    doc.build(elements)