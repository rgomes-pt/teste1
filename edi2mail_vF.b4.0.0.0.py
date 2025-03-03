version_txt = "versão 4.0.0.0 by RGomes"
import os.path
import xml.etree.ElementTree as ET
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle, PageTemplate, Frame#, BaseDocTemplate
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import registerFont
from reportlab.pdfbase.ttfonts import TTFont
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import os
import json
from datetime import datetime
import time
import sys
import logging
import math


if getattr(sys, 'frozen', False):  # Verifica se está em modo executável
    diretorio_atual = os.path.dirname(sys.executable)
else:
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))

caminho_arquivo_db = os.path.join(diretorio_atual, 'database.json')

# Cria caminho para o pasta de armezenamento dos pdfs
caminho_pdf = os.path.join(diretorio_atual, 'pdf_gerados')
# Criar a pasta se ela não existir
os.makedirs(caminho_pdf, exist_ok=True)

# Criar pasta para os ficheiros Excel
caminho_excel = os.path.join(diretorio_atual, 'excel_gerados')
os.makedirs(caminho_excel, exist_ok=True)

# carrega arquivo com o dicionario de artigos
with open(caminho_arquivo_db, 'r') as arquivo:
    dados = json.load(arquivo)
    #Carrega lista de artigos e compradores
    artigos = dados['artigos']
    compradores = dados['compradores']
    ver_db = dados['infos']['vercao']

caminho_arquivo_conf = os.path.join(diretorio_atual, 'config_local.json')
with open(caminho_arquivo_conf, 'r') as arquivo_c:
    config = json.load(arquivo_c)   
    # Carrega local das pastas
    path_prog = config["paths"]["path_prog"]
    lista_xml = config["paths"]["lista_xml"]
    path_XML = config["paths"]["path_XML"]
    # Carrega lista de emails
    lista_email = config["emails"]
    send_email = config["send_email"]
    log_reg = config["log"]
    del_pdf = config["del_pdf"]
    up_lista_xml = config["up_lista_xml"]
    make_excel = config["make_excel"]

pdf_gerado = []     # Criar listas dos pdf gerados
lista_comp = []     # Criar listas dos Compradores

# Configura o logging para salvar as mensagens em um arquivo
logging.basicConfig(
    filename='debug.log',  # Nome do arquivo de log
    level=logging.DEBUG,   # Nível de log
    format='%(asctime)s - %(levelname)s - %(message)s'  # Formato das mensagens
)
if log_reg == True: logging.info("Programa iniciado")

# Função para remover namespaces
def remove_namespace(tag):
    return tag.split('}')[-1] if '}' in tag else tag

# verificar se existe o ficheiro listagem, se nao existir cira-o
if os.path.exists(os.path.join(path_prog, lista_xml)) == False:
    #with open(os.path.join(path_prog, lista_xml), 'w') as fp:
    files = [f for f in os.listdir(path_XML) if f.endswith('.xml')]
    files = [f for f in files if os.path.isfile(path_XML+'/'+f)]  
    list_files_acts = files
    with open(os.path.join(path_prog, lista_xml), 'w') as f:
        f.write(f"{list_files_acts}")
    logging.info("Nao existia lista XML e foi criada:")
    sys.exit()
    
# Lê a lista ficheiros antigos dentro do arquivo
with open(os.path.join(path_prog, lista_xml), 'r') as f:
    list_files_old=f.read() 

files = [f for f in os.listdir(path_XML) if f.endswith('.xml')]
files = [f for f in files if os.path.isfile(path_XML+'/'+f)]  
list_files_acts = files


novos_ficheiros = [x for x in list_files_acts if x not in list_files_old]

if not novos_ficheiros:
    if log_reg == True: logging.info("Não há novos ficheiros.")
    sys.exit()

else:
    # para cada ficheiro novo fazer:
    for i, novo_ficheiro in enumerate(novos_ficheiros):
        Fich_Proc = os.path.join(path_XML, novo_ficheiro)

        # Criar listas para os elementos de OrderDetail
        item_details = []
        
        # extrair as informação do fichiro XML
        # Carregar o arquivo XML
        tree = ET.parse(Fich_Proc)
        root = tree.getroot()

        # Extrair e formatar a Delivery Date
        order_number = root.find('.//{*}OrderNumber')
        order_date = root.find('.//{*}OrderDate')
        doc_type = root.find('.//{*}DocType')
        order_currency = root.find('.//{*}OrderCurrency')
        delivery_date = root.find('.//{*}OtherOrderDates/{*}DeliveryDate')
        buyer_ean = root.find('.//{*}BuyerInformation/{*}EANCode')
        delivery_place_ean = root.find('.//{*}DeliveryPlaceInformation/{*}EANCode')

        order_number=f"{order_number.text if order_number is not None else 'N/A'}"
        order_date=f"{order_date.text if order_date is not None else 'N/A'}"
        data_objeto = datetime.strptime(order_date, "%Y-%m-%dT%H:%M:%S")
        order_date = data_objeto.strftime("%d-%m-%Y")
        doc_type=f"{doc_type.text if doc_type is not None else 'N/A'}"
        order_currency =f"{order_currency.text if order_currency is not None else 'N/A'}"
        delivery_date =f"{delivery_date.text if delivery_date is not None else 'N/A'}"
        data_objeto = datetime.strptime(delivery_date, "%Y-%m-%dT%H:%M:%S")
        delivery_date = data_objeto.strftime("%d-%m-%Y")

        buyer_ean=f"{buyer_ean.text if buyer_ean is not None else 'N/A'}"
        if buyer_ean in compradores:
            buyer_name = compradores[buyer_ean]["n_cliente"]
        else: buyer_name ="Geral"

        delivery_place_ean=f"{delivery_place_ean.text if delivery_place_ean is not None else 'N/A'}"
        if delivery_place_ean in compradores:
            delivery_place_ean = compradores[delivery_place_ean]["tipo"]

        for item in root.findall('.//{*}ItemDetail'):
            line_item_num = item.find('./{*}LineItemNum')
            line_item_num =line_item_num.text if line_item_num is not None else ''
            standard_part_number = item.find('./{*}StandardPartNumber')
            standard_part_number = standard_part_number.text if standard_part_number is not None else ''
            buyer_part_number = item.find('./{*}BuyerPartNumber')
            buyer_part_number = buyer_part_number.text if buyer_part_number is not None else ''
            description = item.find('./{*}ItemDescriptions/{*}Description')  
            description = description.text if description is not None else ''
            quantity_value = item.find('./{*}Quantity/{*}QuantityValue')
            quantity_value = quantity_value.text if quantity_value is not None else ''
            unit_of_measurement = item.find('./{*}Quantity/{*}UnitOfMeasurement')
            unit_of_measurement = unit_of_measurement.text if unit_of_measurement is not None else ''
            net_price = item.find('./{*}Price/{*}NetPrice')
            net_price =  net_price.text if net_price is not None else ''
            line_delivery_date_full = item.find('./{*}LineDeliveryInformation/{*}DeliveryDate')
            line_delivery_date_full = line_delivery_date_full.text if line_delivery_date_full is not None else ''
            data_objeto = datetime.strptime(line_delivery_date_full, "%Y-%m-%dT%H:%M:%S")
            line_delivery_date_full = data_objeto.strftime("%d-%m-%Y")
            int_part_number = "NA"
            uni_medida = "NA"
            peso_medio = "NA"
            uni_per_cx = "NA"           

            if standard_part_number in artigos[buyer_name]:
                int_part_number = artigos[buyer_name][standard_part_number]["cod_interno"]
                uni_medida = artigos[buyer_name][standard_part_number]["uni_med"]
                quantity_value = float(quantity_value) 
                description = artigos[buyer_name][standard_part_number]["descrição"]

                uni_per_cx = int(artigos[buyer_name][standard_part_number]["uni_cx"])
                peso_medio = float(artigos[buyer_name][standard_part_number]["peso_med"])
                enc_por = artigos[buyer_name][standard_part_number]["tipo_quant"]

                if enc_por == "unidade":
                    num_cx = int(round(quantity_value/uni_per_cx, 0))
                    uni_tot_art = int(round(quantity_value,0))

                elif enc_por == "caixa":
                    num_cx = int(round(quantity_value,0))
                    uni_tot_art = int(round(quantity_value * uni_per_cx, 0))

                elif enc_por == "peso":
                    num_cx = int(round(quantity_value/peso_medio/uni_per_cx, 0))
                    uni_tot_art = int(round(quantity_value,0))

            else:
                int_part_number = "Não Listado"
                uni_tot_art = quantity_value
                num_cx = " "

            item_details.append({
                'LineItemNum': line_item_num,
                'StandardPartNumber': standard_part_number,
                'IntPartNumber': int_part_number,
                'BuyerPartNumber': buyer_part_number,
                'BuyerName' : buyer_name,
                'Description': description,
                'QuantityValue': quantity_value,
                'UnitOfMeasurement': uni_medida,
                'NetPrice': net_price,
                'DeliveryDate': line_delivery_date_full,
                'UnidadesTotais': uni_tot_art,
                'TotalCx': num_cx,
                'UnidadePorCx':uni_per_cx,
            })
    
    # Criar o PDF___________________________________
        # Nome do arquivo PDF
        pdf_filename = f"{buyer_name}_{order_number}.pdf" if item_details else 'order_report.pdf'
        
        # Caminho completo do arquivo PDF
        pdf_salvar = os.path.join(caminho_pdf, pdf_filename)
        
        # Criar o PDF
        pdf = SimpleDocTemplate(pdf_salvar, pagesize=letter)
        styles = getSampleStyleSheet()
        
        # Criar um frame que ocupa toda a página, menos o rodapé
        width, height = letter
        frame = Frame(0, 50, width, height - 100, id='normal')  # Margem inferior de 2 cm 
        
        # Função de textos especiais
        def textos(canvas, doc):
            canvas.saveState()
            text = "Este documento DEVE SER VERIFICADO COM A NOTA QUE CHEGA POR EDI"
            canvas.drawString(72, 30, text)
            
            canvas.setFont("Helvetica-Bold", 6)
            text2 = ver_db + " & " + version_txt
            canvas.drawString(72, 15, text2)

            canvas.setFont("Helvetica-Bold", 23)
            text3 = buyer_name
            canvas.drawString(105, 721, text3)

            # Desenhar o checkbox para "Local Carga: LD"
            canvas.setFont("Helvetica", 10)
            canvas.drawString(312, 725, "Data Saída: ______________________")           
            canvas.drawString(312, 702, "Local Carga: LS-             QSD-")
            canvas.rect(390, 700, 12, 12, stroke=1, fill=0)
            canvas.rect(452, 700, 12, 12, stroke=1, fill=0)
            canvas.drawString(312, 670, "Transportador: ____________________")
            canvas.restoreState()

        # Adicionar o frame ao template
        pdf.addPageTemplates([PageTemplate(id='MyPage', frames=[frame], onPage=textos)])
        
        # Criar um novo estilo para o título alinhado à esquerda
        left_aligned_title_style = ParagraphStyle(
            name='LeftAlignedTitle',
            fontName='Helvetica-Bold',
            fontSize=12,
            alignment=0,
            spaceAfter=6,
            leftIndent=20
        )
        # Criar um novo estilo para o título alinhado à esquerda
        left_aligned_title_style2 = ParagraphStyle(
            name='LeftAlignedTitle',
            fontName='Helvetica',
            fontSize=9,
            alignment=0,
            spaceAfter=6,
            leftIndent=20
        )

        # Criar um novo estilo para o rodapé
        footer_style = ParagraphStyle(
            name='FooterStyle',
            fontName='Helvetica',
            fontSize=10,
            alignment=1,
            spaceAfter=12
        )

        # Criar um estilo personalizado para o texto com checkbox
        checkbox_style = ParagraphStyle(
            'CheckboxStyle',
            parent=styles['Normal'],
            spaceBefore=0,
            spaceAfter=0
        )

        def create_checkbox_text(canvas, x, y):
            canvas.saveState()
            canvas.rect(x, y, 10, 10, stroke=1, fill=0)  # Desenha um quadrado de 10x10 pontos
            canvas.restoreState()

        # Criar cabeçalho como uma tabela
        header_data = [
            [Paragraph("Comprador:", left_aligned_title_style)],
            [Paragraph(f"Entrega: {delivery_place_ean}", left_aligned_title_style)],
            [Paragraph(f"Order Number: {order_number}", left_aligned_title_style)],  # Removido o Paragraph do checkbox
            [Paragraph(f"Data Entrega: {delivery_date}", left_aligned_title_style)],
            [Paragraph(f"Order Date: {order_date}", left_aligned_title_style2)],
        ]       

        # Criar tabela para o cabeçalho
        header_table = Table(header_data)
        header_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('ALIGN', (0, 0), (-1, -1), 'CENTER') ]))
        
        elements = []
        elements.append(header_table)
        elements.append(Paragraph("<br/><br/>", styles['Normal']))

        # Tabela
        table_data = [['C.Ext', 'C.Int', 'Descrição', 'Cx', 'Un/Cx', 'Quant.', 'Uni.', 'LOTE', 'Entrega']]
        for item in item_details:
            table_data.append([
                item['StandardPartNumber'], #1
                item['IntPartNumber'],      #2
                item['Description'],        #3
                item['TotalCx'],            #4
                item['UnidadePorCx'],       #5
                item['UnidadesTotais'],     #6
                item['UnitOfMeasurement'],  #7
                '',                    
                item['DeliveryDate']        #9
            ])

        # Definir larguras fixas para as colunas
        col_widths = [0.8 * inch, 0.4 * inch, 3.5 * inch, 0.45 * inch, 0.45 * inch, 0.5 * inch, 0.35 * inch, 0.8 * inch, 0.6 * inch]

        # Definir altura fixa para as linhas (em pontos)
        row_heights = [19] * len(table_data)  # 20 pontos para cada linha

        # Criar a tabela com larguras de coluna e alturas de linha fixas
        table = Table(table_data, colWidths=col_widths, rowHeights=row_heights)

        # Definir estilo da tabela com tamanho de fonte menor
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Alinhamento vertical central para todas as células
            # Fonte padrão Helvetica para todas as células
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            # Cabeçalho em negrito
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            # Tamanho 6 para colunas específicas a partir da segunda linha
            ('FONTSIZE', (0, 1), (0, -1), 6),  # Coluna 'C.Ext'
            ('FONTSIZE', (1, 1), (1, -1), 6),  # Coluna 'C.Int'
            ('FONTSIZE', (-1, 1), (-1, -1), 6),  # Última coluna 'Entrega'
            ('TOPPADDING', (0, 1), (1, -1), 4),  # Padding superior para fontes pequenas
            ('TOPPADDING', (-1, 1), (-1, -1), 4),  # Padding superior para última coluna
            ('TOPPADDING', (0, 1), (0, -1), 6),  # Padding superior adicional para C.Ext
            ('TOPPADDING', (1, 1), (1, -1), 6),  # Padding superior adicional para C.Int
            ('TOPPADDING', (-1, 1), (-1, -1), 6),  # Padding superior adicional para Entrega
            # Restante das configurações específicas
            ('FONTSIZE', (2, 1), (3, -1), 9),                   # Coluna 'Descrição' - Tamalho Fonte
            ('FONTNAME', (2, 1), (3, -1), 'Helvetica-Bold'),    # Coluna 'Descrição'
            ('FONTSIZE', (3, 1), (3, -1), 11),                   # Coluna 'Cx'             
            ('FONTNAME', (3, 1), (3, -1), 'Courier-Bold'),      # Coluna 'Cx'
            ('FONTSIZE', (4, 1), (4, -1), 11),  # Coluna 'Cx'            
            ('FONTNAME', (4, 1), (4, -1), 'Courier'),  # Coluna 'Un/Cx'
            ('FONTSIZE', (4, 1), (4, -1), 11),  # Coluna 'Un/Cx'
            ('FONTNAME', (5, 1), (5, -1), 'Courier'),  # Coluna 'Quant.'
            ('FONTSIZE', (5, 1), (5, -1), 11),  # Coluna 'Quant.'
            ('TOPPADDING', (3, 1), (5, -1), -1),  # Padding inferior para fontes grandes
            ('BOTTOMPADDING', (0, 0), (-1, 0), 0),
            ('GRID', (0, 0), (-1, -1), 0.8, colors.black),
            ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ]))
        
        elements.append(table)
        
        # Espaçamento entre as tabelas
        from reportlab.platypus import Spacer
        elements.append(Spacer(1, 40))
        
        
        # Cria tabela de divisão de paletes se o cliente for a Mercadona ############################
        if buyer_name == 'Mercadona':
            # Espaçamento entre as tabelas
            from reportlab.platypus import Spacer
            elements.append(Spacer(1, 40))

            table_carga = [['Descrição', 'NºPal.', 'NºPal.Comp', 'Kg/Pal', 'Altura']]
            for item in item_details:

                standard_part_number = item['StandardPartNumber']
                buyer_name = item['BuyerName']

                uni_per_cx = int(artigos[buyer_name][standard_part_number]["uni_cx"])

                peso_medio= (artigos[buyer_name][standard_part_number]["peso_med"])

                gra_base = int(artigos[buyer_name][standard_part_number]["gra_base"])
                gra_alt = int(artigos[buyer_name][standard_part_number]["gra_alt"])

                peso_medio_inf= peso_medio - 0.05
                peso_medio_sup= peso_medio + 0.05

                kg_totais = item['UnidadesTotais']  # Vai buscar o número total de unidades (no caso da mercadona é em kg)

                unidades_totais_inf=int(round((kg_totais/peso_medio_inf),0))
                unidades_totais_sup=int(round((kg_totais/peso_medio_sup),0))
                
                un_cx = item['UnidadePorCx']

                Maximo_Uni_Pal=gra_base * gra_alt * un_cx
                
                if enc_por == "peso":
                    Pal_min = round((unidades_totais_inf/Maximo_Uni_Pal),1)
                    Pal_max = round((unidades_totais_sup/Maximo_Uni_Pal),1)
                    
                    Pal_Min_Ared=math.ceil(Pal_min)
                    Pal_Max_Ared=math.ceil(Pal_max)
                    
                    Kg_Pal_min = int(kg_totais / Pal_Min_Ared)
                    Kg_Pal_max = int(kg_totais / Pal_Max_Ared)

                    Alt_min = round(Kg_Pal_min/peso_medio_inf/(un_cx*gra_base),1)
                    Alt_max = round(Kg_Pal_max/peso_medio_sup/(un_cx*gra_base),1)

                    table_carga.append([
                        item['Description'], #descrição
                        f"{Pal_min} - {Pal_max}", #nº de paletes
                        f"{Pal_Min_Ared} - {Pal_Max_Ared}", #Paletes aredondadas para cima
                        f"{Kg_Pal_min} - {Kg_Pal_max}", #unidades totais
                        f"{Alt_min} - {Alt_max}", #altura das grades
                        ])

                if enc_por == "unidade":
                    Pal_min = round((kg_totais/Maximo_Uni_Pal),1)
                    Pal_Min_Ared=math.ceil(Pal_min)
                    Kg_Pal_min = int(kg_totais / Pal_Min_Ared)
                    Alt_min = round(Kg_Pal_min/(un_cx*gra_base),1)
                    table_carga.append([
                        item['Description'], #descrição
                        f"{Pal_min}", #nº de paletes
                        f"{Pal_Min_Ared}", #Paletes aredondadas para cima
                        f"{Kg_Pal_min} - Un/Pal", #unidades totais
                        f"{Alt_min}", #altura das grades
                        ])

            # Define column widths for the new table
            new_col_widths = [3.2 * inch, 0.8 * inch, 0.8 * inch, 0.8 * inch, 0.8 * inch]  # Tamanho das colunas da tabela 2

            # Create the new table
            Palete_table = Table(table_carga, colWidths=new_col_widths)

            # Style the new table
            Palete_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('GRID', (0, 0), (-1, -1), 0.8, colors.black),
            ]))

            # Add spacing between tables
            elements.append(Paragraph("<br/>", styles['Normal']))

            # Add the new table to elements
            elements.append(Palete_table)
            # Fim tabela de divisão de paletes ########################################################
        if buyer_name == 'Pingo Doce' and order_number[:2] == '49':
                            
            # Espaçamento entre as tabelas
            from reportlab.platypus import Spacer
            elements.append(Spacer(1, 40))        
        
        
            styles = getSampleStyleSheet()
            style = styles['Normal']  # Pode usar o estilo padrão ou customizar um novo
            style.fontSize = 24  # Tamanho da fonte em pontos (24 = letras grandes)
            style.fontName = 'Helvetica-Bold'  # Usando a fonte Helvetica em negrito
            style.textColor = colors.red  # Cor vermelha para o texto
            style.alignment = 1  # Alinhamento centralizado (0 = esquerda, 1 = centralizado, 2 = direita)
            
            # Texto a ser adicionado
            texto = "POSSÍVEL DEVOLUÇÃO"
            
            # Criar o parágrafo com o texto e o estilo
            paragrafo = Paragraph(texto, style)
            
            # Adicionar o parágrafo à lista de elementos
            elements.append(paragrafo)
    
        # Construir o PDF
        pdf.build(elements)
        pdf_gerado.append(pdf_filename)
        if buyer_name not in lista_comp:        #Adiciona o nome do comprador à lista de compradores caso nao exista
            lista_comp.append(buyer_name)
        if log_reg == True: logging.debug(f"PDF '{pdf_filename}' foi gerado com sucesso.")

    