from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import tkinter as tk
import os
import openpyxl
from openpyxl import Workbook, load_workbook
from tkinter import messagebox



geckodriver_path = r'C:\Users\User\Desktop\RonaldoBot\geckodriver.exe'
profile_path = r'C:\Users\User\AppData\Roaming\Mozilla\Firefox\Profiles\ddxg8c5z.default-release-1'


firefox_profile = webdriver.FirefoxProfile(profile_path)

# Configurar o serviço do Firefox
service = Service(geckodriver_path)

# Configurar o perfil do Firefox
options = Options()
options.profile = profile_path

# driver = webdriver.Firefox(service=service, options=options)
driver = webdriver.Firefox(service=service, options=options)


url = 'https://app.gtowizard.com/'
# url = 'https://google.com/'

driver.get(url)




########################################################## Função para extrair informações das três primeiras cartas
def extract_first_three_cards_info(driver):
    try:
        # Esperar até 20 segundos para o elemento aparecer
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.gw_table_body_row.repfloptblrow.gw_table_body_row_hoverable.gw_hvr.gw_hvr_pcn.cursor-pointer'))
        )
        
        container = driver.find_element(By.CSS_SELECTOR, '.gw_table_body_row.repfloptblrow.gw_table_body_row_hoverable.gw_hvr.gw_hvr_pcn.cursor-pointer')

        card_info = []
        
        # Encontrar todos os blocos de cartas até 3
        card_blocks = container.find_elements(By.CSS_SELECTOR, ".cardsymbols_block")[:3]
        
        # Iterar sobre os três primeiros blocos de cartas
        for card_block in card_blocks:
            # Extrair o valor textual da carta
            card_value = card_block.find_element(By.CSS_SELECTOR, ".cardsymbols_value").text.strip()

            # Encontrar o código SVG do ícone da carta
            card_symbol_element = card_block.find_element(By.CSS_SELECTOR, ".cardsymbols_symbol svg")
            svg_path = card_symbol_element.find_element(By.CSS_SELECTOR, 'path').get_attribute('d')
            
            # Verifica o conteúdo do path para atribuir o nome correspondente
            if "M3.89" in svg_path:
                suit_name = "PAUS"
            elif "M30.97" in svg_path:
                suit_name = "COPAS"
            elif "M28.29" in svg_path:
                suit_name = "OUROS"
            else:
                suit_name = "ESPADAS"

  
            card_info.append(f"{card_value} de {suit_name}")
        
        return card_info
    
    except Exception as e:
        print(f"Erro ao extrair informações das três primeiras cartas: {e}")
        return None







################################################################## Função para extrair outras informações únicas ignorando textos e capturando apenas números
def extract_other_info(element):
    other_info = []
    other_texts = element.find_elements(By.XPATH, ".//*[not(contains(@class, 'cardsymbols_block')) and not(contains(@class, 'cardsymbols_value')) and not(contains(@class, 'cardsymbols_symbol'))]")
    
    for other_text in other_texts:
        text = other_text.text.strip()
        
        # Tentar converter o texto para float (se for número)
        try:
            number_value = float(text)
            if str(number_value) not in other_info:  # Verificar se o número já foi adicionado
                other_info.append(str(number_value))  # Adicionar como string para manter consistência
        except ValueError:
            pass  # Ignorar se não for possível converter para número
    
    return other_info








######################################################################## Chama
def extract_and_print_info(driver, excel_file_path, new_sheet_name):
    try:
        # Esperar até 20 segundos para o elemento aparecer
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.gw_table_body_row.repfloptblrow.gw_table_body_row_hoverable.gw_hvr.gw_hvr_pcn.cursor-pointer'))
        )
        
        container = driver.find_element(By.CSS_SELECTOR, '.gw_table_body_row.repfloptblrow.gw_table_body_row_hoverable.gw_hvr.gw_hvr_pcn.cursor-pointer')

        # Extrair e imprimir informações das três primeiras cartas
        first_three_cards_info = extract_first_three_cards_info(driver)
        if first_three_cards_info:
            cards_info = ' - '.join(first_three_cards_info)
            print(cards_info)

        # Extrair e imprimir outras informações únicas
        other_info = extract_other_info(container)
        # other_info = [str(info) for info in other_info]  # Convertendo para string
        for other in other_info:
            print(other)

        # Salvar dados no Excel
        save_to_excel(excel_file_path, cards_info, other_info, new_sheet_name)
        return True, "Informações extraídas e salvas com sucesso."
    
    except Exception as e:
        print(f"Erro ao extrair informações: {e}")
        return False, f"Erro ao extrair informações: {e}"








##################################################################### Função para salvar no excel
def save_to_excel(excel_file_path, cards_info, other_info, new_sheet_name):
    # Verificar a extensão do arquivo
    if not excel_file_path.endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
        raise ValueError("O formato do arquivo não é suportado. Use um arquivo com extensão .xlsx, .xlsm, .xltx, ou .xltm")

    # Verificar se o arquivo existe
    if not os.path.exists(excel_file_path):
        workbook = Workbook()
        workbook.save(excel_file_path)
        sheet = workbook.active
        sheet.title = new_sheet_name
    else:
        workbook = load_workbook(excel_file_path)
        sheet = workbook.create_sheet(title=new_sheet_name)

    # Determinar a primeira linha vazia na coluna A
    first_empty_row = sheet.max_row + 1
    sheet.cell(row=first_empty_row, column=1).value = cards_info
    for idx, info in enumerate(other_info):
        sheet.cell(row=first_empty_row, column=idx + 2).value = info

    # Salvar o arquivo Excel
    workbook.save(excel_file_path)




########################################################## Caixa de diálogo
def config_handler():
    global excel_file_path

    # Limpar mensagens de erro anteriores
    label_error.config(text="")

    file_name = entry.get()  # Obter o nome do arquivo da entrada
    if not file_name:
        label_error.config(text="Insira um nome para o arquivo Excel.", fg="red")
        return

    # Diretório base onde o arquivo está localizado
    base_dir = "C:\\Users\\User\\Desktop\\"
    
    # Construir o caminho completo
    excel_file_path = f"{base_dir}{file_name}.xlsx"

    # Obter o nome da nova aba
    new_sheet_name = entry_sheet_name.get()

    if not new_sheet_name:
        label_error.config(text="Insira um nome para a nova aba.", fg="red")
        return

    # Chamada para salvar no Excel com o nome da nova aba
    success, message = extract_and_print_info(driver, excel_file_path, new_sheet_name)

    
    if success:
        label_error.config(text=message, fg="green")
    else:
        label_error.config(text=message, fg="red")


# Configurar a interface gráfica com tkinter
root = tk.Tk()
root.title("Configurações")
root.geometry("400x250")

# Adicionar um campo de entrada para o caminho do arquivo Excel
label_file_name = tk.Label(root, text="Nome do arquivo Excel:")
label_file_name.pack(pady=(20, 5))

entry = tk.Entry(root, width=30)
entry.pack(pady=5)

# Adicionar um campo de entrada para o nome da aba no Excel
label_sheet_name = tk.Label(root, text="Nome da nova aba:")
label_sheet_name.pack(pady=(20, 5))

entry_sheet_name = tk.Entry(root, width=30)
entry_sheet_name.pack(pady=5)

# Label para exibir mensagens de erro
label_error = tk.Label(root, text="", fg="red")
label_error.pack(pady=10)

# Adicionar um botão à janela do tkinter
button = tk.Button(root, text="Tudo Pronto!", command=config_handler)
button.pack(pady=5)

# Iniciar o loop do tkinter
root.mainloop()