from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import tkinter as tk
import openpyxl
from openpyxl import Workbook



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




# Função para extrair informações das três primeiras cartas
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


# Função para extrair outras informações únicas ignorando textos e capturando apenas números
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




def extract_and_print_info(driver, excel_file_path):
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
        save_to_excel(excel_file_path, cards_info, other_info)
    
    except Exception as e:
        print(f"Erro ao extrair informações: {e}")

def save_to_excel(excel_file_path, cards_info, other_info):
    # Verificar a extensão do arquivo
    if not excel_file_path.endswith(('.xlsx', '.xlsm', '.xltx', '.xltm')):
        raise ValueError("O formato do arquivo não é suportado. Use um arquivo com extensão .xlsx, .xlsm, .xltx, ou .xltm")

    # Abrir o arquivo Excel ou criar um novo se não existir
    try:
        workbook = openpyxl.load_workbook(excel_file_path)
    except FileNotFoundError:
        workbook = Workbook()

    sheet = workbook.active

    # Determinar a primeira linha vazia na coluna A
    first_empty_row = sheet.max_row + 1
    sheet.cell(row=first_empty_row, column=1).value = cards_info
    for idx, info in enumerate(other_info):
        sheet.cell(row=first_empty_row, column=idx + 2).value = info


    # Salvar o arquivo Excel
    workbook.save(excel_file_path)






def btn_handler():
    global excel_file_path
    file_name = entry.get()  # Obter o nome do arquivo da entrada
    if file_name:
        # Diretório base onde o arquivo está localizado
        base_dir = "C:\\Users\\User\\Desktop\\"
        
        # Construir o caminho completo
        excel_file_path = f"{base_dir}{file_name}.xlsx"
        print(excel_file_path)

        extract_and_print_info(driver, excel_file_path)
        # Fechar a janela
        root.destroy()

# Configurar a interface gráfica com tkinter
root = tk.Tk()
root.title("Configurações")
root.geometry("300x150")

# Adicionar um campo de entrada para o caminho do arquivo Excel
label = tk.Label(root, text="Digite o nome do arquivo Excel (sem extensão):")
label.pack(pady=10)

entry = tk.Entry(root, width=30)
entry.pack(pady=10)

# Adicionar um botão à janela do tkinter
button = tk.Button(root, text="Tudo Pronto!", command=btn_handler)
button.pack(pady=20)

# Iniciar o loop do tkinter
root.mainloop()

