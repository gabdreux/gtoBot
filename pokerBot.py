from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
import os
from openpyxl import Workbook, load_workbook
import time
import re
from collections import Counter
import pyautogui


# Caminhos para o geckodriver e perfil do Firefox
geckodriver_path = r'C:\Users\User\Desktop\RonaldoBot\geckodriver.exe'
profile_path = r'C:\Users\User\AppData\Roaming\Mozilla\Firefox\Profiles\ddxg8c5z.default-release-1'

# Configurar o serviço e perfil do Firefox
service = Service(geckodriver_path)
options = Options()
options.profile = profile_path

driver = webdriver.Firefox(service=service, options=options)

url = 'https://app.gtowizard.com/'
driver.get(url)





# Função para extrair informações das três primeiras cartas
def extract_first_three_cards_info(container):
    card_info = []
    try:
        card_blocks = container.find_elements(By.CSS_SELECTOR, ".cardsymbols_block")[:3]
        for card_block in card_blocks:
            card_value = card_block.find_element(By.CSS_SELECTOR, ".cardsymbols_value").text.strip()
            card_symbol_element = card_block.find_element(By.CSS_SELECTOR, ".cardsymbols_symbol svg")
            svg_path = card_symbol_element.find_element(By.CSS_SELECTOR, 'path').get_attribute('d')
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




# Função para salvar no excel
def save_to_excel(workbook, sheet_name, cards_info, other_info):
    if sheet_name not in workbook.sheetnames:
        sheet = workbook.create_sheet(title=sheet_name)
    else:
        sheet = workbook[sheet_name]
    first_empty_row = sheet.max_row + 1
    sheet.cell(row=first_empty_row, column=1).value = cards_info
    for idx, info in enumerate(other_info):
        sheet.cell(row=first_empty_row, column=idx + 2).value = info
    workbook.save(excel_file_path)






# Atualização da função extract_and_save_all_info para usar índice dinâmico
def extract_and_save_all_info(driver, excel_file_path, sheet_name):
    try:
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.gw_table_body_row.repfloptblrow.gw_table_body_row_hoverable.gw_hvr.gw_hvr_pcn.cursor-pointer'))
        )
        containers = driver.find_elements(By.CSS_SELECTOR, '.gw_table_body_row.repfloptblrow.gw_table_body_row_hoverable.gw_hvr.gw_hvr_pcn.cursor-pointer')

        print(f"Número de containers encontrados: {len(containers)}")

        if not os.path.exists(excel_file_path):
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = sheet_name
            workbook.save(excel_file_path)
        else:
            workbook = load_workbook(excel_file_path)

        
        # Contador para controlar as iterações
        iteration_count = 0 
        

        for i, container in enumerate(containers, start=1):
            iteration_count += 1
            print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" + str(iteration_count) )   
            # Construir o XPath dinâmico com base no índice i
            xpath = f"(//div[@class='gw_table_body_row repfloptblrow gw_table_body_row_hoverable gw_hvr gw_hvr_pcn cursor-pointer'])[{i}]"
            div_pai = driver.find_element(By.XPATH, xpath)

            # Função para extrair outras informações únicas ignorando textos e capturando apenas números
            def extract_other_info():
                try:
                    # Encontrar todas as divs filhas com as classes específicas dentro da div pai encontrada
                    divs_filhas_position = div_pai.find_elements(By.XPATH, ".//div[@class='position-absolute w-100 f-center']")
                    divs_filhas_text_center = div_pai.find_elements(By.XPATH, ".//div[@class='text-center']")

                    other_info = []
                    for div_filha in divs_filhas_position + divs_filhas_text_center:
                        # Obter o texto da div filha
                        texto = div_filha.text.strip()
                        other_info.append(texto)

                    return other_info

                except Exception as e:
                    print(f"Erro ao extrair informações: {str(e)}")
                    return []

            first_three_cards_info = extract_first_three_cards_info(container)
            if first_three_cards_info:
                cards_info = ' - '.join(first_three_cards_info)
                print(cards_info)
            other_info = extract_other_info()
            for other in other_info:
                print(other)

                
            save_to_excel(workbook, sheet_name, cards_info, other_info)



        if iteration_count % 105 == 0:
            time.sleep(5)
            print("OOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOPPPPPPPPPPPPPPAAAAAAAAAAAAAAAA")
            for _ in range(12):  # ajuste o número de vezes conforme necessário
                pyautogui.scroll(-100)
            time.sleep(3)
            print("CAPTURANDO NOVAMENTE!!!!!!!!!!")
            btn_call()    



        return True, "Informações extraídas e salvas com sucesso."
    
    except Exception as e:
        print(f"Erro ao extrair informações: {e}")
        return False, f"Erro ao extrair informações: {e}"





def btn_call():
    time.sleep(5)
    print("INICIANDO SEQUENCIA!")
    global excel_file_path
    label_error.config(text="")
    file_name = entry.get()
    if not file_name:
        label_error.config(text="Insira um nome para o arquivo Excel.", fg="red")
        return
    base_dir = "C:\\Users\\User\\Desktop\\"
    excel_file_path = f"{base_dir}{file_name}.xlsx"
    sheet_name = entry_sheet_name.get()
    if not sheet_name:
        label_error.config(text="Insira um nome para a nova aba.", fg="red")
        return
    success, message = extract_and_save_all_info(driver, excel_file_path, sheet_name)
    if success:
        label_error.config(text=message, fg="green")
    else:
        label_error.config(text=message, fg="red")




# Caixa de diálogo
def config_handler():
    time.sleep(5)
    for _ in range(1):  # ajuste o número de vezes conforme necessário
        pyautogui.scroll(-100)
    time.sleep(3)    
    btn_call()





# Configurar a interface gráfica com tkinter
root = tk.Tk()
root.title("Configurações")
root.geometry("400x250")

label_file_name = tk.Label(root, text="Nome do arquivo Excel:")
label_file_name.pack(pady=(20, 5))

entry = tk.Entry(root, width=30)
entry.pack(pady=5)

label_sheet_name = tk.Label(root, text="Nome da nova aba:")
label_sheet_name.pack(pady=(20, 5))

entry_sheet_name = tk.Entry(root, width=30)
entry_sheet_name.pack(pady=5)

label_error = tk.Label(root, text="", fg="red")
label_error.pack(pady=10)

button = tk.Button(root, text="Tudo Pronto!", command=config_handler)
button.pack(pady=5)

root.mainloop()
