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







def exctract_other_info():
    try:
        # Encontrar apenas a primeira div pai que contém as informações desejadas
        div_pai = driver.find_element(By.XPATH, "(//div[@class='gw_table_body_row repfloptblrow gw_table_body_row_hoverable gw_hvr gw_hvr_pcn cursor-pointer'])[1]")

        # Encontrar todas as divs filhas com as classes específicas dentro da div pai encontrada
        divs_filhas_position = div_pai.find_elements(By.XPATH, ".//div[@class='position-absolute w-100 f-center']")
        divs_filhas_text_center = div_pai.find_elements(By.XPATH, ".//div[@class='text-center']")

        for div_filha in divs_filhas_position + divs_filhas_text_center:
            # Obter o texto da div filha
            texto = div_filha.text.strip()
            print("Informação encontrada:", texto)
    
    except Exception as e:
        print(f"Erro ao extrair informações: {str(e)}")









        

# Configurar a interface gráfica com tkinter
root = tk.Tk()
root.title("Configurações")
root.geometry("400x250")

button = tk.Button(root, text="Tudo Pronto!", command=exctract_other_info)
button.pack(pady=5)

# Iniciar o loop principal do tkinter
root.mainloop()