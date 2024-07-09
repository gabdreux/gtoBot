from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import re



geckodriver_path = r'C:\Users\User\Desktop\RonaldoBot\geckodriver.exe'
# service = Service(r'C:\Users\User\Desktop\botRafa\geckodriver.exe')


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

driver.get(url)






# Esperar um pouco para garantir que a página carregue completamente
time.sleep(60)





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




# Função principal para extrair e imprimir todas as informações
def extract_and_print_info(url):
    try:
        # Esperar até 20 segundos para o elemento aparecer
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.gw_table_body_row.repfloptblrow.gw_table_body_row_hoverable.gw_hvr.gw_hvr_pcn.cursor-pointer'))
        )
        
        container = driver.find_element(By.CSS_SELECTOR, '.gw_table_body_row.repfloptblrow.gw_table_body_row_hoverable.gw_hvr.gw_hvr_pcn.cursor-pointer')

        # Extrair e imprimir informações das três primeiras cartas
        first_three_cards_info = extract_first_three_cards_info(driver)
        if first_three_cards_info:
            for card in first_three_cards_info:
                print(card)

        # Extrair e imprimir outras informações únicas
        other_info = extract_other_info(container)
        for other in other_info:
            print(other)
    
    except Exception as e:
        print(f"Erro ao extrair informações: {e}")

# Chamada da função principal
extract_and_print_info(url)