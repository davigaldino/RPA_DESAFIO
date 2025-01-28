from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import time

# Configurações
excel_path = r"D:\OneDrive\SANTANDER\RPA\challenge.xlsx"
edge_driver_path = r"D:\OneDrive\SANTANDER\RPA\edgedriver_win64\msedgedriver.exe"
start_button_xpath = '//button[text()="Start"]'  # XPath do botão Start

# Inicializar o WebDriver do Edge
service = Service(edge_driver_path)
driver = webdriver.Edge(service=service)

try:
    # Abrir o site do desafio
    driver.get("https://www.rpachallenge.com/")

    # Esperar até que o botão Start esteja presente na página
    wait = WebDriverWait(driver, 10)
    start_button = wait.until(EC.presence_of_element_located((By.XPATH, start_button_xpath)))
    start_button.click()

    # Ler dados do Excel
    workbook = openpyxl.load_workbook(excel_path)
    sheet = workbook.active
    excel_data = []

    for row in sheet.iter_rows(min_row=2, values_only=True):
        excel_data.append({
            "First Name": row[0],
            "Last Name": row[1],
            "Company Name": row[2],
            "Role in Company": row[3],
            "Address": row[4],
            "Email": row[5],
            "Phone Number": row[6]
        })

    # Preencher o formulário para cada item
    for item in excel_data:
        # Preencher os campos do formulário
        driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelFirstName"]').send_keys(item["First Name"])
        driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelLastName"]').send_keys(item["Last Name"])
        driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelCompanyName"]').send_keys(item["Company Name"])
        driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelRole"]').send_keys(item["Role in Company"])
        driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelAddress"]').send_keys(item["Address"])
        driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelEmail"]').send_keys(item["Email"])
        driver.find_element(By.CSS_SELECTOR, 'input[ng-reflect-name="labelPhone"]').send_keys(item["Phone Number"])

        # Clicar no botão Submit
        driver.find_element(By.CSS_SELECTOR, 'input[value="Submit"]').click()

        # Esperar um pouco para o próximo envio
        time.sleep(1)

    # Esperar o término do desafio e capturar a mensagem de sucesso
    success_message = wait.until(EC.presence_of_element_located((By.CLASS_NAME, 'congratulations')))
    print("Desafio concluído com sucesso!")
    print(success_message.text)

    # Manter o navegador aberto após a execução
    print("O navegador permanecerá aberto. Você pode fechá-lo manualmente.")
    input("Pressione Enter para fechar o script...")  # Aguarda entrada do usuário para encerrar o script

except Exception as e:
    print(f"Ocorreu um erro: {e}")
    input("Pressione Enter para fechar o script...")  # Aguarda entrada do usuário em caso de erro

finally:
    # Removido o driver.quit() para manter o navegador aberto
    pass