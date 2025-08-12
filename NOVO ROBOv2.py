from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from openpyxl import Workbook, load_workbook
from datetime import date, timedelta
from tkinter.filedialog import asksaveasfilename
import time
import os

# --- CONFIGURAÇÕES DE USUÁRIO ---
USER = "emanuele@sevensuprimentos.com.br"
PASS = "*Eas251080"
HOJE = (date.today() - timedelta(days=3)).strftime("%d/%m/%y")
ONTEM = (date.today() - timedelta(days=4)).strftime("%d/%m/%y")

ESTADOS = [
    'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG',
    'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
]

# --- SALVAR PLANILHA ---
EXCEL_PATH = asksaveasfilename(defaultextension=".xlsx",
    filetypes=[("Excel files", "*.xlsx")],
    title="Salvar planilha como")

if os.path.exists(EXCEL_PATH):
    os.remove(EXCEL_PATH)

wb = Workbook()
ws = wb.active
ws.title = "Eventos"
ws.append(["Numero do evento", "UF(VALE)", "DATA", "DESCRIÇÃO", "QTDE", "UNID. MED", "pagina de descrição"])
wb.save(EXCEL_PATH)

# --- CONFIGURAÇÃO DO CHROME ---
options = webdriver.ChromeOptions()
options.add_argument("--disable-extensions")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--start-maximized")

driver = webdriver.Chrome(options=options)
wait = WebDriverWait(driver, 30)

# --- LOGIN MANUAL ---
driver.get("https://supplier.coupahost.com/sessions/new")
print("⚠ Faça login manual e clique em 'Sourcing'.")

while True:
    try:
        driver.find_element(By.ID, "main_nav_sourcing")
        print("✅ Login detectado! Continuando...")
        break
    except NoSuchElementException:
        time.sleep(1)

# --- IR PARA LISTA DE EVENTOS ---
driver.get("https://supplier.coupahost.com/quotes/private_events/")

# --- BOTÃO 90 ----
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
while True:
 try:
   bot_90 = driver.find_element(By.CSS_SELECTOR, 'button[aria-label="Results per page 90"]')
   # bot_90.click
   print(bot_90.id)
   print(bot_90.size)
   print(bot_90.text)
   bot_90.click
   bot_90.click(1)
   break
 except Exception as e:
   print ('BOTÃO NAO ENCONTRADO')
time.sleep(3)
while True:
 # Robo irá buscar todos os casos que a data inicio 
 tbody = driver.find_element(By.XPATH, '/html/body/div[2]/div[3]/div/div[2]/div[1]/div/div[2]/div/div[2]/div[2]/div/div[2]/table/tbody"]')
 # Pega todas as linhas dentro do tbody
 linhas = tbody.find_elements(By.TAG_NAME, "tr")
 for linha in linhas:
        try:
            colunas = linha.find_elements(By.TAG_NAME, "td")
            dados = [coluna.text for coluna in colunas]
            if dados[2] != HOJE:
                continue
            numero_do_evento = dados[0]
            dataFinal = dados[3]
            vazio = ''
            ws.append([numero_do_evento, vazio, dataFinal])
        except:
            pass

# Salva as alterações
wb.save(arquivo_excel)


