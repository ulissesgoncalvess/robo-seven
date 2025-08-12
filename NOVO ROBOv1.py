from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from openpyxl import Workbook, load_workbook
from datetime import date, timedelta
from tkinter.filedialog import asksaveasfilename
import time
import os
import re

USER = "emanuele@sevensuprimentos.com.br"
PASS = "*Eas251080"
#HOJE = date.today().strftime("%d/%m/%y")  # Data que será filtrada
HOJE =(date.today() - timedelta(days=3)).strftime("%d/%m/%y")
ONTEM = (date.today() - timedelta(days=4)).strftime("%d/%m/%y")

ESTADOS = ['AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG',
           'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO']

# --- CONFIGURAÇÃO ---
EXCEL_PATH = asksaveasfilename(defaultextension=".xlsx",
 filetypes=[("Excel files", "*.xlsx")],
 title="Salvar planilha como")
# --- PREPARA PLANILHA ---
if os.path.exists(EXCEL_PATH):
    os.remove(EXCEL_PATH)
wb = Workbook()
ws = wb.active
ws.title = "Eventos"
ws.append(["Numero do evento", "UF(VALE)", "DATA", "DESCRIÇÃO", "QTDE", "UNID. MED", "pagina de descrição"])
wb.save(EXCEL_PATH)

# --- INICIA SELENIUM ---
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 60)
driver.get("https://supplier.coupahost.com/sessions/new")

# Login manual (aguarda menu)
while True:
    try:
        driver.find_element(By.ID, "main_nav_sourcing")
        print("Elemento apareceu! Continuando...")
        break
    except NoSuchElementException:
        time.sleep(1)

# Vai direto para a lista de eventos
driver.get("https://supplier.coupahost.com/quotes/private_events/")
time.sleep(5)
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
# --- COLETA EVENTOS POR DATA ---
while True:
    try:
       botao90 = driver.find_element(By.CSS_SELECTOR,'#header-container > div > div.coupaUI > div > div.coupaTable__outerContainer > div.coupaTableWithSidePanel__tablePanel.s-coupaTableWithSidePanel__tablePanel > div > div.coupaTable__pagination > div.rowsPerPage.s-rowsPerPage > button:nth-child(6)')
       break
    except:
     print('Naõ encontrado')

while True:
    try:
        botao90.click
        break
    except:
        print("Botão não clicado")
