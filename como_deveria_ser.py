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
# --- PREPARA PLANILHA ---
if os.path.exists(EXCEL_PATH):
    os.remove(EXCEL_PATH)
wb = Workbook()
ws = wb.active
ws.title = "Eventos"
ws.append(["Numero do evento", "UF(VALE)", "DATA",
          "DESCRIÇÃO", "QTDE", "UNID. MED","pagina de descrição"])
wb.save(EXCEL_PATH)

# --- INICIA SELENIUM ---
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 10)
driver.get("https://supplier.coupahost.com/sessions/new")
while True:
    try:
        driver.find_element(By.ID, "main_nav_sourcing")
        print("✅ Login detectado! Continuando...")
        break
    except NoSuchElementException:
        time.sleep(1)

# --- IR PARA LISTA DE EVENTOS ---
driver.get("https://supplier.coupahost.com/quotes/private_events/")