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
import tkinter as tk
from tkinter import simpledialog
# --- CONFIGURAÇÕES DE USUÁRIO ---
USER = "emanuele@sevensuprimentos.com.br"
PASS = "*Eas251080"

# Iput de data
root = tk.Tk()
root.withdraw()
data_usuario = simpledialog.askstring(
    title="Input",
    prompt="Digite a data desejada no formato DDMMAA (ex: 190825 para 19/08/25):"
)
if data_usuario and len(data_usuario.strip()) == 6 and data_usuario.isdigit():
    HOJE = f"{data_usuario[:2]}/{data_usuario[2:4]}/{data_usuario[4:]}"
else:
    raise ValueError("Data inválida! Use o formato DDMMAA, ex: 190825")
root.destroy()

#HOJE = (date.today() - timedelta(days=11)).strftime("%d/%m/%y")
#ONTEM = (date.today() - timedelta(days=12)).strftime("%d/%m/%y")

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

time.sleep(3)
encontrou_ontem = False
# --- Buscar a tabela ---
while True:
    try:
        tbody = driver.find_element(By.XPATH, '//*[@id="dataTableSourcingView"]/tbody')
        print ("tabela encontrada")
    except:
        print ("tabela não encontrada")
        continue
# --- Buscar a linhas ---
    try:
        linhas = tbody.find_elements(By.TAG_NAME, "tr")
        print(f"✅ Encontradas {len(linhas)} linhas na tabela.")
    except:
        print ("linhas não encontradas")
        continue
    # --- Buscar a numero do ---
    for linha in linhas:
     try:
        colunas = linha.find_elements(By.TAG_NAME, "td")
        if not colunas or len(colunas) < 6:
            continue

        # data inicio (quinto <td>, dentro de <a>)
        data_inicio = colunas[4].text.strip()

        if data_inicio < HOJE:
            encontrou_ontem = True
            print(f"❌ Encontrou data anterior a HOJE ({HOJE}): {data_inicio}. Parando a coleta.")
            break
        

        if data_inicio != HOJE:
            print(f"⚠️ Data {data_inicio} não é igual a HOJE ({HOJE}). Ignorando linha.")
            continue

        # Número do evento (primeiro <td>, dentro de <a>)
        numero_evento = colunas[0].find_element(By.TAG_NAME, "a").text.strip()

        # Data final do evento (sexto <td>)
        data_final = colunas[5].text.strip()

        print(f"Número do evento: {numero_evento} | Data final: {data_final}")

        # Salva na planilha nas colunas corretas
        ws.append([numero_evento, '', data_final, '', '', '', ''])

     except Exception as e:
        print(f"⚠️ Não foi possível extrair dados da linha: {e}")
    try:
        botao_avancar = driver.find_element(By.XPATH, '//button[.//span[text()="Avançar"] and not(@disabled)]')
        print("✅ Botão 'Avançar' encontrado, clicando...")
    except:
        print("⚠️ Botão 'Avançar' não encontrado.")

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_avancar)
    time.sleep(1)  # Aguarda animações ou overlays sumirem
    # Se encontrou data anterior, para tudo
    if encontrou_ontem:
        break
    try:
        botao_avancar.click()
        print("✅ Botão 'Avançar' clicado.")
        time.sleep(3)  # Aguarda a próxima página carregar
    except Exception as e:
        print(f"Não tem mais páginas ou erro ao clicar no botão 'Avançar': {e}")
        break
    
# Salva a planilha ao final
wb.save(EXCEL_PATH)
print(f"💾 Planilha salva em: {EXCEL_PATH}")

# --- DETALHA CADA EVENTO ---
