import ttkbootstrap as tb
from ttkbootstrap.constants import *
import subprocess
import threading
import os
from PIL import Image, ImageTk  # Para manipular imagem
from tkinter import BOTH
import sys
import queue
import builtins
from tkinter.scrolledtext import ScrolledText

def executar_funcao(EXCEL_PATH, data_usuario):
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from openpyxl import Workbook, load_workbook
    from datetime import date, timedelta, datetime
    from tkinter.filedialog import asksaveasfilename
    import time
    import os
    import re
    import sys
    import tkinter as tk
    from tkinter import simpledialog

    # --- CONFIGURAÇÃO ---
    # EXCEL_PATH e data_usuario são recebidos como parâmetros (coletados na thread principal)

    USER = "emanuele@sevensuprimentos.com.br"
    PASS = "*Eas251080"

    # validação simples do data_usuario recebido
    if not (data_usuario and len(data_usuario.strip()) == 6 and data_usuario.isdigit()):
        raise ValueError("Data inválida! Use o formato DDMMAA, ex: 190825")

    # transforma input DDMMAA em string legível e em objeto date para comparações corretas
    HOJE_str = f"{data_usuario[:2]}/{data_usuario[2:4]}/{data_usuario[4:]}"
    
    def parse_date_str(s: str):
        """Tenta vários formatos e retorna datetime.date ou None."""
        for fmt in ("%d/%m/%y", "%d/%m/%Y"):
            try:
                return datetime.strptime(s.strip(), fmt).date()
            except Exception:
                continue
        return None

    HOJE = parse_date_str(HOJE_str)
    if not HOJE:
        raise ValueError("Data inválida após parse. Use DDMMAA ou DDMMAAAA.")

    ESTADOS = [
        'AC', 'AL', 'AP', 'AM', 'BA', 'CE', 'DF', 'ES', 'GO', 'MA', 'MT', 'MS', 'MG',
        'PA', 'PB', 'PR', 'PE', 'PI', 'RJ', 'RN', 'RS', 'RO', 'RR', 'SC', 'SP', 'SE', 'TO'
    ]

    # --- PREPARA PLANILHA ---
    if os.path.exists(EXCEL_PATH):
        os.remove(EXCEL_PATH)
    wb = Workbook()
    ws = wb.active
    ws.title = "Eventos"
    ws.append(["Numero do evento", "UF(VALE)", "DATA", "DESCRIÇÃO", "QTDE", "UNID. MED", "pagina de descrição"])
    wb.save(EXCEL_PATH)

    # --- INÍCIO DOS LOGS VISÍVEIS ---
    # usar a função 'log' (definida mais abaixo) para mostrar apenas mensagens de alto nível na GUI
    try:
        log("Iniciando robô...")
    except Exception:
        pass

    # --- INICIA SELENIUM ---
    try:
        log("Abrindo navegador...")
    except Exception:
        pass
    driver = webdriver.Chrome()
    wait = WebDriverWait(driver, 10)
    try:
        log("Acessando site...")
    except Exception:
        pass
    driver.get("https://vale.coupahost.com/sessions/supplier_login")

    # login
    try:
        log("Fazendo login...")
    except Exception:
        pass
    wait.until(EC.presence_of_element_located((By.ID, "user_login")))
    driver.find_element(By.ID, "user_login").send_keys(USER)
    driver.find_element(By.ID, "user_password").send_keys(PASS, Keys.RETURN)

    # Clica no elemento de data duas vezes
    try:
        time_filter = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ch_start_time"]')))
        time_filter.click()
        time.sleep(5)
        time_filter = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ch_start_time"]')))
        time_filter.click()
    except:
        pass

    # Robo irá buscar todos os casos que a data inicio = data atual, até a ENCONTRAR ONTEM
    try:
        log("Capturando dados...")
    except Exception:
        pass

    count = 1
    countBotao = 4
    encontrou_ontem = False
    while True:
        time.sleep(5)
        tbody = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="quote_request_table_tag"]')))
        print("tabela encontrada")
        linhas = tbody.find_elements(By.TAG_NAME, "tr")
        print(f"✅ Encontradas {len(linhas)} linhas na tabela.")

        # --- Buscar os números do evento ---
        for linha in linhas:
            try:
                colunas = linha.find_elements(By.TAG_NAME, "td")
                if not colunas or len(colunas) < 7:
                    continue

                # pula a linha se existir o ícone amarelo (flag_yellow) em qualquer lugar da linha
                yellow_flags = linha.find_elements(By.CSS_SELECTOR, "img[src*='flag_yellow']")
                if yellow_flags:
                    print("Pulando linha porque contém flag_yellow")
                    continue

                data_inicio_str = colunas[2].text.strip()
                data_inicio = parse_date_str(data_inicio_str)
                if data_inicio is None:
                    print(f"⚠️ Não foi possível interpretar a data '{data_inicio_str}'. Pulando linha.")
                    continue

                if data_inicio < HOJE:
                    encontrou_ontem = True
                    print(f"❌ Encontrou data anterior a HOJE ({HOJE.strftime('%d/%m/%Y')}): {data_inicio.strftime('%d/%m/%Y')}. Parando a coleta.")
                    break

                if data_inicio != HOJE:
                    print(f"⚠️ Data {data_inicio.strftime('%d/%m/%Y')} não é igual a HOJE ({HOJE.strftime('%d/%m/%Y')}). Ignorando linha.")
                    continue
                numero_evento = colunas[0].find_element(By.TAG_NAME, "a").text.strip()
                data_final = colunas[3].text.strip()
                print(f"Número do evento: {numero_evento} | Data final: {data_final}")

                ws.append([numero_evento, '', data_final, '', '', '', ''])

            except Exception as e:
                print(f"⚠️ Não foi possível extrair dados da linha: {e}")

        if encontrou_ontem:
            break

        try:
            proximo = driver.find_element(By.CLASS_NAME, "next_page")
            print("✅ Botão 'Avançar' encontrado, clicando...")
        except:
            print("⚠️ Botão 'Avançar' não encontrado.")

        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", proximo)
        time.sleep(1)

        if encontrou_ontem:
            break

        try:
            proximo.click()
            print("✅ Botão 'Avançar' clicado.")
            time.sleep(3)
        except Exception as e:
            print(f"Não tem mais páginas ou erro ao clicar no botão 'Avançar': {e}")
            break

    wb.save(EXCEL_PATH)
    try:
        log(f"💾 Planilha salva em: {EXCEL_PATH}")
    except Exception:
        pass

    # --- DETALHA CADA EVENTO ---
    try:
        log("Detalhando eventos...")
    except Exception:
        pass

    wb = load_workbook(EXCEL_PATH)
    ws = wb["Eventos"]

    for row in ws.iter_rows(min_row=2):
        evento = row[0].value
        driver.get(f"https://vale.coupahost.com/quotes/external_responses/{evento}/edit")
        wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))

        # --- VERIFICA EXISTÊNCIA DA PÁGINA DE DESCRIÇÃO ---
        try:
            botoes1 = driver.find_elements(By.XPATH, '//*[@id="pageContentWrapper"]/div[3]/div[2]/a[2]/span')
            if not botoes1:
                driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                botoes2 = driver.find_elements(By.ID, 'quote_response_submit')
                botoes2[0].click()
        except:
            row[6].value = "Erro ao verificar página de descrição"

        # Scroll e abre seção das informações
        try:
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            wait.until(EC.presence_of_element_located((By.CLASS_NAME, "s-expandLines")))
            elementos = driver.find_elements(By.CLASS_NAME, "s-expandLines")
            elementos[0].click()
        except:
            pass

        # quantidade
        try:
            quantidade = driver.find_element(By.XPATH, '//*[@id="itemsAndServicesApp"]/div/div/div[1]/div[2]/div[2]/div/form/div/div/div[2]/div/div[2]/div/p/span[1]').text
            row[4].value = quantidade
        except:
            row[4].value = 'Não foi possivel coletar a quantidade'

        # unidade
        try:
            unidade = driver.find_element(By.XPATH, '//*[@id="itemsAndServicesApp"]/div/div/div[1]/div[2]/div[2]/div/form/div/div/div[2]/div/div[2]/div/p/span[2]').text
            row[5].value = unidade
        except:
            row[5].value = 'Não foi possivel coletar a unidade'

        # descrição
        try:
            descri = driver.find_element(By.XPATH, f'//*[@id="itemsAndServicesApp"]/div/div/div[1]/div[2]/div[2]/div/form/div/div/div[1]/div/div[2]/div/p').text
            desejado = re.search(r'PT\s*\|\|\s*(.*?)\*{3,}', descri, re.DOTALL)
            if desejado:
                row[3].value = desejado.group(1).strip()
            else:
                row[3].value = descri
        except:
            pass

        # UF
        try:
            uf_text = driver.find_element(By.XPATH, f'//*[@id="itemsAndServicesApp"]/div/div/div[1]/div[2]/div[2]/div/form/div/div/div[1]/div/div[8]/div/ul/li[1]/span').text
            for sig in ESTADOS:
                if sig in uf_text:
                    row[1].value = sig
                    break
        except:
            row[1].value = 'Não foi possivel coletar a UF'

        wb.save(EXCEL_PATH)

    driver.quit()
    try:
        log("Finalizando execução.")
        log(f"Concluído! Planilha em: {EXCEL_PATH}")
    except Exception:
        # fallback para console se log não estiver disponível
        print("Concluído! Planilha em:", EXCEL_PATH)



# --- INTERFACE ---
janela = tb.Window(themename="flatly")
janela.title("Robô de Eventos - Seven")
janela.geometry("800x400")
janela.resizable

frame = tb.Frame(janela, padding=20)
frame.pack(fill=BOTH, expand=True)

# Título
titulo = tb.Label(frame, text="Robô de Eventos Seven", font=("Segoe UI", 18, "bold"))
titulo.pack(pady=(0, 20))

# Botão de iniciar
botao_iniciar = tb.Button(frame, text="Iniciar Robô", bootstyle=SUCCESS, width=30, command=executar_funcao)
botao_iniciar.pack(pady=5)

# Status
status_var = tb.StringVar(value="Aguardando início...")
status_label = tb.Label(frame, textvariable=status_var, bootstyle=INFO)
status_label.pack(pady=(20, 0))

# --- LOG (apenas status azul embaixo do botão) ---
# removido o widget ScrolledText e a fila de logs.
_original_print = builtins.print

def log(msg):
    try:
        _original_print(msg)  # continua indo ao terminal
    except Exception:
        pass
    try:
        def _set():
            status_var.set(msg if len(msg) < 200 else msg[:197] + "...")
        # agenda atualização na thread da interface
        janela.after(0, _set)
    except Exception:
        pass

# função que inicia o robô em thread e faz controle básico de interface
def _start_robot_thread():
    # desabilita botão para evitar múltiplas execuções
    botao_iniciar.config(state='disabled')
    status_var.set("Inicializando...")
    # coletar parâmetros NA THREAD PRINCIPAL (tkinter não é threadsafe)
    import tkinter as tk
    from tkinter import simpledialog
    from tkinter.filedialog import asksaveasfilename

    root = tk.Tk()
    root.withdraw()

    EXCEL_PATH = asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="Salvar planilha como"
    )
    if not EXCEL_PATH:
        status_var.set("Operação cancelada.")
        botao_iniciar.config(state='normal')
        root.destroy()
        return

    data_usuario = simpledialog.askstring(
        title="Input",
        prompt="Digite a data desejada no formato DDMMAA (ex: 190825 para 19/08/25):"
    )
    root.destroy()
    if not (data_usuario and len(data_usuario.strip()) == 6 and data_usuario.isdigit()):
        status_var.set("Data inválida. Cancelado.")
        botao_iniciar.config(state='normal')
        return

    def target():
        try:
            executar_funcao(EXCEL_PATH, data_usuario)
            print("Robo finalizado com sucesso.")
        except Exception as e:
            print(f"Erro na execução do robô: {e}")
        finally:
            def _reenable():
                botao_iniciar.config(state='normal')
                status_var.set("Pronto.")
            janela.after(0, _reenable)

    t = threading.Thread(target=target, daemon=True)
    t.start()

# atualiza o comando do botão para a versão que roda em thread
botao_iniciar.config(command=_start_robot_thread)

# não há mais _process_log_queue nem janela.after para ele

janela.mainloop()