from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
from datetime import date
import tkinter as tk
from tkinter.filedialog import askopenfilename
import time
import pyperclip
import pyautogui
import subprocess
from tkinter import messagebox
import threading

# Função única para rodar o WebDriver em uma thread separada
def rodar_webdriver():
    def rodar_webdriver():
        comando = r'"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\edge_selenium"'
        subprocess.run(comando, shell=True)
    
    
    thread = threading.Thread(target=rodar_webdriver)
    thread.start()

# O código principal do seu app continua aqui
print("WebDriver iniciado em uma thread separada. O aplicativo não irá travar.")


def criar_nf_thuban():

    def ler_excel():
        explorador_arquivos = tk.Tk()
        explorador_arquivos.withdraw()
        caminho_planilha = askopenfilename(filetypes=[("Excel files", "*.xlsm;*.xlsx")])
        return caminho_planilha

    caminho_plnailha = ler_excel()
    df = pd.read_excel(caminho_plnailha)

    def index_linhas():
        arquivo = 'numero_inicial_criar_nf.txt'
        qtde_nf = df.iloc[5, 5]

        for _ in range(int(qtde_nf)):

            for _ in range(1):
                try:
                    with open(arquivo, 'r') as f:
                        conteudo = f.read().strip()
                        if conteudo:
                            numero_inicial = int(conteudo)
                        else:
                            numero_inicial = 7
                except FileNotFoundError:
                    numero_inicial = 7  
                segundo_numero = 1
                resultado = numero_inicial + segundo_numero
                numero_inicial += 1
                print(resultado)
                with open(arquivo, 'w') as f:
                    f.write(str(numero_inicial))

            tipo_nf = df.iloc[int(resultado), 2]
            nro_nf = df.iloc[int(resultado), 3]
            cnpj = df.iloc[int(resultado), 4]
            ins_estadual = '00000000'
            ins_municipal = df.iloc[int(resultado), 6]
            data_emissao = df.iloc[int(resultado), 7]
            data_vencimento = df.iloc[int(resultado), 9]
            descricao_subir_nf = df.iloc[int(resultado), 10]
            valor_nf = df.iloc[int(resultado), 11]
            cod_boleto = df.iloc[int(resultado), 12]
            chave_pix = df.iloc[int(resultado), 17]
            nro_banco = df.iloc[int(resultado), 18]
            agencia = df.iloc[int(resultado), 19]
            conta_corrente = df.iloc[int(resultado), 20]
            
            tipo_pg = df.iloc[int(resultado), 15]

            x = 269
            y = 251

            time.sleep(2)
            pyautogui.click(515, 264)

            options = webdriver.EdgeOptions()
            options.debugger_address = "localhost:9222"
            driver = webdriver.Edge(options=options)
            driver.get("https://proveedores-hh.com.ar/thuban-web/secure/main.zul#mainScreen")

            wait = WebDriverWait(driver, 30)
            abas = driver.window_handles
            driver.switch_to.window(abas[-1])
            driver.switch_to.window(driver.current_window_handle)
            driver.maximize_window()

            time.sleep(5)

            try:         
                clicar = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div/div/div[1]/div[1]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/table/tbody/tr/td[2]/table/tbody/tr/td[1]"
                clicar_elemento = wait.until(EC.element_to_be_clickable((By.XPATH, clicar)))
                clicar_elemento.click()
                print('Elemento encontrado e clicado com sucesso!')
            except Exception as e:
                print(f'Erro ao encontrar o elemento: {e}')
            time.sleep(1)
            for _ in range(1):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.press('down')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(0.25)
            pyperclip.copy(tipo_nf)
            time.sleep(0.5)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(0.25)
            pyperclip.copy(nro_nf)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(0.75)
            for _ in range(5):
                pyautogui.press('enter')
            time.sleep(0.25)
            pyperclip.copy(cnpj)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            for _ in range(2):
                pyautogui.press('tab')
            time.sleep(0.75)
            for _ in range(5):
                pyautogui.press('enter')
            time.sleep(0.25)
            pyperclip.copy(ins_estadual)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(1)
            for _ in range(5):
                pyautogui.press('enter')
            time.sleep(0.25)
            pyperclip.copy(ins_municipal)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(1)
            for _ in range(5):
                pyautogui.press('enter')
            time.sleep(0.25)
            pyperclip.copy(data_emissao)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            for _ in range(9):
                pyautogui.press('backspace')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(0.25)
            pyperclip.copy(data_vencimento)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            for _ in range(9):
                pyautogui.press('backspace')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(0.25)
            pyperclip.copy(descricao_subir_nf)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            for _ in range(2):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyperclip.copy(valor_nf)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            if tipo_pg == "Pix":
                for _ in range(3):
                    pyautogui.press('tab')
                time.sleep(0.25)
                pyperclip.copy(chave_pix)
                time.sleep(0.25)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.25)
            elif tipo_pg == "Boleto":
                for _ in range(2):
                    pyautogui.press('tab')
                time.sleep(0.25)
                pyperclip.copy(cod_boleto)
                time.sleep(0.25)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.25)
            elif tipo_pg == "Dados Bancarios":
                for _ in range(4):
                    pyautogui.press('tab')
                time.sleep(0.25)
                pyperclip.copy(nro_banco)
                time.sleep(0.25)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.25)
                pyautogui.press('tab')
                time.sleep(0.25)
                pyperclip.copy(agencia)
                time.sleep(0.25)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.25)
                pyautogui.press('tab')
                time.sleep(0.25)
                pyperclip.copy(conta_corrente)
                time.sleep(0.25)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.25)
            else:
                print("Tipo de Pagamento Não Localizado!")
            time.sleep(0.25)
            try: 
                clicar = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div/div/div[1]/div[1]/div/div/div[2]/div/div/div/div/div/div/div[2]/div[1]/div[2]/div/div/div/table/tbody/tr/td[2]/table/tbody/tr/td[1]"
                clicar_elemento = wait.until(EC.element_to_be_clickable((By.XPATH, clicar)))
                clicar_elemento.click()
                print('Elemento encontrado e clicado com sucesso!')
            except Exception as e:
                print(f'Erro ao encontrar o elemento: {e}')
            time.sleep(5)
            for _ in range(19):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyperclip.copy(valor_nf)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            for _ in range(8):
                pyautogui.press('tab')
            time.sleep(0.25)
            for _ in range(2):
                pyautogui.press('down')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(0.25)
            for _ in range(2):
                pyautogui.press('down')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(0.25)
            for _ in range(2):
                pyautogui.press('down')
            time.sleep(0.25)
            for _ in range(2):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.press('enter')
            time.sleep(1)
            pyautogui.press('enter')
            time.sleep(1)
            for _ in range(6):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.press('enter')
            time.sleep(0.25)
            pyautogui.write(rf"L:\FC\_CONTAS A PAGAR\_THUBAN")
            time.sleep(0.25)
            pyautogui.press('enter')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.write(f'{nro_nf} {cnpj}.pdf')
            time.sleep(5)
            for _ in range(3):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.press('up')
            time.sleep(0.25)
            pyautogui.press('enter')
            for _ in range(2):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.press('enter')
            time.sleep(5)
            pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.press('enter')
            time.sleep(3)
            pyautogui.press('enter')
            time.sleep(2)

    index_linhas()
    
    messagebox.showinfo("Status", "Notas Criadas")

def contabilizar_nf_thuban():

    def ler_excel():
        explorador_arquivos = tk.Tk()
        explorador_arquivos.withdraw()
        caminho_planilha = askopenfilename(filetypes=[("Excel files", "*.xlsm;*.xlsx")])
        return caminho_planilha

    caminho_plnailha = ler_excel()
    df = pd.read_excel(caminho_plnailha)

    arquivo = 'numero_inicial_contabilizar_nf.txt'
    valor_rateio_txt = 'valor_rateio_txt.txt'
    conta_contabil_rateio_txt = 'conta_contabio_rateio_txt.txt'
    centro_custo_rateio_txt = 'centro_custo_rateio_txt.txt'
    descricao_rateio_txt = 'descricao_rateio_txt.txt'
    qtde_nf = df.iloc[5, 5]

    def contabilizar_nf_thuban_rpa():

        for _ in range(int(qtde_nf)):

            for _ in range(1):
                try:
                    with open(arquivo, 'r') as f:
                        conteudo = f.read().strip()
                        if conteudo:
                            numero_inicial = int(conteudo)
                        else:
                            numero_inicial = 7
                except FileNotFoundError:
                    numero_inicial = 7  
                segundo_numero = 1
                resultado = numero_inicial + segundo_numero
                numero_inicial += 1
                print(resultado)
                with open(arquivo, 'w') as f:
                    f.write(str(numero_inicial))

            tipo_nf = df.iloc[int(resultado), 2]
            nro_nf = df.iloc[int(resultado), 3]
            cnpj = df.iloc[int(resultado), 4]
            ins_estadual = '00000000'
            ins_municipal = df.iloc[int(resultado), 6]
            data_emissao = df.iloc[int(resultado), 7]
            data_vencimento = df.iloc[int(resultado), 9]
            descricao_contabilizar_nf = df.iloc[int(resultado), 21]
            valor_nf = df.iloc[int(resultado), 11]
            cod_boleto = df.iloc[int(resultado), 12]
            conta_contabil = df.iloc[int(resultado), 13]
            centro_custo = df.iloc[int(resultado), 14]
            chave_pix = df.iloc[int(resultado), 17]
            nro_banco = df.iloc[int(resultado), 18]
            agencia = df.iloc[int(resultado), 19]
            conta_corrente = df.iloc[int(resultado), 20]
            tipo_pg = df.iloc[int(resultado), 15]
            data_hoje = date.today()
            status = 'Enviado para SSO'
            qtde_rateio = df.iloc[int(resultado), 62]

            time.sleep(2)
            pyautogui.click(515, 264)

            options = webdriver.EdgeOptions()
            options.debugger_address = "localhost:9222"
            driver = webdriver.Edge(options=options)

            driver.get("https://proveedores-hh.com.ar/thuban-web/secure/main.zul#mainScreen")
            wait = WebDriverWait(driver, 30)

            abas = driver.window_handles
            driver.switch_to.window(abas[-1])
            driver.switch_to.window(driver.current_window_handle)
            driver.maximize_window()

            time.sleep(0.25)
            try: 
                clicar = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div/div/div[1]/div[1]/div/div/div[1]/div[1]/div/div"
                clicar_elemento = wait.until(EC.element_to_be_clickable((By.XPATH, clicar)))
                clicar_elemento.click()
                print('Elemento encontrado e clicado com sucesso!')
            except Exception as e:
                print(f'Erro ao encontrar o elemento: {e}')
            time.sleep(5)
            for _ in range(4):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.press('down')
            time.sleep(0.25)
            for _ in range(2):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyperclip.copy(tipo_nf)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            for _ in range(2):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyperclip.copy(nro_nf)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            for _ in range(12):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyperclip.copy(data_hoje)
            time.sleep(0.25)
            pyautogui.hotkey('ctrl', 'v')
            time.sleep(0.25)
            for _ in range(16):
                pyautogui.press('tab')
            time.sleep(0.25)
            pyautogui.press('enter')
            time.sleep(1)
       
            baixar_notas_xpath = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div/div/div[2]/div/div/div/div[1]/div[1]/div/div/div[2]/div/div/div/table[3]/tbody/tr/td/div/div/div[2]/table/tbody[2]/tr/td[3]/div/a"
            elemento_nfs_e = wait.until(EC.element_to_be_clickable((By.XPATH, baixar_notas_xpath)))

            # Criando uma instância de ActionChains
            actions = ActionChains(driver)

            # Dando um duplo clique no elemento
            actions.double_click(elemento_nfs_e).perform()

            # Aguarda um pouco para que a nova aba abra e se estabilize
            time.sleep(5)

            # Troca para a nova aba se necessário
            abas = driver.window_handles
            driver.switch_to.window(abas[-1])

            time.sleep(0.25)
            pyautogui.press('F11')
            time.sleep(0.75)

            # Verifica se a nova aba está carregada e visível
            dynamic_links_xpath = "/html/body/div/div/div/div/div[1]/div[1]/div[2]/div/div/div[1]/div[3]/ul/li[2]/div/div/div/span"
            try:
                exe_dynamic_links = wait.until(EC.presence_of_element_located((By.XPATH, dynamic_links_xpath)))
                exe_dynamic_links.click()
                print("Elemento encontrado e clicado com sucesso.")
            except Exception as e:
                print(f"Erro ao encontrar o elemento: {e}")

            time.sleep(5)

            if int(qtde_rateio) == 0:
                for _ in range(2):
                    pyautogui.press('tab')
                time.sleep(0.25)
                pyautogui.press('enter')
                time.sleep(0.25)
                pyautogui.press('down')
                time.sleep(0.25)
                pyautogui.press('enter')
                time.sleep(3)
                for _ in range(25):
                    pyautogui.press('tab')
                time.sleep(0.25)
                pyperclip.copy(valor_nf)
                time.sleep(0.25)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.25)
                pyautogui.press('tab')
                time.sleep(0.25)
                pyperclip.copy(conta_contabil)
                time.sleep(0.25)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.25)
                pyautogui.press('tab')
                time.sleep(0.25)
                if centro_custo == 0:
                    pyautogui.write("")
                else:
                    pyperclip.copy(centro_custo)
                    time.sleep(0.25)
                    pyautogui.hotkey('ctrl', 'v')
                    time.sleep(0.25)
                time.sleep(0.25)
                pyautogui.press('tab')
                pyperclip.copy(descricao_contabilizar_nf)
                time.sleep(0.25)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.25)

                criar_cont_xpath = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[1]/div/div/div[3]/div[1]/div/table/tbody/tr[17]/td/fieldset/div/span[1]/table/tbody/tr[2]/td[2]"
                try:
                    exe_criar_cont = wait.until(EC.presence_of_element_located((By.XPATH, criar_cont_xpath)))
                    exe_criar_cont.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")

                time.sleep(5)
                pyautogui.press('enter')
                time.sleep(0.25)
                                    
                fields_cont_xpath = "/html/body/div[1]/div/div/div/div[1]/div[1]/div[2]/div/div/div[1]/div[3]/ul/li[1]/div/div/div/span"
                try:
                    exe_fields_cont = wait.until(EC.presence_of_element_located((By.XPATH, fields_cont_xpath)))
                    exe_fields_cont.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")

                cadiado_xpath = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div[1]/div[1]/div[2]/table/tbody/tr[1]/td/div/div/div/div/fieldset/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/span/div/img"
                try:
                    exe_cadiado = wait.until(EC.presence_of_element_located((By.XPATH, cadiado_xpath)))
                    exe_cadiado.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")

                time.sleep(2)
                try:
                    exe_fields_cont = wait.until(EC.presence_of_element_located((By.XPATH, fields_cont_xpath)))
                    exe_fields_cont.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")
                time.sleep(0.25)
                for _ in range(14):
                    pyautogui.press('tab')
                time.sleep(0.25)
                for _ in range(2):
                    pyautogui.press('up')
                time.sleep(0.25)
                pyautogui.press('tab')
                time.sleep(5)
                                    
                salvar_cont_xpath = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div[1]/div[1]/div[2]/table/tbody/tr[1]/td/div/div/div/div/fieldset/div/table/tbody/tr/td[1]/table/tbody/tr/td[15]/span/div/img"
                try:
                    exe_salvar_cont = wait.until(EC.presence_of_element_located((By.XPATH, salvar_cont_xpath)))
                    exe_salvar_cont.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")

                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(2)
                pyautogui.press('F11')
                time.sleep(0.25)
                pyautogui.hotkey('alt', 'F4')
                time.sleep(1)

                driver.switch_to.window(abas[0])

                clear_xpath = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div/div/div[1]/div[1]/div/div/div[3]/div[1]/div/fieldset/div/span[1]/table/tbody/tr[2]/td[2]"
                try:
                    exe_clear = wait.until(EC.presence_of_element_located((By.XPATH, clear_xpath)))
                    exe_clear.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")

                time.sleep(2)
            
            elif int(qtde_rateio) > 1:
                try:
                    for _ in range(int(qtde_rateio)):
                        for _ in range(1):
                            try:
                                with open(valor_rateio_txt, 'r') as f:
                                    conteudo_valor_rateio = f.read().strip()
                                    if conteudo_valor_rateio:
                                        numero_inicial_valor_rateio = int(conteudo_valor_rateio)
                                    else:
                                        numero_inicial_valor_rateio = 18
                            except FileNotFoundError:
                                numero_inicial_valor_rateio = 18
                            segundo_numero_valor_rateio = 4
                            resultado_valor_rateio = numero_inicial_valor_rateio + segundo_numero_valor_rateio
                            numero_inicial_valor_rateio += 4
                            print(resultado_valor_rateio)
                            with open(valor_rateio_txt, 'w') as f:
                                f.write(str(numero_inicial_valor_rateio))

                        vlr_rateio = df.iloc[int(resultado), int(resultado_valor_rateio)]

                        for _ in range(1):
                            try:
                                with open(conta_contabil_rateio_txt, 'r') as f:
                                    conteudo_conta_contabio_rateio = f.read().strip()
                                    if conteudo_conta_contabio_rateio:
                                        numero_inicial_conta_contabio_rateio = int(conteudo_conta_contabio_rateio)
                                    else:
                                        numero_inicial_conta_contabio_rateio = 19
                            except FileNotFoundError:
                                numero_inicial_conta_contabio_rateio = 19
                            segundo_numero_conta_contabio_rateio = 4
                            resultado_conta_contabio_rateio = numero_inicial_conta_contabio_rateio + segundo_numero_conta_contabio_rateio
                            numero_inicial_conta_contabio_rateio += 4
                            print(resultado_conta_contabio_rateio)
                            with open(conta_contabil_rateio_txt, 'w') as f:
                                f.write(str(numero_inicial_conta_contabio_rateio))

                        conta_contabil_rateio = df.iloc[int(resultado), int(resultado_conta_contabio_rateio)]

                        for _ in range(1):
                            try:
                                with open(centro_custo_rateio_txt, 'r') as f:
                                    conteudo_cc_rateio = f.read().strip()
                                    if conteudo_cc_rateio:
                                        numero_inicial_cc_rateio = int(conteudo_cc_rateio)
                                    else:
                                        numero_inicial_cc_rateio = 20
                            except FileNotFoundError:
                                numero_inicial_cc_rateio = 20
                            segundo_numero_cc_rateio = 4
                            resultado_cc_rateio = numero_inicial_cc_rateio + segundo_numero_cc_rateio
                            numero_inicial_cc_rateio += 4
                            print(resultado_cc_rateio)
                            with open(centro_custo_rateio_txt, 'w') as f:
                                f.write(str(numero_inicial_cc_rateio))

                        centro_custo_rateio = df.iloc[int(resultado), int(resultado_cc_rateio)]

                        for _ in range(1):
                            try:
                                with open(descricao_rateio_txt, 'r') as f:
                                    conteudo_desc_rateio = f.read().strip()
                                    if conteudo_desc_rateio:
                                        numero_inicial_desc_rateio = int(conteudo_desc_rateio)
                                    else:
                                        numero_inicial_desc_rateio = 21
                            except FileNotFoundError:
                                numero_inicial_desc_rateio = 21
                            segundo_numero_desc_rateio = 4
                            resultado_desc_rateio = numero_inicial_desc_rateio + segundo_numero_desc_rateio
                            numero_inicial_desc_rateio += 4
                            print(resultado_desc_rateio)
                            with open(descricao_rateio_txt, 'w') as f:
                                f.write(str(numero_inicial_desc_rateio))

                        descricao_rateio = df.iloc[int(resultado), int(resultado_desc_rateio)]
                    
                        for _ in range(2):
                            pyautogui.press('tab')
                        time.sleep(0.25)
                        pyautogui.press('enter')
                        time.sleep(0.25)
                        pyautogui.press('down')
                        time.sleep(0.25)
                        pyautogui.press('enter')
                        time.sleep(3)
                        for _ in range(25):
                            pyautogui.press('tab')
                        time.sleep(0.25)
                        pyperclip.copy(vlr_rateio)
                        time.sleep(0.25)
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(0.25)
                        pyautogui.press('tab')
                        time.sleep(0.25)
                        pyperclip.copy(conta_contabil_rateio)
                        time.sleep(0.25)
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(0.25)
                        pyautogui.press('tab')
                        time.sleep(0.25)
                        if centro_custo_rateio == 0:
                            pyautogui.write("")
                        else:
                            pyperclip.copy(centro_custo_rateio)
                            time.sleep(0.25)
                            pyautogui.hotkey('ctrl', 'v')
                            time.sleep(0.25)
                        time.sleep(0.25)
                        pyautogui.press('tab')
                        pyperclip.copy(descricao_rateio)
                        time.sleep(0.25)
                        pyautogui.hotkey('ctrl', 'v')
                        time.sleep(0.25)                           

                        criar_cont_xpath = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div[2]/div/div/div/div[2]/div[1]/div/div/div/div/div[1]/div[1]/div/div/div[3]/div[1]/div/table/tbody/tr[17]/td/fieldset/div/span[1]/table/tbody/tr[2]/td[2]"
                        try:
                            exe_criar_cont = wait.until(EC.presence_of_element_located((By.XPATH, criar_cont_xpath)))
                            exe_criar_cont.click()
                            print("Elemento encontrado e clicado com sucesso.")
                        except Exception as e:
                            print(f"Erro ao encontrar o elemento: {e}")

                        time.sleep(5)
                        pyautogui.press('enter')
                        time.sleep(0.25)

                        dynamic_links_xpath = "/html/body/div/div/div/div/div[1]/div[1]/div[2]/div/div/div[1]/div[3]/ul/li[2]/div/div/div/span"
                        try:
                            exe_dynamic_links = wait.until(EC.presence_of_element_located((By.XPATH, dynamic_links_xpath)))
                            exe_dynamic_links.click()
                            print("Elemento encontrado e clicado com sucesso.")
                        except Exception as e:
                            print(f"Erro ao encontrar o elemento: {e}")

                        time.sleep(5)
                finally:
                    with open(valor_rateio_txt, 'w') as f:
                        f.write('')

                    with open(conta_contabil_rateio_txt, 'w') as f:
                        f.write('')

                    with open(centro_custo_rateio_txt, 'w') as f:
                        f.write('')

                    with open(descricao_rateio_txt, 'w') as f:
                        f.write('')
                                        
                fields_cont_xpath = "/html/body/div[1]/div/div/div/div[1]/div[1]/div[2]/div/div/div[1]/div[3]/ul/li[1]/div/div/div/span"
                try:
                    exe_fields_cont = wait.until(EC.presence_of_element_located((By.XPATH, fields_cont_xpath)))
                    exe_fields_cont.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")

                cadiado_xpath = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div[1]/div[1]/div[2]/table/tbody/tr[1]/td/div/div/div/div/fieldset/div/table/tbody/tr/td[1]/table/tbody/tr/td[3]/span/div/img"
                try:
                    exe_cadiado = wait.until(EC.presence_of_element_located((By.XPATH, cadiado_xpath)))
                    exe_cadiado.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")

                time.sleep(2)
                try:
                    exe_fields_cont = wait.until(EC.presence_of_element_located((By.XPATH, fields_cont_xpath)))
                    exe_fields_cont.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")
                time.sleep(0.25)
                for _ in range(14):
                    pyautogui.press('tab')
                time.sleep(0.25)
                for _ in range(2):
                    pyautogui.press('up')
                time.sleep(0.25)
                pyautogui.press('tab')
                time.sleep(5)
                                        
                salvar_cont_xpath = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div[1]/div[1]/div[2]/table/tbody/tr[1]/td/div/div/div/div/fieldset/div/table/tbody/tr/td[1]/table/tbody/tr/td[15]/span/div/img"
                try:
                    exe_salvar_cont = wait.until(EC.presence_of_element_located((By.XPATH, salvar_cont_xpath)))
                    exe_salvar_cont.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")

                time.sleep(1)
                pyautogui.press('enter')
                time.sleep(2)
                pyautogui.press('F11')
                time.sleep(0.25)
                pyautogui.hotkey('alt', 'F4')
                time.sleep(1)

                driver.switch_to.window(abas[0])

                clear_xpath = "/html/body/div[1]/div/div/div/div[2]/div/div/div/div/div/div[1]/div[1]/div/div/div[3]/div[1]/div/fieldset/div/span[1]/table/tbody/tr[2]/td[2]"
                try:
                    exe_clear = wait.until(EC.presence_of_element_located((By.XPATH, clear_xpath)))
                    exe_clear.click()
                    print("Elemento encontrado e clicado com sucesso.")
                except Exception as e:
                    print(f"Erro ao encontrar o elemento: {e}")

                time.sleep(2)


    contabilizar_nf_thuban_rpa()

    messagebox.showinfo("Status", "Notas Contabilizadas!")

def interface():
    root= tk.Tk()
    root.title('Thuban RPA')
    root.geometry("400x250")

    def webdriver():
        rodar_webdriver()
    
    def criar():
        criar_nf_thuban()

    def contabilizar():
        contabilizar_nf_thuban()
    
    frame = tk.Frame(root)
    frame.pack(pady=20)
    label = tk.Label(frame, text="Selecione Uma Opção")
    label.pack()

    rodar_webdriver_button = tk.Button(frame, text ="1. Rodar Edge", command=webdriver)
    rodar_webdriver_button.pack(pady=10)

    criar_nf_button = tk.Button(frame, text="2. Criar Notas", command=criar)
    criar_nf_button.pack(pady=10)

    contabilizar_nf_button = tk.Button(frame, text="3. Contabilizar Notas", command=contabilizar)
    contabilizar_nf_button.pack(pady=10)

    sair_button = tk.Button(frame, text="0. Sair do Sistema", command=root.quit)
    sair_button.pack(pady=10)
    root.mainloop()

if __name__ == "__main__":
    interface()



