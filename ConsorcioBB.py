# -- ANOTAÇÕES -- #
# 1.0 Testes Realizados com Chrome Driver v 125.0.6422.141 - https://googlechromelabs.github.io/chrome-for-testing/#stable
# 1.1 Testes Realizados com Selenium v 4.21.0

try:
    # Importações Selenium
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    # Importação para criar serviço com driver chrome
    from selenium.webdriver.chrome.service import Service
    # Importações analise de dados
    import pandas as pd
    # Importações padrões
    import time
    import os
except Exception as err:
    print(f"[ERROR] \n{err}")

class ConsorcioBB:
    def __init__(self, login, senha):

        # -- fluxo de execução --#
        os.system('cls')
        print("Sistema Rodando Selenium Versão (" + webdriver.__version__ + ").")
        
        self.login = login
        self.senha = senha
        self.chrome_driver_path = r"C:\Users\pedro\OneDrive\Área de Trabalho\AUTOMAÇÕES ISF\chromedriver-win64\chromedriver.exe"

        self.iniciar_navegador()
        self.fazer_login()
        self.fechar_popup()

    # --- FUNÇÕES PARA GERAR GRUPOS ATIVOS --- #
  
    def generate_dataFrame_gruposAtivos(self, sigla):
        try:
            grupos_ativos = []
            self.browser.execute_script("document.getElementById('ctl00_Conteudo_rptFormularios_ctl08_lnkFormulario').click();") # Clicando em simulador de contratação
            time.sleep(3)

            if sigla == "IM240": # Grupos IMOVEIS 240
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoGrupo"]/option[3]'))).click()
                time.sleep(2)
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]'))).click()
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]/option[2]'))).click()
            elif sigla == "IMP": # Grupos IMOVEIS PADROES
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoGrupo"]/option[3]'))).click()
                time.sleep(2)
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]'))).click()
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]/option[3]'))).click()
            else:
                self.browser.execute_script(f"document.querySelector('#ctl00_Conteudo_cbxTipoGrupo option[value=\"{sigla}\"]').selected = true")
            
            self.browser.execute_script("document.getElementById('ctl00_Conteudo_lnkProximo').click();") # Pesquisar categoria
            cod = 1

            while cod != 0:
                time.sleep(2)
                tabela = self.browser.find_element(By.XPATH, '//*[@id="ctl00_Conteudo_grdGruposDisponiveis"]')  # Localizar a tabela pelo ID
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')  # Encontrar todas as linhas da tabela, ignorando o cabeçalho
                just_one_table = True

                for linha in linhas:
                    if linha.find_elements(By.XPATH, './/input[contains(@src, "next.png")]'):
                        just_one_table = False

                for linha in linhas:
                    if linha.find_elements(By.XPATH, './/input[contains(@src, "next.png")]'):
                        self.browser.execute_script("__doPostBack('ctl00$Conteudo$grdGruposDisponiveis','Page$Next')")
                        break 
                    elif linha.find_elements(By.XPATH,'.//input[contains(@src, "back.png")]'):
                        cod = 0
                        break
                    elif len(linha.find_elements(By.TAG_NAME, 'td'))>3:
                        celulas = linha.find_elements(By.TAG_NAME, 'td')  # Encontrar todas as células da linha

                        codigo_grupo = celulas[0].text
                        prazo = celulas[1].text
                        vagas = celulas[2].text
                        txAdm = celulas[4].text
                        fReserv = celulas[5].text
                        valorBem = celulas[7].text

                        cartas_Grupo = {
                            "TipoGrupo": self.traduzir_sigla(sigla),
                            "Codigo": int(codigo_grupo[2:]), # Tirar 2 primeiros caracter
                            "Prazo": int(prazo),
                            "Vagas": int(vagas),
                            "TaxaAdm": f"{float(txAdm.replace(',', '.')):.2f}%",
                            "FReserva": f"{float(fReserv.replace(',', '.')):.2f}%",
                            "CartasCredito": float(valorBem.replace('.', '').replace(',', '.'))
                        }
                        grupos_ativos.append(cartas_Grupo)
                
                # Verifica se tabela tem apenas uma página para encerrar metodo.
                if just_one_table == True:
                    self.browser.find_element(By.XPATH,'//*[@id="ctl00_lnkHome"]').click() # Volta para menu principal.
                    df = pd.DataFrame(grupos_ativos)
                    df = df.loc[df.groupby(['Codigo', 'Prazo'])['CartasCredito'].idxmin()] # Filtrando apenas menor carta
                    return df # Retorna df com grupos ativos da sigla.
                
            self.browser.find_element(By.XPATH,'//*[@id="ctl00_lnkHome"]').click() # Volta para menu principal.
            df = pd.DataFrame(grupos_ativos)
            df = df.loc[df.groupby(['Codigo', 'Prazo'])['CartasCredito'].idxmin()] # Filtrando apenas menor carta
            return df # Retorna df com grupos ativos da sigla.
                
        except Exception as err:
            print(f"[ERROR] Falha ao levantar grupos ativos. \n {err}")

    # --- FUNÇÕES PARA GERAR DADOS ASSEMBLEIAS --- #
    def generate_dataFrame_dadosAssembleia(self, lista_grupos):
        try:
            print("\nEXTRAINDO DADOS DA ASSEMBLEIA")
            time.sleep(3)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_rptFormularios_ctl27_lnkFormulario"]'))).click() # Clicando na página da assembleia
            
            dados_grupos = []

            for grupo in lista_grupos:
                dados_assembleias_grupo = {
                    "Codigo": int(grupo),
                    "VolumeTotal": None,
                    "A 05/24": None, "QtdCont 05/24": None,  "MEL 05/24": None,
                    "A 04/24": None, "QtdCont 04/24": None,  "MEL 04/24": None,
                    "A 03/24": None, "QtdCont 03/24": None,  "MEL 03/24": None,
                }

                script = f"document.evaluate('//*[@id=\"ctl00_Conteudo_edtCD_Grupo\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{str(grupo).zfill(6)}';"
                self.browser.execute_script(script) # Alterando valor do grupo
                script = "document.getElementById('ctl00_Conteudo_lnkPesquisar').click();"
                self.browser.execute_script(script) # Clicando em pesquisar

                script = "return document.getElementById('ctl00_Conteudo_lblVolumeTotal').textContent;"# Script para pegar Volume total
                volume_total_label = self.browser.execute_script(script) # Executar o script usando execute_script
                dados_assembleias_grupo["VolumeTotal"] = volume_total_label

                tabela = self.browser.find_element(By.ID, 'ctl00_Conteudo_grdAssembleias')# Localizar a tabela pelo ID
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')[1:] # Encontrar todas as linhas da tabela, ignorando o cabeçalho

                # Iterar sobre as linhas e extrair os valores das primeiras tds de cada linha
                for linha in linhas:
                    celulas = linha.find_elements(By.TAG_NAME, 'td') # Encontrar todas as células da linha

                    assembleia = celulas[1].text
                    data = celulas[2].text
                    qtde_contemp = celulas[3].text
                    # maior_lance = celulas[4].text
                    menor_lance = celulas[5].text

                    # Adiciona os dados à estrutura do grupo
                    if data == '27/05/2024':
                        dados_assembleias_grupo["A 05/24"] = assembleia
                        dados_assembleias_grupo["QtdCont 05/24"] = qtde_contemp
                        # dados_assembleias_grupo["MAL 05/24"] = "{:.2f}%".format(float(maior_lance.replace(",", ".")))
                        dados_assembleias_grupo["MEL 05/24"] = "{:.2f}%".format(float(menor_lance.replace(",", ".")))
                    elif data == '25/04/2024':
                        dados_assembleias_grupo["A 04/24"] = assembleia
                        dados_assembleias_grupo["QtdCont 04/24"] = qtde_contemp
                        # dados_assembleias_grupo["MAL 04/24"] = "{:.2f}%".format(float(maior_lance.replace(",", ".")))
                        dados_assembleias_grupo["MEL 04/24"] = "{:.2f}%".format(float(menor_lance.replace(",", ".")))
                    elif data == '25/03/2024':
                        dados_assembleias_grupo["A 03/24"] = assembleia
                        dados_assembleias_grupo["QtdCont 03/24"] = qtde_contemp
                        # dados_assembleias_grupo["MAL 03/24"] = "{:.2f}%".format(float(maior_lance.replace(",", ".")))
                        dados_assembleias_grupo["MEL 03/24"] = "{:.2f}%".format(float(menor_lance.replace(",", ".")))
                
                dados_grupos.append(dados_assembleias_grupo)

            df = pd.DataFrame(dados_grupos)
            return df
        except Exception as err:
            print(f"\n[ERROR] Falha ao extrair dados da assembleia.\n{err}")

    # --- FUNÇÕES COMPLEMENTARES ---#
    def traduzir_sigla(self, sigla):
        dicionario_siglas = {
            "MO": "MOTO DEMAIS",
            "EE": "OUTROS BENS MOVEIS",
            "TC": "TRATOR E CAMINHAO GERAL",
            "IM240": "IMOVEIS 240",
            "IMP": "IMOVEIS PADRAO",
            "AI": "AUTO IPCA",
            "AU": "AUTO DEMAIS"
        }
        return dicionario_siglas.get(sigla, "Sigla desconhecida")
    
    # --- FUNÇÕES PADRÕES -- #
    def iniciar_navegador(self):
        try:
            print("INICIANDO NAVEGADOR CHROME")
            chrome_service = Service(self.chrome_driver_path)
            self.browser = webdriver.Chrome(service=chrome_service)
            self.wait = WebDriverWait(self.browser, 10) 
            print("NAVEGADOR INICIADO COM SUCESSO!\nACESSANDO PORTAL PARCEIROS BB CONSÓRCIO (www.parceirosbbconsorcios.com.br)")
            self.browser.get('https://www.parceirosbbconsorcios.com.br/')
        except Exception as err:
            print(f"\n[ERROR] Falha ao inicializar navegador.\n{err}")
    
    def fazer_login(self):
        try:
            print("REALIZANDO LOGIN NO PORTAL PARCEIROS BB")
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_edtUsuario"]'))).send_keys(self.login)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_edtSenha"]'))).send_keys(self.senha)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_lnkEntrar"]'))).click()
            print("LOGON EFETUADO COM SUCESSO.")
        except Exception as err:
            print(f"\n[ERROR] Falha ao realizar login.\n{err}")
    
    def fechar_popup(self):
        try:
            print('FECHANDO POUP-UP')
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_LinkButton1"]'))).click()
            print("POUP-UP FECHADO")
        except Exception as err:
            print(f"\n[ERROR] Falha ao fechar poup-Up pós login.\n{err}")
            return
    
    def encerrar_navegador(self):
        try:
            # Encerra navegador.
            self.browser.quit()
        except Exception as err:
            print(f"\n[ERROR] Falha ao encerrar browser.\n{err}")


