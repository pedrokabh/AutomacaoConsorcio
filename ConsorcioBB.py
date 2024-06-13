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
    import os, sys
except Exception as err:
    print(f"[ConsorcioBB] FALHA AO CARREGAR IMPORTAÇÕES \n{err}")
    sys.exit(1)

class ConsorcioBB:
    def __init__(self, login, senha, endereçoChromeDriver):

        # 1.0 # VERIFICANDO VERSÃO SELENIUM
        os.system('cls')
        print("[ConsorcioBB] Versão Selenium Em Execução (V" + webdriver.__version__ + ").")
        # 1.0 #
        
        # 1.1 # INICIANDO NAVEGADOR, FAZENDO LOGON E FECHANDO POUP-UP.
        self.login = login
        self.senha = senha
        self.chrome_driver_path = endereçoChromeDriver
        self.iniciar_navegador()
        self.fazer_login()
        self.fechar_popup()
        # 1.1 #

    # --- FUNÇÕES PARA GERAR GRUPOS ATIVOS --- #
    def generate_dataFrame_gruposAtivos(self, sigla):
        try:
            grupos_ativos = []
            self.browser.execute_script("document.getElementById('ctl00_Conteudo_rptFormularios_ctl08_lnkFormulario').click();")
            time.sleep(3)

            if sigla == "IM240":
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoGrupo"]/option[3]'))).click()
                time.sleep(2)
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]'))).click()
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]/option[2]'))).click()
            elif sigla == "IMP":
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoGrupo"]/option[3]'))).click()
                time.sleep(2)
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]'))).click()
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]/option[3]'))).click()
            else:
                self.browser.execute_script(f"document.querySelector('#ctl00_Conteudo_cbxTipoGrupo option[value=\"{sigla}\"]').selected = true")
            
            self.browser.execute_script("document.getElementById('ctl00_Conteudo_lnkProximo').click();")
            cod = 1

            while cod != 0:
                time.sleep(2)
                tabela = self.browser.find_element(By.XPATH, '//*[@id="ctl00_Conteudo_grdGruposDisponiveis"]')
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')
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
                    elif len(linha.find_elements(By.TAG_NAME, 'td')) > 3:
                        celulas = linha.find_elements(By.TAG_NAME, 'td')  # Encontrar todas as células da linha

                        codigo_grupo = celulas[0].text
                        prazo = int(celulas[1].text)
                        vagas = int(celulas[2].text)
                        txAdm = float(celulas[4].text.replace(',', '.'))  # Converte string para float
                        fReserv = float(celulas[5].text.replace(',', '.'))  # Converte string para float
                        valorBem = float(celulas[7].text.replace('.', '').replace(',', '.'))  # Converte string para float

                        # Calculando novas colunas
                        txAdmFr = f"{txAdm:.2f}% + {fReserv:.2f}%"
                        calcTxAdmMensal = f"{(txAdm / prazo):.3f}% a.m\n{(txAdm / prazo * 12):.3f}% a.a"
                        calcFRMensal = f"{(fReserv / prazo):.3f}% a.m\n{(fReserv / prazo * 12):.3f}% a.a"
                        calcTotalMensal = f"{((txAdm + fReserv) / prazo):.3f}% a.m\n{((txAdm + fReserv) / prazo * 12):.3f}% a.a"

                        cartas_Grupo = {
                            "TipoGrupo": self.traduzir_sigla(sigla),
                            "Codigo": int(codigo_grupo[2:]),  # Tirar 2 primeiros caracteres
                            "Prazo": prazo,
                            "Vagas": vagas,
                            "TxAdm+FR": txAdmFr,
                            "CalcTxAdm": calcTxAdmMensal,
                            "CalcFR": calcFRMensal,
                            "CalcTotal": calcTotalMensal,
                            "CartasCredito": valorBem
                        }
                        grupos_ativos.append(cartas_Grupo)

                
                if just_one_table == True:
                    cod = 0
            
            self.browser.find_element(By.XPATH,'//*[@id="ctl00_lnkHome"]').click()
            df = pd.DataFrame(grupos_ativos)
            df = df.loc[df.groupby(['Codigo', 'Prazo'])['CartasCredito'].idxmin()]
            return df      
        except Exception as err:
            print(f"[ConsorcioBB] Falha no metodo generate_dataFrame_gruposAtivos. \n {err}")
            sys.exit(1)

    # --- FUNÇÕES PARA GERAR DADOS ASSEMBLEIAS --- #
    def generate_dataFrame_dadosAssembleia(self, lista_grupos):
        try:
            time.sleep(3)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_rptFormularios_ctl27_lnkFormulario"]'))).click()
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
                self.browser.execute_script(script)
                script = "document.getElementById('ctl00_Conteudo_lnkPesquisar').click();"
                self.browser.execute_script(script)

                script = "return document.getElementById('ctl00_Conteudo_lblVolumeTotal').textContent;"
                volume_total_label = self.browser.execute_script(script)
                dados_assembleias_grupo["VolumeTotal"] = volume_total_label

                tabela = self.browser.find_element(By.ID, 'ctl00_Conteudo_grdAssembleias')
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')[1:]

                for linha in linhas:
                    celulas = linha.find_elements(By.TAG_NAME, 'td')

                    assembleia = celulas[1].text
                    data = celulas[2].text
                    qtde_contemp = celulas[3].text
                    menor_lance = celulas[5].text

                    if data == '27/05/2024':
                        dados_assembleias_grupo["A 05/24"] = assembleia
                        dados_assembleias_grupo["QtdCont 05/24"] = qtde_contemp
                        dados_assembleias_grupo["MEL 05/24"] = "{:.2f}%".format(float(menor_lance.replace(",", ".")))
                    elif data == '25/04/2024':
                        dados_assembleias_grupo["A 04/24"] = assembleia
                        dados_assembleias_grupo["QtdCont 04/24"] = qtde_contemp
                        dados_assembleias_grupo["MEL 04/24"] = "{:.2f}%".format(float(menor_lance.replace(",", ".")))
                    elif data == '25/03/2024':
                        dados_assembleias_grupo["A 03/24"] = assembleia
                        dados_assembleias_grupo["QtdCont 03/24"] = qtde_contemp
                        dados_assembleias_grupo["MEL 03/24"] = "{:.2f}%".format(float(menor_lance.replace(",", ".")))
                
                dados_grupos.append(dados_assembleias_grupo)

            df = pd.DataFrame(dados_grupos)
            return df
        except Exception as err:
            print(f"\n[ConsorcioBB] Falha no Metodo generate_dataFrame_dadosAssembleia.\n{err}")
            sys.exit(1)

    # --- FUNÇÕES COMPLEMENTARES ---#
    def traduzir_sigla(self, sigla):
        dicionario_siglas = {
            "MO": "MOTO DEMAIS",
            "EE": "BENS MOVEIS",
            "TC": "TRATOR E CAMINHAO",
            "IM240": "IMOVEIS 240",
            "IMP": "IMOVEIS PADRAO",
            "AI": "AUTO IPCA",
            "AU": "AUTO DEMAIS"
        }
        return dicionario_siglas.get(sigla, "Sigla desconhecida")
    
    # --- FUNÇÕES PADRÕES -- #
    def iniciar_navegador(self):
        try:
            print("\n[ConsorcioBB] Iniciando Chrome...")
            chrome_service = Service(self.chrome_driver_path)
            self.browser = webdriver.Chrome(service=chrome_service)
            self.wait = WebDriverWait(self.browser, 10) 
        except Exception as err:
            print(f"\n[ConsorcioBB] Falha no método. iniciar_navegador.\n{err}")
            sys.exit(1)
    
    def fazer_login(self):
        try:
            print("\n[ConsorcioBB] Acessando URL (https://www.parceirosbbconsorcios.com.br).")
            self.browser.get('https://www.parceirosbbconsorcios.com.br/')
            print("[ConsorcioBB] Efetuando login PP BB Consórcio.")
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_edtUsuario"]'))).send_keys(self.login)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_edtSenha"]'))).send_keys(self.senha)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_lnkEntrar"]'))).click()
            print("[ConsorcioBB] Logado com sucesso no PP BB Consórcio.")
        except Exception as err:
            print(f"\n[ConsorcioBB] Falha no metodo fazer_login.\n{err}")
            sys.exit(1)
    
    def fechar_popup(self):
        try:
            print('[ConsorcioBB] Fechando Poup-Up Inicial.')
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_LinkButton1"]'))).click()
            print("[ConsorcioBB] Poup-Up fechado.")
        except Exception as err:
            print(f"\n[ConsorcioBB] Falha no metodo fechar_popUp.\n{err}")
            sys.exit(1)
    
    def encerrar_navegador(self):
        try:
            # Encerra navegador.
            print("[ConsorcioBB] Encerrando Chrome.")
            self.browser.quit()
            sys.exit(0)
        except Exception as err:
            print(f"\n[ConsorcioBB] Falha no método encerrar_navegador.\n{err}")
            sys.exit(1)


