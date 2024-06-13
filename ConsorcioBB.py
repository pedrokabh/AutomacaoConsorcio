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
                        taxas = f"{txAdm:.2f}% + {fReserv:.2f}%"
                        calcTxAdmMensal = f"{(txAdm / prazo):.3f}% A.M\n{(txAdm / prazo * 12):.3f}% A.A"
                        calcFRMensal = f"{(fReserv / prazo):.3f}% A.M\n{(fReserv / prazo * 12):.3f}% A.A"
                        calcTotalMensal = f"{((txAdm + fReserv) / prazo):.3f}% A.M\n{((txAdm + fReserv) / prazo * 12):.3f}% A.A"

                        cartas_Grupo = {
                            "Modalidade": self.traduzir_sigla(sigla),
                            "Grupo": int(codigo_grupo[2:]),  # Tirar 2 primeiros caracteres
                            "Prazo": prazo,
                            "Vagas": vagas,
                            "Taxas": taxas,
                            "Calculo TxAdm": calcTxAdmMensal,
                            "Calculo FR": calcFRMensal,
                            "Calculo Total": calcTotalMensal,
                            "CartasCredito": valorBem
                        }
                        grupos_ativos.append(cartas_Grupo)

                
                if just_one_table == True:
                    cod = 0
            
            self.browser.find_element(By.XPATH,'//*[@id="ctl00_lnkHome"]').click()
            df = pd.DataFrame(grupos_ativos)
            df = df.loc[df.groupby(['Grupo', 'Prazo'])['CartasCredito'].idxmin()]
            return df      
        except Exception as err:
            print(f"[ConsorcioBB] Falha no metodo generate_dataFrame_gruposAtivos. \n {err}")
            sys.exit(1)

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
    
    # --- FUNÇÕES PARA GERAR DADOS ASSEMBLEIAS --- #
    def generate_dataFrame_dadosAssembleia(self, lista_grupos):
        try:
            time.sleep(3)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_rptFormularios_ctl27_lnkFormulario"]'))).click()
            dados_grupos = []

            for grupo in lista_grupos:

                script = f"document.evaluate('//*[@id=\"ctl00_Conteudo_edtCD_Grupo\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{str(grupo).zfill(6)}';"
                self.browser.execute_script(script)
                script = "document.getElementById('ctl00_Conteudo_lnkPesquisar').click();"
                self.browser.execute_script(script)

                script = "return document.getElementById('ctl00_Conteudo_lblVolumeTotal').textContent;"
                liquidez = self.browser.execute_script(script)

                # Definindo dados.
                dados_assembleias_grupo = {
                    "Grupo": int(grupo),
                    "Liquidez": liquidez,
                    "N° Assembleia M5": None, "Contemplados M5": None,  "Menor Lance M5": None,
                    "N° Assembleia M4": None, "Contemplados M4": None,  "Menor Lance M4": None,
                    "N° Assembleia M3": None, "Contemplados M3": None,  "Menor Lance M3": None,
                }

                tabela = self.browser.find_element(By.ID, 'ctl00_Conteudo_grdAssembleias')
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')[1:]

                for linha in linhas:
                    celulas = linha.find_elements(By.TAG_NAME, 'td')

                    assembleia = celulas[1].text
                    data = celulas[2].text
                    qtde_contemp = celulas[3].text
                    menor_lance = celulas[5].text

                    if data == '27/05/2024':
                        dados_assembleias_grupo["N° Assembleia M5"] = assembleia
                        dados_assembleias_grupo["Contemplados M5"] = qtde_contemp
                        dados_assembleias_grupo["Menor Lance M5"] = "{:.2f}%".format(float(menor_lance.replace(",", ".")))
                    elif data == '25/04/2024':
                        dados_assembleias_grupo["N° Assembleia M4"] = assembleia
                        dados_assembleias_grupo["Contemplados M4"] = qtde_contemp
                        dados_assembleias_grupo["Menor Lance M4"] = "{:.2f}%".format(float(menor_lance.replace(",", ".")))
                    elif data == '25/03/2024':
                        dados_assembleias_grupo["N° Assembleia M3"] = assembleia
                        dados_assembleias_grupo["Contemplados M3"] = qtde_contemp
                        dados_assembleias_grupo["Menor Lance M3"] = "{:.2f}%".format(float(menor_lance.replace(",", ".")))
                
                dados_grupos.append(dados_assembleias_grupo)

            self.browser.find_element(By.XPATH,'//*[@id="ctl00_lnkHome"]').click()
            df = pd.DataFrame(dados_grupos)
            return df
        except Exception as err:
            print(f"\n[ConsorcioBB] Falha no Metodo generate_dataFrame_dadosAssembleia.\n{err}")
            sys.exit(1)

    # --- FUNÇÕES PARA CALCULAR MEDIA DE CONTEMPLAÇÃO DOS GRUPOS --- #
    def extrair_MediaContemplacao(self, df_dadosAssembleia, lista_grupos):
        try:
            dados_grupos = [] # Adiciona todos os dados dos grupos.
            for grupo in lista_grupos:
                dados_assembleias_grupo = {
                    "Grupo": int(grupo),
                    "Media Lance M5": None,
                    "Media Lance M4": None,
                    "Media Lance M3": None
                }
                countLinhas = 0 # Conta quantas linhas tem na tabela de assembleias.
                execution = 0 # Contabiliza qual linha processar.
                
                try:
                    # Navega até assembleia.
                    self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_rptFormularios_ctl27_lnkFormulario"]'))).click()
                    # Insere o grupo.
                    self.browser.execute_script(f"document.evaluate('//*[@id=\"ctl00_Conteudo_edtCD_Grupo\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{str(grupo).zfill(6)}';")
                    # Pesquisa dados do grupo.
                    self.browser.execute_script("document.getElementById('ctl00_Conteudo_lnkPesquisar').click();")
                    
                    # Acha a tabela.
                    tabela = self.browser.find_element(By.ID, 'ctl00_Conteudo_grdAssembleias')
                    
                    # Contabiliza linhas pulando cabeçalho.
                    linhas = tabela.find_elements(By.TAG_NAME, 'tr')[1:]
                    countLinhas = len(linhas) - 1
                    execution = 0

                    while execution <= countLinhas:
                        # Volta no menu
                        self.wait.until(EC.presence_of_element_located((By.ID, 'ctl00_lnkHome'))).click()
                        # Entra na página da assembleia novamente
                        self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_rptFormularios_ctl27_lnkFormulario"]'))).click()
                        # Insere o grupo.
                        self.browser.execute_script(f"document.evaluate('//*[@id=\"ctl00_Conteudo_edtCD_Grupo\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{str(grupo).zfill(6)}';")
                        # Pesquisa dados do grupo.
                        self.browser.execute_script("document.getElementById('ctl00_Conteudo_lnkPesquisar').click();")
                        # Acha a tabela.
                        tabela = self.browser.find_element(By.ID, 'ctl00_Conteudo_grdAssembleias')
                        # Pega as linhas da tabela.
                        linhas = tabela.find_elements(By.TAG_NAME, 'tr')[1:]
                        # Pega a linha desejada.
                        linha = linhas[execution]

                        # Pegando valores das colunas da linha e criando variavel de media.
                        celulas = linha.find_elements(By.TAG_NAME, 'td')
                        # NumeroAssembleia = celulas[1].text
                        data = celulas[2].text
                        # contemplados = celulas[3].text
                        # menor_lance = celulas[5].text
                        buton_CotasContempladas = celulas[6].find_element(By.TAG_NAME, 'input')
                        media_lance = None
                        
                        # ABRINDO ABA DA ASSEMBLEIA PARA CALCULAR MEDIA LANCE
                        buton_CotasContempladas.click()
                        media_lance = self.dados_cotas_contemplados(grupo, df_dadosAssembleia)

                        if data == '27/05/2024':
                            dados_assembleias_grupo["Media Lance M5"] = f"{float(media_lance):.2f}".replace('.', ',') + '%'
                        elif data == '25/04/2024':
                            dados_assembleias_grupo["Media Lance M4"] = f"{float(media_lance):.2f}".replace('.', ',') + '%'
                        elif data == '25/03/2024':
                            dados_assembleias_grupo["Media Lance M3"] = f"{float(media_lance):.2f}".replace('.', ',') + '%'
                        
                        # Voltando para menu principal.
                        self.browser.find_element(By.XPATH,'//*[@id="ctl00_lnkHome"]').click()
                        execution += 1

                    dados_grupos.append(dados_assembleias_grupo)
                except Exception as err:
                    print(f"[ConsorcioBB] error\n{err}")
            
            df = pd.DataFrame(dados_grupos)
            self.browser.find_element(By.XPATH,'//*[@id="ctl00_lnkHome"]').click()
            return df
        except Exception as err:
                    print(f"[ConsorcioBB] error\n{err}")

    def dados_cotas_contemplados(self, grupo, df_dadosAssembleia):
        try:
            # MÉTODO PARA CALCULAR MÉDIA DE LANCE DA ASSEMBLEIA.
            # CALCULAR PARA CADA ASSEMBLEIA OU SEJA PARA CADA LINHA DA TABELA.
            end_table = False
            list_cotasContempladas = []

            # Pega informações dos contemplados da assembleia.
            while end_table == False:
                tabelaContemplados = WebDriverWait(self.browser, timeout=10).until(EC.presence_of_element_located((By.ID, 'ctl00_Conteudo_grdDetalhesAssembleia')))
                linhasContemplados = tabelaContemplados.find_elements(By.TAG_NAME, 'tr')[1:]  # Ignorar a primeira linha de cabeçalho

                # Leitura dos dados de contemplação.
                for index, linhaAtual in enumerate(linhasContemplados):
                    if linhaAtual.find_elements(By.XPATH,'.//input[contains(@src, "next.png")]'):  
                            linhaAtual.find_element(By.XPATH,'.//input[contains(@src, "next.png")]').click()
                            break
                    elif linhaAtual.find_elements(By.XPATH,'.//input[contains(@src, "back.png")]'):
                            end_table = True
                            break
                    else:
                        celulas = linhaAtual.find_elements(By.TAG_NAME, 'td')
                        cota = celulas[0].text
                        percentual_lance = celulas[2].text
                        dados_contemplacao = {
                            "Cota": cota,
                            "Percentual": percentual_lance,
                        }
                        list_cotasContempladas.append(dados_contemplacao)
                        
                        # Verifica se é a última linha da tabela
                        if index == len(linhasContemplados) - 1:
                            end_table = True
                            break

            # Inicialize as variáveis para calcular a média e contar o número de valores válidos
            soma_percentuais = 0
            contagem = 0
            media_percentuais = None
            pecentuais = []

            # Obtenha os valores das colunas "MenorLance M5", "MenorLance M4" e "MenorLance M3" para o grupo correspondente
            dados_grupo = df_dadosAssembleia[df_dadosAssembleia['Grupo'] == grupo]
            if not dados_grupo.empty:
                mel_05_24 = dados_grupo['Menor Lance M5'].values[0]
                mel_04_24 = dados_grupo['Menor Lance M4'].values[0]
                mel_03_24 = dados_grupo['Menor Lance M3'].values[0]

                for CotaContemplada in list_cotasContempladas:

                    percentual = CotaContemplada['Percentual']
                    if percentual == '0,0000': # Desconsidera sorteio.
                        continue
                    elif percentual == "20,0000" and ( mel_05_24 == '20.00%' or None) and (mel_04_24 == '20.00%' or None) and (mel_03_24 == '20.00%' or None):
                        # lanceFixo20 = True
                        continue
                    elif percentual == "30,0000" and (mel_05_24 == '30.00%' or None) and (mel_04_24 == '30.00%' or None) and (mel_03_24 == '30.00%' or None):
                        # lanceFixo30 = True
                        continue
                    else:
                        pecentuais.append(percentual)
                        soma_percentuais += float(percentual.replace(',', '.'))  # Converta para float, substituindo ',' por '.'
                        contagem += 1

            # Calcule a média
            if contagem > 0:
                media_percentuais = soma_percentuais / contagem
            else:
                print(f"[WARINING] Não há valores válidos para calcular a média. Grupo {grupo}°")
                return media_percentuais == 0
            
            # Debug.
            # if lanceFixo20:
            #     print(f"[ConsorcioBB] Grupo {grupo}° Lance Fixo 20%.")
            #     print("Percentuais ->",pecentuais)
            #     print("Contagem: ",contagem)
            #     print("Media Lance: ", media_percentuais)
            #     print()
            # elif lanceFixo30:
            #     print(f"[ConsorcioBB] Grupo {grupo}° Lance Fixo 30%.")
            #     print("Percentuais ->",pecentuais)
            #     print("Contagem: ",contagem)
            #     print("Media Lance: ", media_percentuais)
            #     print()

            self.browser.find_element(By.XPATH,"//*[@id='ctl00_Conteudo_lnkFechaDiv']").click()
            return media_percentuais

        except Exception as err:
            print(f"error\n{err}")

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


