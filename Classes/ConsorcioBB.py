try:
    # Importações Selenium
    from selenium import webdriver
    from selenium.webdriver.support.ui import Select
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
    import datetime
except Exception as err:
    print(f"[ConsorcioBB] FALHA AO CARREGAR IMPORTAÇÕES \n{err}")
    raise

class ConsorcioBB:
    def __init__(self, login, senha, endereco_chrome_driver, logger, cwd, data_assembleia_mais_recente=None, data_assembleia_passada=None, data_assembleia_retrasada=None):
        if data_assembleia_mais_recente != None and data_assembleia_passada != None and data_assembleia_retrasada != None:
            # Variaveis Assembleia Mais Recente
            self.data_assembleia_mais_recente = data_assembleia_mais_recente
            self.sigla_assembleia_mais_recente = self.ReturnsMonthCode(data_assembleia_mais_recente)
            # Variaveis Assembleia Passada
            self.data_assembleia_passada = data_assembleia_passada
            self.sigla_assembleia_passada = self.ReturnsMonthCode(data_assembleia_passada)
            # Variaveis Assembleia Retrasada
            self.data_assembleia_retrasada = data_assembleia_retrasada
            self.sigla_assembleia_retrasada = self.ReturnsMonthCode(data_assembleia_retrasada)

        self.logger = logger
        self.cwd = cwd
        self.login = login
        self.senha = senha
        self.chrome_driver_path = endereco_chrome_driver
        self.browser, self.wait = self.StartBrowser()

        # 1.0 # VERIFICANDO VERSÃO SELENIUM
        self.logger.warning("[ConsorcioBB] Versão Selenium Em Execução (V" + webdriver.__version__ + ").")

    # --- FUNÇÕES PARA GERAR GRUPOS ATIVOS --- #
    def ReturnsDataFrameWithActiveGroups(self, sigla):
        try:
            # 1.0 # Entrando em Simulador/Contratação e selecionando a categoria.
            self.browser.execute_script("document.getElementById('ctl00_Conteudo_rptFormularios_ctl08_lnkFormulario').click();")
            self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_Conteudo_cbxTipoGrupo"]')))
            Select(self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoGrupo"]')))).select_by_index(0)
            if sigla == "IM240":
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoGrupo"]/option[3]'))).click()
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]'))).click()
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]/option[2]'))).click()
            elif sigla == "IMP":
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoGrupo"]/option[3]'))).click()
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]'))).click()
                self.wait.until(EC.element_to_be_clickable((By.XPATH,'//*[@id="ctl00_Conteudo_cbxTipoVenda"]/option[3]'))).click()
            else:
                self.browser.execute_script(f"document.querySelector('#ctl00_Conteudo_cbxTipoGrupo option[value=\"{sigla}\"]').selected = true")
            self.browser.execute_script("document.getElementById('ctl00_Conteudo_lnkProximo').click();")
            #
            
            # 1.2 # Capturando dados e gerando Data Frame dos dados.
            grupos_ativos = [] # df para guardar dicionário de dados.
            cod = 1 # Variavel de controle para encontrar ultima página da tabela.
            while cod != 0:
                tabela = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_Conteudo_grdGruposDisponiveis"]')))
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')
                just_one_table = True
                for linha in linhas: # Verifica se tabela tem botão 'next' para assim saber que é uma unica página.
                    if linha.find_elements(By.XPATH, './/input[contains(@src, "next.png")]'):
                        just_one_table = False
                for linha in linhas: # Realiza a captura dos dados da tabela.
                    if linha.find_elements(By.XPATH, './/input[contains(@src, "next.png")]'): # Se for a linha do botão 'Next' vai para proxima página de dados da tabela.
                        self.browser.execute_script("__doPostBack('ctl00$Conteudo$grdGruposDisponiveis','Page$Next')")
                        break 
                    elif linha.find_elements(By.XPATH,'.//input[contains(@src, "back.png")]'): # Se achar somente o botão 'Back' é a última página da tabela.
                        cod = 0
                        break
                    elif len(linha.find_elements(By.TAG_NAME, 'td')) > 3: # Verifica se é a linha com dados para ser capturado.
                        coluna = linha.find_elements(By.TAG_NAME, 'td')
                        codigo_grupo = coluna[0].text
                        prazo = int(coluna[1].text)
                        vagas = int(coluna[2].text)
                        txAdm = float(coluna[4].text.replace(',', '.'))  # Converte string para float
                        fReserv = float(coluna[5].text.replace(',', '.'))  # Converte string para float
                        valorBem = float(coluna[7].text.replace('.', '').replace(',', '.'))  # Converte string para float

                        # Calculando novas colunas
                        taxas = f"{txAdm:.2f}% + {fReserv:.2f}%"
                        calcTxAdmMensal = f"{(txAdm     / prazo):.3f}% A.M {(txAdm / prazo * 12):.3f}% A.A"
                        calcFRMensal    = f"{(fReserv   / prazo):.3f}% A.M {(fReserv / prazo * 12):.3f}% A.A"
                        calcTotalMensal = f"{((txAdm    + fReserv) / prazo):.3f}% A.M {((txAdm + fReserv) / prazo * 12):.3f}% A.A"

                        cartas_Grupo = {
                            "Modalidade": self.AcronymTranslator(sigla),
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
                if just_one_table == True: # Se página não tem botão 'next', nem botão 'back' encerra a procura de dados na tabela.
                    cod = 0
            #

            # 1.3 # Criando Data Frame, Pegando somente a linha com menor valor de Carta de Crédito e retornando.
            self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_lnkHome"]'))).click() # Voltando para Menu.
            df = pd.DataFrame(grupos_ativos) # Transformando dicionário em data frame.
            del grupos_ativos
            df = df.loc[df.groupby(['Grupo', 'Prazo'])['CartasCredito'].idxmin()]
            self.logger.info(f"[ConsorcioBB] Data frame Grupos Ativos [{sigla}] extraído com sucesso. ")
            return df
            #
        
        except Exception as err:
            self.logger.error(f"[ConsorcioBB] Falha ReturnsDataFrameWithActiveGroups()\n {err}")
            self.TakeBrowserScreenshot()
            raise

    def AcronymTranslator(self, sigla):
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

    def ReturnsMonthCode(self, date_str):
        # Converte a string de data para um objeto datetime
        date_obj = datetime.datetime.strptime(date_str, "%d/%m/%Y")
        # Extrai o número do mês
        month_number = date_obj.month
        # Formata a string 'M'
        month_code = f'M{month_number}'
        return month_code

    def ReturnsDataFrameAssemblyData(self, lista_grupos, sigla):
        try:
            # 1.1 # Navegando até página 'Resultado de Assembleias'. 
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_rptFormularios_ctl27_lnkFormulario"]'))).click()
            dados_grupos = []
            #

            # 1.2 # Capturando informações das últimas assembleias.
            for grupo in lista_grupos:
                script = f"document.evaluate('//*[@id=\"ctl00_Conteudo_edtCD_Grupo\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{str(grupo).zfill(6)}';"
                self.browser.execute_script(script) # Preenche input com código do grupo.
                script = "document.getElementById('ctl00_Conteudo_lnkPesquisar').click();"
                self.browser.execute_script(script) # Clica no botão pesquisar.
                script = "return document.getElementById('ctl00_Conteudo_lblVolumeTotal').textContent;"
                liquidez = self.browser.execute_script(script) # Pegando valor da 'Label' Volume Total.
                # Definindo nome dos dados no dicionario / Nome colunas no xlsx.
                dados_assembleias_grupo = {
                    "Grupo": int(grupo),
                    "Liquidez": liquidez,
                    f"N° Assembleia {self.sigla_assembleia_mais_recente}": None, f"Contemplados {self.sigla_assembleia_mais_recente}": None,  f"Menor Lance {self.sigla_assembleia_mais_recente}": None,
                    f"N° Assembleia {self.sigla_assembleia_passada}": None, f"Contemplados {self.sigla_assembleia_passada}": None,  f"Menor Lance {self.sigla_assembleia_passada}": None,
                    f"N° Assembleia {self.sigla_assembleia_retrasada}": None, f"Contemplados {self.sigla_assembleia_retrasada}": None,  f"Menor Lance {self.sigla_assembleia_retrasada}": None,
                }

                tabela = self.wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_Conteudo_grdAssembleias"]')))
                linhas = tabela.find_elements(By.TAG_NAME, 'tr')[1:]
                for linha in linhas:
                    coluna = linha.find_elements(By.TAG_NAME, 'td')
                    assembleia = coluna[1].text
                    data = coluna[2].text
                    qtde_contemp = coluna[3].text

                    if coluna[5].text == None or coluna[5].text == " ":
                        menor_lance = "0"
                    else:
                        menor_lance = coluna[5].text

                    # Verificando de qual assembleia são os dados da linha correspondente.
                    if data == self.data_assembleia_mais_recente or data == "28/06/2024": # Assembleia Extraordinaria
                        dados_assembleias_grupo[f"N° Assembleia {self.sigla_assembleia_mais_recente}"] = assembleia
                        dados_assembleias_grupo[f"Contemplados {self.sigla_assembleia_mais_recente}"] = qtde_contemp
                        dados_assembleias_grupo[f"Menor Lance {self.sigla_assembleia_mais_recente}"] = "{:.2f}%".format(float(menor_lance.replace(",", "."))) 
                    elif data == self.data_assembleia_passada:
                        dados_assembleias_grupo[f"N° Assembleia {self.sigla_assembleia_passada}"] = assembleia
                        dados_assembleias_grupo[f"Contemplados {self.sigla_assembleia_passada}"] = qtde_contemp
                        dados_assembleias_grupo[f"Menor Lance {self.sigla_assembleia_passada}"] = "{:.2f}%".format(float(menor_lance.replace(",", "."))) 
                    elif data == self.data_assembleia_retrasada:
                        dados_assembleias_grupo[f"N° Assembleia {self.sigla_assembleia_retrasada}"] = assembleia
                        dados_assembleias_grupo[f"Contemplados {self.sigla_assembleia_retrasada}"] = qtde_contemp
                        dados_assembleias_grupo[f"Menor Lance {self.sigla_assembleia_retrasada}"] = "{:.2f}%".format(float(menor_lance.replace(",", "."))) 
                    
                    #
                dados_grupos.append(dados_assembleias_grupo)

            self.browser.find_element(By.XPATH,'//*[@id="ctl00_lnkHome"]').click() # Voltando para menu Portal Parceiros.
            df = pd.DataFrame(dados_grupos)
            self.logger.info(f"[ConsorcioBB] Dados das Assembleias [{sigla}] Extraídos com Sucesso.")
            return df
        
        except Exception as err:
            self.logger.error(f"\n[ConsorcioBB] Falha ReturnsDataFrameAssemblyData.\n{err}")
            self.TakeBrowserScreenshot()
            raise

    # --- FUNÇÕES PARA CALCULAR MEDIA DE CONTEMPLAÇÃO DOS GRUPOS --- #

    def ReturnsDataFrameGroupsMedia(self, df_dadosAssembleia, lista_grupos, sigla):
        try:
            dados_grupos = []  # Adiciona todos os dados dos grupos.

            for grupo in lista_grupos:
                try:
                    # Chama o método de execução para processar o grupo
                    dados_assembleias_grupo = self.ReturnAveragesAssemblies(grupo, df_dadosAssembleia)
                    dados_grupos.append(dados_assembleias_grupo)
                except Exception as err:
                    self.logger.warning(f"[ConsorcioBB] Tentando novamente para o grupo {grupo} devido a uma falha de execução.")
                    try:
                        # Tenta executar novamente o método de execução para o grupo
                        dados_assembleias_grupo = self.ReturnAveragesAssemblies(grupo, df_dadosAssembleia)
                        dados_grupos.append(dados_assembleias_grupo)
                    except Exception as err:
                        self.logger.warning(f"[ConsorcioBB] Média do grupo {grupo} não pode ser gerada.")
                        continue  # Ignora o grupo se falhar novamente

            df = pd.DataFrame(dados_grupos)
            self.browser.find_element(By.XPATH, '//*[@id="ctl00_lnkHome"]').click()
            self.logger.info(f"[ConsorcioBB] Média contemplações [{sigla}] extraídas com sucesso.")
            return df

        except Exception as err:
            self.logger.error(f"[ConsorcioBB] Erro geral: {str(err)}")
            raise

    def ReturnAveragesAssemblies(self, grupo, df_dadosAssembleia):
        dados_assembleias_grupo = {
            "Grupo": int(grupo),
            f"Media Lance {self.sigla_assembleia_mais_recente}": None,
            f"Media Lance {self.sigla_assembleia_passada}": None,
            f"Media Lance {self.sigla_assembleia_retrasada}": None
        }
        
        try:
            # Navega até assembleia.
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_rptFormularios_ctl27_lnkFormulario"]'))).click()
            # Insere o grupo.
            self.browser.execute_script(f"document.evaluate('//*[@id=\"ctl00_Conteudo_edtCD_Grupo\"]', document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '{str(grupo).zfill(6)}';")
            # Pesquisa dados do grupo.
            self.browser.execute_script("document.getElementById('ctl00_Conteudo_lnkPesquisar').click();")
            
            # Acha a tabela.
            time.sleep(1)
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
                data = celulas[2].text
                buton_CotasContempladas = celulas[6].find_element(By.TAG_NAME, 'input')
                media_lance = None
                
                # ABRINDO ABA DA ASSEMBLEIA PARA CALCULAR MEDIA LANCE
                buton_CotasContempladas.click()
                media_lance = self.CalculateMediaAssembly(grupo, df_dadosAssembleia)

                if data == self.data_assembleia_mais_recente:
                    dados_assembleias_grupo[f"Media Lance {self.sigla_assembleia_mais_recente}"] = f"{float(media_lance):.2f}".replace('.', ',') + '%'
                elif data == self.data_assembleia_passada:
                    dados_assembleias_grupo[f"Media Lance {self.sigla_assembleia_passada}"] = f"{float(media_lance):.2f}".replace('.', ',') + '%'
                elif data == self.data_assembleia_retrasada:
                    dados_assembleias_grupo[f"Media Lance {self.sigla_assembleia_retrasada}"] = f"{float(media_lance):.2f}".replace('.', ',') + '%'
                
                # Voltando para menu principal.
                self.browser.find_element(By.XPATH, '//*[@id="ctl00_lnkHome"]').click()
                execution += 1

        except Exception as err:
            self.browser.find_element(By.XPATH,'//*[@id="ctl00_lnkHome"]').click() # Voltando para menu Portal Parceiros.
            self.logger.error(f"[ConsorcioBB] Erro ao processar grupo {grupo}")
            raise  # Relança a exceção para ser tratada no método principal

        return dados_assembleias_grupo

    def CalculateMediaAssembly(self, grupo, df_dadosAssembleia):
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
                menor_lance_recente = dados_grupo[f'Menor Lance {self.sigla_assembleia_mais_recente}'].values[0]
                menor_lance_passada = dados_grupo[f'Menor Lance {self.sigla_assembleia_passada}'].values[0]
                menor_lance_retrasada = dados_grupo[f'Menor Lance {self.sigla_assembleia_retrasada}'].values[0]

                for CotaContemplada in list_cotasContempladas:

                    percentual = CotaContemplada['Percentual']
                    if percentual == '0,0000': # Desconsidera sorteio.
                        continue
                    elif percentual == "20,0000" and ( menor_lance_passada == '20.00%' or None) and (menor_lance_retrasada == '20.00%' or None) and (menor_lance_recente == '20.00%' or None):
                        # lanceFixo20 = True
                        continue
                    elif percentual == "30,0000" and (menor_lance_passada == '30.00%' or None) and (menor_lance_retrasada == '30.00%' or None) and (menor_lance_recente == '30.00%' or None):
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
                self.logger.info(f"[WARINING] Não há valores válidos para calcular a média. Grupo {grupo}°")
                return media_percentuais == 0

            self.browser.find_element(By.XPATH,"//*[@id='ctl00_Conteudo_lnkFechaDiv']").click()
            return media_percentuais

        except Exception as err:
            self.logger.error(f"error\n{err}")
            raise

    # --- FUNÇÕES PADRÕES -- #
    def StartBrowser(self):
        try:
            self.logger.info("[ConsorcioBB] Iniciando Chrome...")
            chrome_service = Service(self.chrome_driver_path)
            browser = webdriver.Chrome(service=chrome_service)
            wait = WebDriverWait(browser, 10)
            return browser, wait
        except Exception as err:
            self.logger.error(f"\n[ConsorcioBB] Falha StartBrowser.\n{err}")
            raise
    
    def Login(self):
        try:
            self.logger.info("[ConsorcioBB] Acessando URL (https://www.parceirosbbconsorcios.com.br).")
            self.browser.get('https://www.parceirosbbconsorcios.com.br/')
            self.logger.info("[ConsorcioBB] Efetuando login PP BB Consórcio.")
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_edtUsuario"]'))).send_keys(self.login)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_edtSenha"]'))).send_keys(self.senha)
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_lnkEntrar"]'))).click()
            self.logger.info("[ConsorcioBB] Logado com sucesso no PP BB Consórcio.")
        except Exception as err:
            self.logger.error(f"\n[ConsorcioBB] Falha Login.\n{err}")
            raise
    
    def ClosePoupUp(self):
        try:
            self.logger.info('[ConsorcioBB] Fechando Poup-Up Inicial.')
            self.wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_Conteudo_LinkButton1"]'))).click()
            self.logger.info("[ConsorcioBB] Poup-Up fechado.")
        except Exception as err:
            self.logger.error(f"\n[ConsorcioBB] Falha no metodo ClosePoupUp.\n{err}")
            raise
    
    def EndBrowser(self):
        try:
            # Encerra navegador.
            self.logger.info("[ConsorcioBB] Encerrando Chrome.")
            self.browser.quit()
        except Exception as err:
            self.logger.error(f"\n[ConsorcioBB] Falha no método EndBrowser.\n{err}")
            raise

    def TakeBrowserScreenshot(self):
        try:
            file_path = f"{self.cwd}error.png"
            # Tira a captura de tela e salva no caminho especificado
            self.browser.save_screenshot(file_path)
            self.logger.info(f"[ERROR] Captura de tela do erro salva em [{file_path}]")
            return True
        except Exception as e:
            self.logger.error(f"[ConsrocioBB] Falha TakeBrowserScreenshot\n {str(e)}")
            raise