try:
    # MELHORIAS PARA A PRÓXIMA VERSÃO #
    # VERIFICAÇÃO DAS DATAS DA ASSEMBLEIA PARA GARANTIR QUE ESTEJAM CERTAS.
    # IDENTIFICAÇÃO DE QUANTAS CONTEMPLAÇÕES POR SORTEIO E QUANTAS POR LANCE (1449 / 1447)

    # 0.0 # IMPORTAÇÕES NECESSÁRIAS PARA EXECUTAR O PROJETO.
    import sys
    import os
    import pandas as pd
    from datetime import datetime
    from ConsorcioBB import ConsorcioBB
    from Logger import Logger

    # 1.0 - VARIAVEIS GLOBAIS. ------------------------------------------------------------------------------------------------------
    executionCount_path = str(os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))+'\\Builds\\execution_count.txt')
    endereco_chrome_driver = str(os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))+'\\chromedriver-win64\\chromedriver.exe')
    global execution_code
    global currentExecution_directory

    # 1.2 - VARIAVEIS DE EXECUÇÃO. ------------------------------------------------------------------------------------------------------
    dados_assembleia, info_assembleia = True, True 
    categoria_processadas = ["AU", "MO"] # TODAS AS CATEGORIAS POSSIVEIS -> "IMP","IM240", "TC", "AI", "AU", "MO", "EE"
    data_assembleia_mais_recente = input("Digite a data da assembleia mais recente: ") if dados_assembleia or info_assembleia else None
    data_assembleia_passada = input("Digite a data da assembleia passada: ") if dados_assembleia or info_assembleia else None
    data_assembleia_retrasada = input("Digite a data da assembleia retrasada: ") if dados_assembleia or info_assembleia else None
    login = input("Digite o seu login: ")
    senha = input("Digite a sua senha: ")
   
    # 1.3 - LENDO LOG COUNT.TXT 
    with open(executionCount_path, 'r') as arquivo:
        conteudo = arquivo.read()
        execution_code = int(conteudo)
    
    # 1.4 - CRIA O DIRETÓRIO PARA EXECUÇÃO DA BUILD ATUAL.
    currentExecution_directory=f'.\\Builds\\Build Number {execution_code}'
    os.makedirs(currentExecution_directory, exist_ok=True)  # Cria o diretório se não existir.
    open(os.path.join(currentExecution_directory, f'Build {execution_code}.log'), 'a').close()

    # 1.5 - INICIALIZANDO CLASSE LOGGER.
    logger = Logger(execution_code, write_log_directory=f'{currentExecution_directory}\\Build {execution_code}.log')
    datetime_start_execution = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    logger.info(f"[Executar] Iniciando execução [{execution_code}] datetime [{datetime_start_execution}]")

    # -- INICIO EXECUÇÃO DA CLASSE ConsorcioBB.py   -- #
    # 1.6 - EXECUÇÃO.
    try:
        # 1.2 # EXTRAINDO DADOS REFENTE A CADA CATEGORIA DE GRUPOS.
        for sigla in categoria_processadas:
            # 1.0 # INSTÂNCIANDO CLASSE CONSORCIO BB.
            consorcio = ConsorcioBB(
                    login=login,
                    senha=senha,
                    endereco_chrome_driver=endereco_chrome_driver,
                    logger=logger,
                    cwd=f'.\\Builds\\Build Number {execution_code}\\Build {execution_code}.log',
                    data_assembleia_mais_recente=data_assembleia_mais_recente,
                    data_assembleia_passada=data_assembleia_passada,
                    data_assembleia_retrasada=data_assembleia_retrasada
            )
            
            # 1.1 # INICIANDO NAVEGADOR, FAZENDO LOGON E FECHANDO POUP-UP.
            consorcio.Login()
            consorcio.ClosePoupUp()

            # 1.0 - DADOS GRUPOS ATIVOS.
            df_todosGruposAtivos = consorcio.ReturnsDataFrameWithActiveGroups(sigla= str(sigla) )
            lista_id_grupos =  [codigo for codigo in df_todosGruposAtivos['Grupo'].tolist() if codigo != 0] # [1453, 1526, 1520]

            # 1.1 - DADOS ASSEMBLEIA.
            if dados_assembleia or info_assembleia:
                df_dados_assembleia = consorcio.ReturnsDataFrameAssemblyData(lista_grupos=lista_id_grupos, sigla=sigla)
                df_dados_assembleia = pd.merge(df_todosGruposAtivos, df_dados_assembleia, on="Grupo", how="left")
            # 1.1 - DECIDE QUAL DADOS VAI CONTER NO EXCEL GERADO.
            # 1° CONDIÇÃO - GERA EXCEL COM DADOS -> (GRUPOS ATIVOS, DADOS DA ASSEMBLEIA, INFORMACOES DE CONTEMPLACAO P/ GRUPO).
            if info_assembleia:
                # 1.1.1 - CÁLCULO INFORMACOES DE CONTEMPLAÇÃO POR ASSEMBLEIA.
                df_mediaContemplacao = consorcio.ReturnsDataFrameGroupsInfo(df_dadosAssembleia=df_dados_assembleia, lista_grupos=lista_id_grupos, sigla=sigla)
                df_excelFinal = pd.merge(df_dados_assembleia, df_mediaContemplacao, on="Grupo", how="left")
                
                # 1.1.2 - ORGANIZA COLUNAS E SALVA EXCEL COM DADOS.
                colunas = [
                    "Modalidade", "Grupo", "Prazo", "Vagas", "Taxas", "Calculo TxAdm", "Calculo FR",
                    "Calculo Total", "CartasCredito", "Liquidez", f"N° Assembleia {consorcio.sigla_assembleia_mais_recente}",
                    f"Contemplados {consorcio.sigla_assembleia_mais_recente}",  f"Qtde Cotas LL {consorcio.sigla_assembleia_mais_recente}", f"Qtde Cotas LF {consorcio.sigla_assembleia_mais_recente}",    f"Qtde Cotas Sorteio {consorcio.sigla_assembleia_mais_recente}", f"Media Lance {consorcio.sigla_assembleia_mais_recente}",   f"Menor Lance {consorcio.sigla_assembleia_mais_recente}",
                    f"Contemplados {consorcio.sigla_assembleia_passada}",       f"Qtde Cotas LL {consorcio.sigla_assembleia_passada}",      f"Qtde Cotas LF {consorcio.sigla_assembleia_passada}"     ,    f"Qtde Cotas Sorteio {consorcio.sigla_assembleia_passada}"     , f"Media Lance {consorcio.sigla_assembleia_passada}"     ,   f"Menor Lance {consorcio.sigla_assembleia_passada}",
                    f"Contemplados {consorcio.sigla_assembleia_retrasada}",     f"Qtde Cotas LL {consorcio.sigla_assembleia_retrasada}",    f"Qtde Cotas LF {consorcio.sigla_assembleia_retrasada}"   ,    f"Qtde Cotas Sorteio {consorcio.sigla_assembleia_retrasada}"   , f"Media Lance {consorcio.sigla_assembleia_retrasada}"   ,   f"Menor Lance {consorcio.sigla_assembleia_retrasada}"
                ]
                # 1.1.3 - SALVANDO ARQUIVO FINAL ORGANIZDO.
                df_excelFinal[colunas].to_excel(f'{currentExecution_directory}\\TodosGruposAtivos [{sigla}] BUILD({execution_code}).xlsx', index=False)
                logger.info(f"[Executar] Arquivo 'TodosGruposAtivos [{sigla}] BUILD({execution_code}).xlsx' criado com sucesso.")
                consorcio.EndBrowser()        
            # 2° CONDIÇÃO - GERA EXCEL COM DADOS -> (GRUPOS ATIVOS).
            elif not dados_assembleia and not info_assembleia:
                # 1.2 - SALVANDO ARQUIVO FINAL ORGANIZDO.
                df_todosGruposAtivos.to_excel(f'{currentExecution_directory}\\TodosGruposAtivos [{sigla}] BUILD({execution_code}).xlsx', index=False)
                logger.info(f"[Executar] Arquivo 'TodosGruposAtivos [{sigla}] BUILD({execution_code}).xlsx' criado com sucesso.")
                consorcio.EndBrowser()
            # 3° CONDIÇÃO - GERA EXCEL COM DADOS -> (GRUPOS ATIVOS, DADOS DA ASSEMBLEIA).
            else:
                # 1.2 - SALVANDO ARQUIVO FINAL ORGANIZDO.
                df_dados_assembleia.to_excel(f'{currentExecution_directory}\\Dados Assembleia [{sigla}] BUILD({execution_code}).xlsx', index=False)
                logger.info(f"[Executar] Arquivo 'Dados Assembleia [{sigla}] BUILD({execution_code}).xlsx' criado com sucesso.")
                consorcio.EndBrowser() 
    except Exception as err:
        # 1.2.1 # TRATANDO EXECUÇÕES COM ERROS.
        print(f"[Executar] FALHA AO EXECUTAR CLASSE ConsorcioBB.py.-> \n{err}")
        consorcio.EndBrowser()
        raise
    # -- FIM EXECUÇÃO DA CLASSE ConsorcioBB.py -- #

    # 1.7 - ATUALIZANDO LOG COUNT.TXT
    logger.info(f"[Executar] Execução Finalizada [{datetime.now().strftime("%H:%M:%S")}] || Finalized Execution [{execution_code}]")
    with open(r'.\\Builds\\execution_count.txt', 'w') as arquivo:
        arquivo.write(str(execution_code + 1))

except Exception as err:
    # 0.0 - MENSAGEM DE ERROR.
    if 'logger' in locals() and hasattr(logger, 'error'):
        logger.error(f"[Executar] Falha ao executar programa.\nERROR-> {err}")
    
    # 0.1 - INSERINDO MENSAGEM NO LOG E ATUALIZANDO LOG COUNT.
    logger.warning(f"[Executar] Execução Finalizada [{datetime.now().strftime("%H:%M:%S")}] || Finalized Execution [{execution_code}]")
    with open(executionCount_path, 'w') as arquivo:
        arquivo.write(str(execution_code + 1))
        print("[Executar] Arquivo de controle de log (execution_count.txt) encerrado com sucesso.")
    sys.exit(1)