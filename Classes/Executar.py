import sys
import os
import pandas as pd
from datetime import datetime
from ConsorcioBB import ConsorcioBB
from Logger import Logger

try:
    """
        !!! ATENÇÃO !!! - Caso seja outra assembleia, atualizar as colunas do mês.
    """
    # 1.0 - VARIAVEIS GLOBAIS.
    executionCount_path = str(os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))+'\\Builds\\execution_count.txt')
    endereco_chrome_driver = r"..\chromedriver-win64\chromedriver.exe"
    execution_code = None
    currentExecution_directory = None

    # 1.2 - VARIAVEIS DE EXECUÇÃO.
    dados_assembleia, media_assembleia = False, False 
    # dados_assembleia, media_assembleia = False, False -> Extraí todos grupos ativos.
    # dados_assembleia, media_assembleia = True , False -> Extraí todos grupos ativos e dados assembleia.
    # dados_assembleia, media_assembleia = True , True  -> Extraí todos grupos ativos, dados assembleia, média contemplação grupo.

    # 1.3 - VARIAVEIS PARA EXECUTAR PARTE 1.
    categoria_processadas = ["TC", "AI", "AU", "MO", "EE", "IM240", "IMP"]
    login = input("Digite o seu login: ")
    senha = input("Digite a sua senha: ")
    if dados_assembleia or media_assembleia:
        data_assembleia_passada = "27/08/2024"      # data_assembleia_mais_recente = input("Digite a data da assembleia mais recente: ")    
        data_assembleia_retrasada = "26/07/2024"    # data_assembleia_passada = input("Digite a data da assembleia passada: ")
        data_assembleia_mais_recente = "25/09/2024" # data_assembleia_retrasada = input("Digite a data da assembleia retrasada: ")

    # 1.4 - LENDO LOG COUNT.TXT
    with open(executionCount_path, 'r') as arquivo:
        conteudo = arquivo.read()
        execution_code = int(conteudo)
    
    # 1.5 - CRIA O DIRETÓRIO PARA EXECUÇÃO DA BUILD ATUAL.
    currentExecution_directory=f'.\\Builds\\Build Number {execution_code}'
    os.makedirs(currentExecution_directory, exist_ok=True)  # Cria o diretório se não existir.
    open(os.path.join(currentExecution_directory, f'Build {execution_code}.log'), 'a').close()

    # 1.6 - INICIALIZANDO CLASSE LOGGER.
    logger = Logger(execution_code, write_log_directory=f'{currentExecution_directory}\\Build {execution_code}.log')
    datetime_start_execution = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    logger.info(f"[Executar] Iniciando execução [{execution_code}] datetime [{datetime_start_execution}]")

    # -- INICIO EXECUÇÃO DA CLASSE ConsorcioBB.py   -- #
    # 1.6 - EXECUÇÃO.
    try:
        ## --- EXECUÇÃO PARTE --- ##
        consorcio = ConsorcioBB(
            login=login,
            senha=senha,
            endereco_chrome_driver=endereco_chrome_driver,
            data_assembleia_mais_recente=data_assembleia_mais_recente,
            data_assembleia_passada=data_assembleia_passada,
            data_assembleia_retrasada=data_assembleia_retrasada,
            logger=logger,
            cwd=f'.\\Builds\\Build Number {execution_code}\\Build {execution_code}.log'
        )

        # 1.0 - DADOS GRUPOS ATIVOS.
        df_todosGruposAtivos = [consorcio.ReturnsDataFrameWithActiveGroups(sigla=categoria) for categoria in categoria_processadas]
        df_todosGruposAtivos = pd.concat(df_todosGruposAtivos, ignore_index=True)
        lista_id_grupos = [codigo for codigo in df_todosGruposAtivos['Grupo'].tolist() if codigo != 0]

        # 1.1 - DADOS ASSEMBLEIA.
        if dados_assembleia or media_assembleia:
            df_dados_assembleia = consorcio.ReturnsDataFrameAssemblyData(lista_grupos=lista_id_grupos)
            df_dados_assembleia = pd.merge(df_todosGruposAtivos, df_dados_assembleia, on="Grupo", how="left")

        # 1.1 - DECIDE QUAL DADOS VAI CONTER NO EXCEL GERADO.
        # 1° CONDIÇÃO - GERA EXCEL COM DADOS -> (GRUPOS ATIVOS, DADOS DA ASSEMBLEIA, MEDIA CONTEMPLACAO P/ GRUPO).
        if media_assembleia:
            # 1.1.1 - CÁLCULO MÉDIA CONTEMPLAÇÃO POR ASSEMBLEIA.
            df_mediaContemplacao = consorcio.ReturnsDataFrameGroupsMedia(df_dadosAssembleia=df_dados_assembleia, lista_grupos=lista_id_grupos)
            df_excelFinal = pd.merge(df_dados_assembleia, df_mediaContemplacao, on="Grupo", how="left")

            # 1.1.2 - ORGANIZA COLUNAS E SALVA EXCEL COM DADOS.
            colunas = [
                "Modalidade", "Grupo", "Prazo", "Vagas", "Taxas", "Calculo TxAdm", "Calculo FR",
                "Calculo Total", "CartasCredito", "Liquidez", f"N° Assembleia {consorcio.sigla_assembleia_mais_recente}",
                f"Contemplados {consorcio.sigla_assembleia_mais_recente}", f"Media Lance {consorcio.sigla_assembleia_mais_recente}", f"Menor Lance {consorcio.sigla_assembleia_mais_recente}",
                f"Contemplados {consorcio.sigla_assembleia_passada}", f"Media Lance {consorcio.sigla_assembleia_passada}", f"Menor Lance {consorcio.sigla_assembleia_passada}",
                f"Contemplados {consorcio.sigla_assembleia_retrasada}", f"Media Lance {consorcio.sigla_assembleia_retrasada}", f"Menor Lance {consorcio.sigla_assembleia_retrasada}"
            ]

            # 1.1.3 - SALVANDO ARQUIVO FINAL ORGANIZDO.
            df_excelFinal[colunas].to_excel(f'{currentExecution_directory}\\TodosGruposAtivos BUILD({execution_code}).xlsx', index=False)
            logger.info(f"[Executar] Arquivo 'TodosGruposAtivos BUILD({execution_code}).xlsx' criado com sucesso.")
            consorcio.EndBrowser()        
        # 2° CONDIÇÃO - GERA EXCEL COM DADOS -> (GRUPOS ATIVOS).
        elif not df_dados_assembleia and not df_dados_assembleia:
            # 1.2 - SALVANDO ARQUIVO FINAL ORGANIZDO.
            df_todosGruposAtivos.to_excel(f'{currentExecution_directory}\\TodosGruposAtivos BUILD({execution_code}).xlsx', index=False)
            logger.info(f"[Executar] Arquivo 'TodosGruposAtivos BUILD({execution_code}).xlsx' criado com sucesso.")
            consorcio.EndBrowser()
        # 3° CONDIÇÃO - GERA EXCEL COM DADOS -> (GRUPOS ATIVOS, DADOS DA ASSEMBLEIA).
        else:
            # 1.2 - SALVANDO ARQUIVO FINAL ORGANIZDO.
            df_dados_assembleia.to_excel(f'{currentExecution_directory}\\Dados Assembleia BUILD({execution_code}).xlsx', index=False)
            logger.info(f"[Executar] Arquivo 'Dados Assembleia BUILD({execution_code}).xlsx' criado com sucesso.")
            consorcio.EndBrowser()
    except Exception as err:
        print(f"[Executar] FALHA AO EXECUTAR CLASSE ConsorcioBB.py.\n{err}")
        sys.exit(1)
    # -- FIM EXECUÇÃO DA CLASSE ConsorcioBB.py -- #

    # 1.7 - ATUALIZANDO LOG COUNT.TXT
    logger.info(f"[Executar] Execução Finalizada [{datetime.now().strftime("%H:%M:%S")}] || Finalized Execution [{execution_code}]")
    with open(r'.\\Builds\\execution_count.txt', 'w') as arquivo:
        arquivo.write(str(execution_code + 1))

except Exception as err:
    # 0.0 - MENSAGEM DE ERROR.
    if 'logger' in locals() and hasattr(logger, 'error'):
        logger.error(f"\n[Executar] Falha ao executar programa.\nERROR-> {err}")
    
    # 0.1 - INSERINDO MENSAGEM NO LOG E ATUALIZANDO LOG COUNT.
    logger.warning(f"[Executar] Execução Finalizada [{datetime.now().strftime("%H:%M:%S")}] || Finalized Execution [{execution_code}]")
    with open(str('.\\execution_count.txt'), 'w') as arquivo:
        arquivo.write(str(execution_code + 1))
        print("[Executar] Arquivo de controle de log (execution_count.txt) encerrado com sucesso.")
    sys.exit(1)
