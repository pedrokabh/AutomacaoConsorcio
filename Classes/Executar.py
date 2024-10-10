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
    parte1 = True # GERA TODOS DADOS GRUPOS ATIVOS E ASSEMBLEIAS.
    dados_assembleia, media_assembleia = False, False 
    # DADOS ASSEMBLEIA - RETORNA INFORMAÇÕES COMO COTAS CONTEMPLADAS, MENOR LANCE, MAIOR LANCE, ETC.
    # MEDIA ASSEMBLEIA - CALCULA A MEDIA DE CONTEMPLAÇÃO DAS ASSEMBLEIAS.

    parte2 = False # REFERENTE A VENDAS EM RELAÇÃO AS ULTIMAS ASSEMBLEIAS.

    # 1.3 - VARIAVEIS PARA EXECUTAR PARTE 1.
    if parte1:
        categoria_processadas = ["TC", "AI", "AU", "MO", "EE", "IM240", "IMP"]
        login = input("Digite o seu login: ")
        senha = input("Digite a sua senha: ")
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

    # -- INICIO EXECUÇÃO DA CLASSE ConsorcioBB.py   --
    # 1.6 - EXECUÇÃO.
    if parte1:
        try:
            ## --- EXECUÇÃO PARTE 1 --- ##
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

            # 1.1 - CALCULAR MÉDIA COMTEMPLAÇÕES
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
            elif not df_dados_assembleia and not df_dados_assembleia:
                # 1.2 - SALVANDO ARQUIVO FINAL ORGANIZDO.
                df_todosGruposAtivos.to_excel(f'{currentExecution_directory}\\TodosGruposAtivos BUILD({execution_code}).xlsx', index=False)
                logger.info(f"[Executar] Arquivo 'TodosGruposAtivos BUILD({execution_code}).xlsx' criado com sucesso.")
                consorcio.EndBrowser()
            else:
                # 1.2 - SALVANDO ARQUIVO FINAL ORGANIZDO.
                df_dados_assembleia.to_excel(f'{currentExecution_directory}\\Dados Assembleia BUILD({execution_code}).xlsx', index=False)
                logger.info(f"[Executar] Arquivo 'Dados Assembleia BUILD({execution_code}).xlsx' criado com sucesso.")
                consorcio.EndBrowser()

        except Exception as err:
            print(f"[Executar] FALHA AO EXECUTAR PARTE 1.\n{err}")
            sys.exit(1)

     #ADICIONA QUANTIDADE DE VENDAS NO ARQUIVO E GERA O 'TODOSGRUPOSATIVOSCOMVENDAS.XLSX'
    if parte2:
        try:
            # Ler os arquivos Excel
            todos_grupos_ativos = pd.read_excel(
                f'.\\Builds\\Build Number {execution_code}\\TodosGruposAtivos.xlsx',
                sheet_name='Sheet1')
            vagasM5 = pd.read_excel(r'.\\Classes\\VagasM5.xlsx', sheet_name='Planilha1')
            vagasM6 = pd.read_excel(r'.\\Classes\\VagasM6.xlsx', sheet_name='Planilha1')
            vagasM7 = pd.read_excel(r'.\\Classes\\VagasM7.xlsx', sheet_name='Planilha1')

            # Selecionar as colunas relevantes dos arquivos de vagas
            vagasM5 = vagasM5[['Grupo', 'Vagas M5']]
            vagasM6 = vagasM6[['Grupo', 'Vagas M6']]
            vagasM7 = vagasM7[['Grupo', 'Vagas M7']]

            # Mesclar os dados com base na coluna 'Grupo'
            merged_data = pd.merge(todos_grupos_ativos, vagasM5, on='Grupo', how='left')
            merged_data = pd.merge(merged_data, vagasM6, on='Grupo', how='left')
            merged_data = pd.merge(merged_data, vagasM7, on='Grupo', how='left')

            # Garantir que os valores ausentes fiquem em branco
            merged_data['Vagas M5'] = pd.to_numeric(merged_data['Vagas M5'], errors='coerce').fillna(0)
            merged_data['Vagas M6'] = pd.to_numeric(merged_data['Vagas M6'], errors='coerce').fillna(0)
            merged_data['Vagas M7'] = pd.to_numeric(merged_data['Vagas M7'], errors='coerce').fillna(0)
            merged_data['Vagas'] = pd.to_numeric(merged_data['Vagas'], errors='coerce').fillna(0)

            # Calcular as colunas 'Vendas M6', 'Vendas M7' e 'Vendas Atuais'
            merged_data['Vendas M6'] = merged_data['Vagas M5'] - merged_data['Vagas M6']
            merged_data['Vendas M7'] = merged_data['Vagas M6'] - merged_data['Vagas M7']
            merged_data['Vendas Atuais'] = merged_data['Vagas M7'] - merged_data['Vagas']
            # merged_data = merged_data.drop(columns=['Vagas M5', 'Vagas M6', 'Vagas M7'])

            # Salvar o novo arquivo atualizado
            merged_data.to_excel(f'.\\Builds\\Build Number {execution_code}\\TodosGruposAtivosComVendas.xlsx', index=False)
            logger.warning(f"[Executar] ARQUIVO 'TodosGruposAtivosComVendas.xlsx' GERADO COM SUCESSO.")
        except Exception as err:
            print(err)
    # -- FIM EXECUÇÃO DA CLASSE ConsorcioBB.py      --

    # 1.7 - ATUALIZANDO LOG COUNT.TXT
    logger.info(f"[Executar] Execução Finalizada [{datetime.now().strftime("%H:%M:%S")}] || Finalized Execution [{execution_code}]")
    with open(r'.\\Builds\\execution_count.txt', 'w') as arquivo:
        arquivo.write(str(execution_code + 1))

except Exception as err:
    # MENSAGEM DE ERROR.
    if 'logger' in locals() and hasattr(logger, 'error'):
        logger.error(f"\n[Executar] Falha ao executar programa.\nERROR-> {err}")
    
    # INSERINDO MENSAGEM NO LOG E ATUALIZANDO LOG COUNT.
    logger.warning(f"[Executar] Execução Finalizada [{datetime.now().strftime("%H:%M:%S")}] || Finalized Execution [{execution_code}]")
    with open(str('.\\execution_count.txt'), 'w') as arquivo:
        arquivo.write(str(execution_code + 1))
        print("[Executar] Arquivo de controle de log (execution_count.txt) encerrado com sucesso.")
    sys.exit(1)
