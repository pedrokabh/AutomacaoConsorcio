import sys
import os
import pandas as pd
from datetime import datetime
from ConsorcioBB import ConsorcioBB
from Logger import Logger

try:

    # LENDO LOG COUNT.TXT
    with open(str(os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))+'\\Classes\\execution_count.txt'), 'r') as arquivo:
        conteudo = arquivo.read()

    """
        !!! ATENÇÃO !!! - Caso seja outra assembleia, atualizar as colunas do mês.
        !!! ATENÇÃO !!! - Caso seja outra assembleia, deve-se extrair as vagas dos grupos na assembleia anterior e gerar o arquivo VagasMX.xlsxx
        e modificar as colunas na mão.
    """
    # !! EXECUTAR DENTRO DO DIRETORIO automacaoconsorcio !!
    # !--- PARAMETROS DE EXECUÇÃO ---! #
    execution_code = int(conteudo)
    parte1 = True # GERA TODOS DADOS GRUPOS ATIVOS E ASSEMBLEIAS.
    parte2 = False # REFERENTE A VENDAS EM RELAÇÃO AS ULTIMAS ASSEMBLEIAS.
    media_assembleia = False # DECIDE EXTRAIR RELATORIO COM OU MEDIA DAS ASSEMBLEIAS

    if parte1:
        categoria_processadas = ["TC", "AI", "AU", "MO", "EE", "IM240", "IMP"]
        endereco_chrome_driver = r"..\chromedriver-win64\chromedriver.exe"
        
        data_assembleia_passada = "27/08/2024" # AGOSTO
        data_assembleia_retrasada = "26/07/2024" # JULHO
        data_assembleia_mais_recente = "25/09/2024" # SETEMBRO
        # data_assembleia_mais_recente = input("Digite a data da assembleia mais recente: ")
        # data_assembleia_passada = input("Digite a data da assembleia passada: ")
        # data_assembleia_retrasada = input("Digite a data da assembleia retrasada: ")
        login = input("Digite o seu login: ")
        senha = input("Digite a sua senha: ")
    # !--- PARAMETROS DE EXECUÇÃO ---! #


    # INICIALIZANDO LOG
    logger = Logger(execution_code, write_log_directory=f'.\\Builds\\Build Number {execution_code}\\Build {execution_code}.log')
    datetime_start_execution = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    logger.info(f"[Executar] Iniciando execução [{execution_code}] datetime [{datetime_start_execution}]")
    
    # GERA ARQUIVO TODOSGRUPOSATIVOS.XLSX (DADOS DOS GRUPOS / MEDIA CONTEMPLAÇÃO / DADOS ASSEMBLEIAS)
    if parte1:
        try:
            os.system('cls')

            # DIRETORIO ARQUIVO LOG.
            directory_mother = f'{os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))}\\Builds\\Build Number {execution_code}\\'
            os.makedirs(os.path.dirname( str(directory_mother) ), exist_ok=True)

            

            ## --- Execução --- ##
            consorcio = ConsorcioBB(
                login=login,
                senha=senha,
                endereco_chrome_driver=endereco_chrome_driver,
                data_assembleia_mais_recente=data_assembleia_mais_recente,
                data_assembleia_passada=data_assembleia_passada,
                data_assembleia_retrasada=data_assembleia_retrasada,
                logger=logger,
                cwd=directory_mother
            )

            # DADOS GRUPOS ATIVOS
            df_gruposAtivos = [consorcio.ReturnsDataFrameWithActiveGroups(sigla=categoria) for categoria in categoria_processadas]
            df_todosGruposAtivos = pd.concat(df_gruposAtivos, ignore_index=True)
            df_todosGruposAtivos.to_excel(f'.\\Builds\\Build Number {execution_code}\\df_todosGruposAtivos.xlsx', index=False)
            logger.info("[Executar] Arquivo 'df_todosGruposAtivos.xlsx' criado com sucesso.")

            # DADOS ASSEMBLIEA.
            lista_codigos = [codigo for codigo in df_todosGruposAtivos['Grupo'].tolist() if codigo != 0]
            df_assembleia = consorcio.ReturnsDataFrameAssemblyData(lista_grupos=lista_codigos)
            df_assembleia.to_excel(f'.\\Builds\\Build Number {execution_code}\\df_dados_assembleia.xlsx', index=False)
            logger.info("[Executar] Arquivo 'df_dados_assembleia.xlsx' criado com sucesso.")

            # CALCULAR MÉDIA COMTEMPLAÇÕES
            if media_assembleia:
                df_mediaContemplacao = consorcio.ReturnsDataFrameGroupsMedia(df_dadosAssembleia=df_assembleia, lista_grupos=lista_codigos)
                df_mediaContemplacao.to_excel(f'.\\Builds\\Build Number {execution_code}\\df_medias_assembleias.xlsx', index=False)
                logger.info("[Executar] Arquivo 'df_medias_assembleias.xlsx' criado com sucesso.")
                df_excelFinal = pd.merge(df_todosGruposAtivos, df_assembleia, on="Grupo", how="left")
                df_excelFinal = pd.merge(df_excelFinal, df_mediaContemplacao, on="Grupo", how="left")
            else:
                df_excelFinal = pd.merge(df_todosGruposAtivos, df_assembleia, on="Grupo", how="left")

            # ORGANIZANDO ORDEM DAS COLUNAS
            if media_assembleia:
                colunas = [
                    "Modalidade", "Grupo", "Prazo", "Vagas", "Taxas", "Calculo TxAdm", "Calculo FR",
                    "Calculo Total", "CartasCredito", "Liquidez", f"N° Assembleia {consorcio.sigla_assembleia_mais_recente}",
                    f"Contemplados {consorcio.sigla_assembleia_mais_recente}", f"Media Lance {consorcio.sigla_assembleia_mais_recente}", f"Menor Lance {consorcio.sigla_assembleia_mais_recente}",
                    f"Contemplados {consorcio.sigla_assembleia_passada}", f"Media Lance {consorcio.sigla_assembleia_passada}", f"Menor Lance {consorcio.sigla_assembleia_passada}",
                    f"Contemplados {consorcio.sigla_assembleia_retrasada}", f"Media Lance {consorcio.sigla_assembleia_retrasada}", f"Menor Lance {consorcio.sigla_assembleia_retrasada}"
                ]
            else:
                colunas = [
                    "Modalidade", "Grupo", "Prazo", "Vagas", "Taxas", "Calculo TxAdm", "Calculo FR",
                    "Calculo Total", "CartasCredito", "Liquidez", f"N° Assembleia {consorcio.sigla_assembleia_mais_recente}",
                    f"Contemplados {consorcio.sigla_assembleia_mais_recente}", f"Menor Lance {consorcio.sigla_assembleia_mais_recente}",
                    f"Contemplados {consorcio.sigla_assembleia_passada}",  f"Menor Lance {consorcio.sigla_assembleia_passada}",
                    f"Contemplados {consorcio.sigla_assembleia_retrasada}", f"Menor Lance {consorcio.sigla_assembleia_retrasada}"
                ]
            
            # SALVANDO ARQUIVO FINAL ORGANIZDO. !
            df_excelFinal = df_excelFinal[colunas]
            df_excelFinal.to_excel(f'.\\Builds\\Build Number {execution_code}\\TodosGruposAtivos.xlsx', index=False)
            logger.info("[Executar] Arquivo 'TodosGruposAtivos.xlsx' criado com sucesso.")
            consorcio.EndBrowser()
            
        except Exception as err:
            print(f"[Executar] FALHA AO CARREGAR IMPORTAÇÕES \n{err}")
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

    # ATUALIZANDO LOG COUNT.TXT
    # with open(str(directory_mother+'\\Classes\\execution_count.txt'), 'w') as arquivo:
    #     arquivo.write(str(execution_code + 1))
    # logger.info(f"[Executar] Execução Finalizada [{datetime.now().strftime("%H:%M:%S")}] | new execution count [{execution_code + 1}]")    

except Exception as err:
    if 'logger' in locals() and hasattr(logger, 'error'):
        logger.error(f"\n[Executar] Falha ao executar programa.\nERROR-> {err}")
    datetime_end_execution = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    execution_code += 1
    if 'logger' in locals() and hasattr(logger, 'error'):
        logger.error(f"[Executar] Execução Finalizada [{datetime_end_execution}] | new execution count [{execution_code}]")
    sys.exit(1)