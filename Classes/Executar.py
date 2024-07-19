import sys
import os
import pandas as pd
from datetime import datetime
from ConsorcioBB import ConsorcioBB
from Logger import Logger

try:
    os.system('cls')
    # Diretório base do projeto
    directory_mother = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))

    with open(str(directory_mother+'\\Classes\\execution_count.txt'), 'r') as arquivo:
        conteudo = arquivo.read()
    execution_code = int(conteudo)

    # Caminho para o arquivo de log
    directory_mother = f'{directory_mother}\\Builds\\Build Number {execution_code}\\'
    os.makedirs(os.path.dirname( str(directory_mother) ), exist_ok=True)

    try:
        # Inicializando o Logger
        logger = Logger(execution_code, write_log_directory=f'.\\Builds\\Build Number {execution_code}\\Build {execution_code}.log')
        datetime_start_execution = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        logger.info(f"[Executar] Iniciando execução [{execution_code}] datetime [{datetime_start_execution}]")

        # Parâmetros de execução
        categoria_processadas = ["TC", "AI", "AU", "MO", "EE", "IM240", "IMP"]
        # categoria_processadas = ["TC"]
        endereco_chrome_driver = r"C:\Users\pedro\OneDrive\Área de Trabalho\AUTOMAÇÕES ISF\chromedriver-win64\chromedriver.exe"
        data_assembleia_mais_recente = "25/06/2024"
        data_assembleia_passada = "27/05/2024"
        data_assembleia_retrasada = "25/04/2024"
        login = input("Digite o seu login: ")
        senha = input("Digite a sua senha: ")

        # Inicializando a classe ConsorcioBB
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

        # Criando arquivo com dados dos grupos ativos
        df_gruposAtivos = [consorcio.ReturnsDataFrameWithActiveGroups(sigla=categoria) for categoria in categoria_processadas]
        df_todosGruposAtivos = pd.concat(df_gruposAtivos, ignore_index=True)
        df_todosGruposAtivos.to_excel(f'.\\Builds\\Build Number {execution_code}\\df_todosGruposAtivos.xlsx', index=False)
        logger.info("[Executar] Arquivo 'df_todosGruposAtivos.xlsx' criado com sucesso.")

        # Criando arquivo com dados da assembleia
        lista_codigos = [codigo for codigo in df_todosGruposAtivos['Grupo'].tolist() if codigo != 0]
        df_assembleia = consorcio.ReturnsDataFrameAssemblyData(lista_grupos=lista_codigos)
        df_assembleia.to_excel(f'.\\Builds\\Build Number {execution_code}\\df_dados_assembleia.xlsx', index=False)
        logger.info("[Executar] Arquivo 'df_dados_assembleia.xlsx' criado com sucesso.")

        # Mesclando informações para criar Excel
        # df_mediaContemplacao = consorcio.ReturnsDataFrameGroupsMedia(df_dadosAssembleia=df_assembleia, lista_grupos=lista_codigos)
        # df_mediaContemplacao.to_excel(f'.\\Builds\\Build Number {execution_code}\\df_medias_assembleias.xlsx', index=False)
        # logger.info("[Executar] Arquivo 'df_medias_assembleias.xlsx' criado com sucesso.")

        df_excelFinal = pd.merge(df_todosGruposAtivos, df_assembleia, on="Grupo", how="left")
        # df_excelFinal = pd.merge(df_excelFinal, df_mediaContemplacao, on="Grupo", how="left")

        # Adicionando colunas específicas
        # COLUNA SEM MEDIA
        colunas = [
            "Modalidade", "Grupo", "Prazo", "Vagas", "Taxas", "Calculo TxAdm", "Calculo FR",
            "Calculo Total", "CartasCredito", "Liquidez", f"N° Assembleia {consorcio.sigla_assembleia_mais_recente}",
            f"Contemplados {consorcio.sigla_assembleia_mais_recente}", f"Menor Lance {consorcio.sigla_assembleia_mais_recente}",
            f"Contemplados {consorcio.sigla_assembleia_passada}",  f"Menor Lance {consorcio.sigla_assembleia_passada}",
            f"Contemplados {consorcio.sigla_assembleia_retrasada}", f"Menor Lance {consorcio.sigla_assembleia_retrasada}"
        ]

        #COLUNA COM MEDIA
        # colunas = [
        #     "Modalidade", "Grupo", "Prazo", "Vagas", "Taxas", "Calculo TxAdm", "Calculo FR",
        #     "Calculo Total", "CartasCredito", "Liquidez", f"N° Assembleia {consorcio.sigla_assembleia_mais_recente}",
        #     f"Contemplados {consorcio.sigla_assembleia_mais_recente}", f"Media Lance {consorcio.sigla_assembleia_mais_recente}", f"Menor Lance {consorcio.sigla_assembleia_mais_recente}",
        #     f"Contemplados {consorcio.sigla_assembleia_passada}", f"Media Lance {consorcio.sigla_assembleia_passada}", f"Menor Lance {consorcio.sigla_assembleia_passada}",
        #     f"Contemplados {consorcio.sigla_assembleia_retrasada}", f"Media Lance {consorcio.sigla_assembleia_retrasada}", f"Menor Lance {consorcio.sigla_assembleia_retrasada}"
        # ]
        df_excelFinal = df_excelFinal[colunas]
        df_excelFinal.to_excel(f'.\\Builds\\Build Number {execution_code}\\TodosGruposAtivos.xlsx', index=False)
        logger.info("[Executar] Arquivo 'TodosGruposAtivos.xlsx' criado com sucesso.")

        # Encerrando o navegador
        consorcio.EndBrowser()

        # Atualizando o contador de execução
        with open(str(directory_mother+'\\Classes\\execution_count.txt'), 'w') as arquivo:
            arquivo.write(str(execution_code + 1))

        datetime_end_execution = datetime.now().strftime("%H:%M:%S")
        logger.info(f"[Executar] Execução Finalizada [{datetime_end_execution}] | new execution count [{execution_code + 1}]")

    except Exception as err:
        if 'logger' in locals() and hasattr(logger, 'error'):
            logger.error(f"\n[Executar] Falha ao executar programa.\nERROR-> {err}")
        datetime_end_execution = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        execution_code += 1
        if 'logger' in locals() and hasattr(logger, 'error'):
            logger.error(f"[Executar] Execução Finalizada [{datetime_end_execution}] | new execution count [{execution_code}]")
        sys.exit(1)

except Exception as err:
    print(f"[Executar] FALHA AO CARREGAR IMPORTAÇÕES \n{err}")
    sys.exit(1)
