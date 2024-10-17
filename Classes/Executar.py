import sys
import os
import pandas as pd
from datetime import datetime
import argparse
from ConsorcioBB import ConsorcioBB
from Logger import Logger

# 0.0 - FUNÇÃO ARGPARSER
def parse_args():
    parser = argparse.ArgumentParser(description="Executa o script ConsorcioBB com argumentos fornecidos pelo usuário.")

    parser.add_argument('--dados_assembleia', type=bool, default=True, help="Flag para ativar dados da assembleia.")
    parser.add_argument('--info_assembleia', type=bool, default=True, help="Flag para ativar informações da assembleia.")
    parser.add_argument('--categoria_processadas', nargs='+', default=["AU", "MO"], help="Lista de categorias processadas.")
    parser.add_argument('--data_assembleia_mais_recente', type=str, required=False, help="Data da assembleia mais recente.")
    parser.add_argument('--data_assembleia_passada', type=str, required=False, help="Data da assembleia passada.")
    parser.add_argument('--data_assembleia_retrasada', type=str, required=False, help="Data da assembleia retrasada.")
    parser.add_argument('--login', type=str, required=True, help="Login do usuário.")
    parser.add_argument('--senha', type=str, required=True, help="Senha do usuário.")
    
    return parser.parse_args()

# 0.1 - MAIN
try:
    # 1.0 - PARSING DOS ARGUMENTOS
    args = parse_args()

    # 2.0 - VARIAVEIS GLOBAIS
    executionCount_path = str(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')) + '\\Builds\\execution_count.txt')
    endereco_chrome_driver = str(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')) + '\\chromedriver-win64\\chromedriver.exe')
    global execution_code
    global currentExecution_directory

    # 3.0 - VARIÁVEIS DE EXECUÇÃO
    dados_assembleia = args.dados_assembleia
    info_assembleia = args.info_assembleia
    categoria_processadas = args.categoria_processadas
    data_assembleia_mais_recente = args.data_assembleia_mais_recente
    data_assembleia_passada = args.data_assembleia_passada
    data_assembleia_retrasada = args.data_assembleia_retrasada
    login = args.login
    senha = args.senha

    # 4.0 - LENDO LOG COUNT.TXT
    with open(executionCount_path, 'r') as arquivo:
        conteudo = arquivo.read()
        execution_code = int(conteudo)

    # 5.0 - CRIA O DIRETÓRIO PARA EXECUÇÃO DA BUILD ATUAL.
    currentExecution_directory = f'.\\Builds\\Build Number {execution_code}'
    os.makedirs(currentExecution_directory, exist_ok=True)
    open(os.path.join(currentExecution_directory, f'Build {execution_code}.log'), 'a').close()

    # 6.0 - INICIALIZANDO CLASSE LOGGER
    logger = Logger(execution_code, write_log_directory=f'{currentExecution_directory}\\Build {execution_code}.log')
    datetime_start_execution = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    logger.info(f"[Executar] Iniciando execução [{execution_code}] datetime [{datetime_start_execution}]")

    # 7.0 - EXECUÇÃO PARA EXTRAÇÃO DO RELATÓRIO - #
    try:
        # 7.1 - LOOP PARA PROCESSAMENTO SEPARADO DE CADA SEGMENTO DE GRUPOS.
        df_todos_segmentos = None
        for sigla in categoria_processadas:
            # 7.2 - INSTÂNCIANDO A CLASSE E REALIZANDO LOGIN NO SITE. 
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
            consorcio.Login()
            consorcio.ClosePoupUp()

            # 7.3 - EXTRAINDO OS GRUPOS ATIVOS DO SEGMENTO ATUAL.
            df_todosGruposAtivos = consorcio.ReturnsDataFrameWithActiveGroups(sigla=str(sigla))
            lista_id_grupos = [codigo for codigo in df_todosGruposAtivos['Grupo'].tolist() if codigo != 0]

            # 7.4 - EXTRAINDO OS DADOS DAS ULTIMAS 3 ASSEMBLEIAS.
            if dados_assembleia or info_assembleia:
                df_dados_assembleia = consorcio.ReturnsDataFrameAssemblyData(lista_grupos=lista_id_grupos, sigla=sigla)
                df_dados_assembleia = pd.merge(df_todosGruposAtivos, df_dados_assembleia, on="Grupo", how="left")

                # 7.5 - EXTRAINDO AS INFORMAÇÕES DE CADA ASSEMBLEIA (MEDIA/COTASCONTEMPLADAS/LL/LF/SORTEIO).
                if info_assembleia:
                    # 7.5.1 - REORDENANDO COLUNAS, CRIANDO ARQUIVO EXCEL, ADICIONANDO AO LOG E ENCERRANDO NAVEGADOR.
                    df_mediaContemplacao = consorcio.ReturnsDataFrameGroupsInfo(df_dadosAssembleia=df_dados_assembleia, lista_grupos=lista_id_grupos, sigla=sigla)
                    df_mediaContemplacao = pd.merge(df_dados_assembleia, df_mediaContemplacao, on="Grupo", how="left")
                    colunas = [
                        "Modalidade", "Grupo", "Prazo", "Vagas", "Taxas", "Calculo TxAdm", "Calculo FR",
                        "Calculo Total", "CartasCredito", "Liquidez", f"N° Assembleia {consorcio.sigla_assembleia_mais_recente}",
                        f"Contemplados {consorcio.sigla_assembleia_mais_recente}",  f"Qtde Cotas LL {consorcio.sigla_assembleia_mais_recente}", f"Qtde Cotas LF {consorcio.sigla_assembleia_mais_recente}",    f"Qtde Cotas Sorteio {consorcio.sigla_assembleia_mais_recente}", f"Media Lance {consorcio.sigla_assembleia_mais_recente}",   f"Menor Lance {consorcio.sigla_assembleia_mais_recente}",
                        f"Contemplados {consorcio.sigla_assembleia_passada}",       f"Qtde Cotas LL {consorcio.sigla_assembleia_passada}",      f"Qtde Cotas LF {consorcio.sigla_assembleia_passada}"     ,    f"Qtde Cotas Sorteio {consorcio.sigla_assembleia_passada}"     , f"Media Lance {consorcio.sigla_assembleia_passada}"     ,   f"Menor Lance {consorcio.sigla_assembleia_passada}",
                        f"Contemplados {consorcio.sigla_assembleia_retrasada}",     f"Qtde Cotas LL {consorcio.sigla_assembleia_retrasada}",    f"Qtde Cotas LF {consorcio.sigla_assembleia_retrasada}"   ,    f"Qtde Cotas Sorteio {consorcio.sigla_assembleia_retrasada}"   , f"Media Lance {consorcio.sigla_assembleia_retrasada}"   ,   f"Menor Lance {consorcio.sigla_assembleia_retrasada}"
                    ]
                    df_mediaContemplacao = df_mediaContemplacao[colunas]
                    df_mediaContemplacao.to_excel(f'{currentExecution_directory}\\InfoAssembleias [{sigla}] BUILD({execution_code}).xlsx', index=False)
                    logger.info(f"[Executar] Arquivo [InfoAssembleias [{sigla}] BUILD({execution_code}).xlsx] criado com sucesso.")
                    consorcio.EndBrowser()
                    # 7.5.2 - ADICIONANDO DADOS AO DF_TODOS_SEGMENTOS. 
                    df_todos_segmentos = df_mediaContemplacao if df_todos_segmentos == None else pd.concat([df_todos_segmentos, df_mediaContemplacao], ignore_index=True)
                
                # 7.5 - SALVANDO SOMENTE O DF COM GRUPOS ATIVOS E DADOS DAS ASSEMBLEIAS.
                else:
                    # 7.5.1 - CRIANDO ARQUIVO EXCEL, ADICIONANDO AO LOG E ENCERRANDO NAVEGADOR.
                    df_dados_assembleia.to_excel(f'{currentExecution_directory}\\DadosAssembleia [{sigla}] BUILD({execution_code}).xlsx', index=False)
                    logger.info(f"[Executar] Arquivo [DadosAssembleia [{sigla}] BUILD({execution_code}).xlsx] criado com sucesso.")
                    consorcio.EndBrowser()
                    # 7.5.2 - ADICIONANDO DADOS AO DF_TODOS_SEGMENTOS. 
                    df_todos_segmentos = df_dados_assembleia if df_todos_segmentos == None else pd.concat([df_todos_segmentos, df_dados_assembleia], ignore_index=True)
            
            # 7.4 - SALVANDO SOMENTE O DF COM DADOS DOS GRUPOS ATIVOS.
            else:
                # 7.4.1 - CRIANDO ARQUIVO EXCEL, ADICIONANDO AO LOG E ENCERRANDO NAVEGADOR.
                df_todosGruposAtivos.to_excel(f'{currentExecution_directory}\\DadosGruposAtivos [{sigla}] BUILD({execution_code}).xlsx', index=False)
                logger.info(f"[Executar] Arquivo [DadosGruposAtivos [{sigla}] BUILD({execution_code}).xlsx] criado com sucesso.")
                consorcio.EndBrowser()
                # 7.4.2 - ADICIONANDO DADOS AO df_todos_segmentos. 
                df_todos_segmentos = df_todosGruposAtivos if df_todos_segmentos == None else pd.concat([df_todos_segmentos, df_todosGruposAtivos], ignore_index=True)
            
        # 7.2 - SALVANDO ARQUIVO COM TODOS OS SEGMENTOS AGRUPADOS.
        df_todos_segmentos.to_excel(f'{currentExecution_directory}\\TodosSegmentos [{categoria_processadas}] - BUILD({execution_code}).xlsx', index=False)
        logger.info(f"[Executar] Arquivo [TodosSegmentos [{categoria_processadas}] - BUILD({execution_code}).xlsx] criado com sucesso.")

    except Exception as err:
        logger.error(f"[Executar] FALHA DE EXECUÇÃO PARA EXTRAÇÃO DO RELATÓRIO ->\n{err}")
        consorcio.EndBrowser()
        raise
    
    # 8.0 - ATUALIZANDO LOG COUNT.TXT
    logger.info(f"[Executar] Execução Finalizada [{datetime.now().strftime('%H:%M:%S')}] || Finalized Execution [{execution_code}]")
    with open(r'.\\Builds\\execution_count.txt', 'w') as arquivo:
        arquivo.write(str(execution_code + 1))

except Exception as err:
    # 8.1 - TRATANDO EXEÇÕES.
    if 'logger' in locals() and hasattr(logger, 'error'):
        logger.error(f"[Executar] Falha ao executar programa.\nERROR-> {err}")
    sys.exit(1)
