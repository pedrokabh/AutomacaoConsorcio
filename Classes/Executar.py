import sys
import os
import pandas as pd
from datetime import datetime
import argparse
from ConsorcioBB import ConsorcioBB
from Logger import Logger

# 0.0 - FUNÇÕES DA CLASSE.
def parse_args():
    parser = argparse.ArgumentParser(description="Executa o script ConsorcioBB com argumentos fornecidos pelo usuário.")

    # parser.add_argument('--dados_assembleia', type=bool, default=True, help="Flag para ativar dados da assembleia.")
    # parser.add_argument('--info_assembleia', type=bool, default=True, help="Flag para ativar informações da assembleia.")
    # parser.add_argument('--categoria_processadas', nargs='+', default=["AU", "MO"], help="Lista de categorias processadas.")
    parser.add_argument('--data_assembleia_mais_recente', type=str, required=False, help="Data da assembleia mais recente.")
    parser.add_argument('--data_assembleia_passada', type=str, required=False, help="Data da assembleia passada.")
    parser.add_argument('--data_assembleia_retrasada', type=str, required=False, help="Data da assembleia retrasada.")
    parser.add_argument('--login', type=str, required=True, help="Login do usuário.")
    parser.add_argument('--senha', type=str, required=True, help="Senha do usuário.")
    
    return parser.parse_args()

def salvar_dataframe_excel(df, caminho, nome_arquivo, logger):
    try:
        df.to_excel(caminho, index=False)
        logger.info(f"[Executar] Arquivo [{nome_arquivo}.xlsx] criado com sucesso.")
    except Exception as err:
        logger.error(f"[Executar] Falha ao salvar arquivo [{nome_arquivo}.xlsx].\nERROR: {err}")

# 0.1 - MAIN
try:
    # 1.0 - PARSING DOS ARGUMENTOS
    args = parse_args()

    # 2.0 - VARIAVEIS GLOBAIS
    execution_code = None
    base_path = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    path = {    
        'currentExecution_directory': None,
        'executionCount_path': os.path.join(base_path, 'Builds', 'execution_count.txt'),
        'endereco_chrome_driver': os.path.join(base_path, 'chromedriver-win64', 'chromedriver.exe'),
        'cwd': None
    }

    # 3.0 - VARIÁVEIS DE EXECUÇÃO
    dados_assembleia = True
    info_assembleia = True
    categoria_processadas = ["IM240"]

    # 4.0 - LENDO LOG COUNT.TXT E INCLUINDO CAMINHOS.
    with open(path['executionCount_path'], 'r') as arquivo:
        conteudo = arquivo.read()
        execution_code = int(conteudo)
        path['currentExecution_directory'] = os.path.join(base_path, 'Builds', f'Build Number {execution_code}')
        path['cwd'] = os.path.join(base_path, 'Builds', f'Build Number {execution_code}', f'Build {execution_code}.log')

    # 5.0 - CRIA O DIRETÓRIO PARA EXECUÇÃO DA BUILD ATUAL.
    os.makedirs(path['currentExecution_directory'], exist_ok=True)
    open(os.path.join(path['currentExecution_directory'], f'Build {execution_code}.log'), 'a').close()

    # 6.0 - INICIALIZANDO CLASSE LOGGER
    logger = Logger(execution_code, write_log_directory=f'{path['currentExecution_directory']}\\Build {execution_code}.log')
    datetime_start_execution = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    logger.info(f"[Executar] Iniciando execução [{execution_code}] datetime [{datetime_start_execution}]")

    # 7.0 - EXECUÇÃO PARA EXTRAÇÃO DO RELATÓRIO - #
    df_todos_segmentos = None
    try:
        # 7.1 - LOOP PARA PROCESSAMENTO SEPARADO DE CADA SEGMENTO DE GRUPOS.
        for sigla in categoria_processadas:
            # 7.1.2 - INSTÂNCIANDO A CLASSE E REALIZANDO LOGIN NO SITE. 
            consorcio = ConsorcioBB(
                login=args.login,
                senha=args.senha,
                endereco_chrome_driver=path['endereco_chrome_driver'],
                logger=logger,
                cwd=path['cwd'],
                data_assembleia_mais_recente=args.data_assembleia_mais_recente,
                data_assembleia_passada=args.data_assembleia_passada,
                data_assembleia_retrasada=args.data_assembleia_retrasada
            )
            consorcio.Login()
            consorcio.ClosePoupUp()

            # 7.1.3 - EXTRAINDO GRUPOS ATIVOS DO SEGMENTO ATUAL.
            df_todosGruposAtivos = consorcio.ReturnsDataFrameWithActiveGroups(sigla=str(sigla))
            lista_id_grupos = [codigo for codigo in df_todosGruposAtivos['Grupo'].tolist() if codigo != 0]

            # 7.1.4 - SALVANDO ARQUIVO COM SOMENTE GRUPOS ATIVOS.
            if not dados_assembleia and not info_assembleia:
                # 7.4.1 - CRIANDO ARQUIVO EXCEL, E ENCERRANDO NAVEGADOR.
                salvar_dataframe_excel(df=df_todosGruposAtivos, caminho=f'{path['currentExecution_directory']}\\DadosGruposAtivos [{sigla}] BUILD ({execution_code}).xlsx', nome_arquivo=f'DadosGruposAtivos [{sigla}] BUILD ({execution_code})', logger=logger)
                consorcio.EndBrowser()
                # 7.4.2 - TENTA ADICIONAR DADOS AO 'df_todos_segmentos'.
                try: df_todos_segmentos = df_todosGruposAtivos if df_todos_segmentos is None else pd.concat([df_todos_segmentos, df_todosGruposAtivos], ignore_index=True)
                except Exception as err: logger.error(f"[Executar] Falha ao concatenar dados [{sigla}] ao  DF_TODOS_SEGMENTOS.")
            # 7.1.5 - SALVANDO ARQUIVO COM SOMENTE GRUPOS ATIVOS + INFORMAÇÕES DE ASSEMBLEIAS.
            else:
                # 7.4.1 - Extraindo Dados Assembleia e já Concatenando o df_todosGruposAtivos.
                df_dados_assembleia = pd.merge(df_todosGruposAtivos, consorcio.ReturnsDataFrameAssemblyData(lista_grupos=lista_id_grupos, sigla=sigla), on="Grupo", how="left")
                # 7.5 - INFO  ASSEMBLIEAS (LIQUIDEZ/ COTAS CONTEMPLADAS/ MENOR LANCE).
                if info_assembleia:
                    # 7.5.1 - Extraindo Informações das Assembleias, e já Concatenando o df_dados_assembleia.
                    df_mediaContemplacao = pd.merge(df_dados_assembleia,consorcio.ReturnsDataFrameGroupsInfo(df_dadosAssembleia=df_dados_assembleia, lista_grupos=lista_id_grupos, sigla=sigla),on="Grupo",how="left")
                    # 7.5.2 - SALVANDO ARQUIVO, ORGANIZANDO A POSIÇÕES DAS COLUNAS, E ENCERRANDO NAVEGADOR.
                    colunas = [
                        "Modalidade", "Grupo", "Prazo", "Vagas", "Taxas", "Calculo TxAdm", "Calculo FR",
                        "Calculo Total", "CartasCredito", "Liquidez", f"N° Assembleia {consorcio.sigla_assembleia_mais_recente}",
                        f"Contemplados {consorcio.sigla_assembleia_mais_recente}",  f"Qtde Cotas LL {consorcio.sigla_assembleia_mais_recente}", f"Qtde Cotas LF {consorcio.sigla_assembleia_mais_recente}",    f"Qtde Cotas Sorteio {consorcio.sigla_assembleia_mais_recente}", f"Media Lance {consorcio.sigla_assembleia_mais_recente}",   f"Menor Lance {consorcio.sigla_assembleia_mais_recente}",
                        f"Contemplados {consorcio.sigla_assembleia_passada}",       f"Qtde Cotas LL {consorcio.sigla_assembleia_passada}",      f"Qtde Cotas LF {consorcio.sigla_assembleia_passada}"     ,    f"Qtde Cotas Sorteio {consorcio.sigla_assembleia_passada}"     , f"Media Lance {consorcio.sigla_assembleia_passada}"     ,   f"Menor Lance {consorcio.sigla_assembleia_passada}",
                        f"Contemplados {consorcio.sigla_assembleia_retrasada}",     f"Qtde Cotas LL {consorcio.sigla_assembleia_retrasada}",    f"Qtde Cotas LF {consorcio.sigla_assembleia_retrasada}"   ,    f"Qtde Cotas Sorteio {consorcio.sigla_assembleia_retrasada}"   , f"Media Lance {consorcio.sigla_assembleia_retrasada}"   ,   f"Menor Lance {consorcio.sigla_assembleia_retrasada}"
                    ]
                    salvar_dataframe_excel(df=df_mediaContemplacao[colunas], caminho=f'{path['currentExecution_directory']}\\InfoAssembleias [{sigla}] BUILD ({execution_code}).xlsx', nome_arquivo=f"InfoAssembleias [{sigla}] BUILD ({execution_code})", logger=logger)
                    consorcio.EndBrowser()
                    # 7.5.3 - ADICIONA DADOS AO DF_TODOS_SEGMENTOS.
                    try: df_todos_segmentos = df_mediaContemplacao[colunas] if df_todos_segmentos is None else pd.concat([df_todos_segmentos, df_mediaContemplacao[colunas]], ignore_index=True)
                    except Exception as err: logger.error(f"[Executar] Falha ao concatenar dados [{sigla}] ao  DF_TODOS_SEGMENTOS.")
                # 7.5 - DADOS ASSEMBLIEAS (MÉDIA CONTEMPLAÇÃO/QTDE LF/QTDE LL/QTDE SORTEIOS).
                else:
                    # 7.5.1 - CRIANDO ARQUIVO EXCEL, ADICIONANDO AO LOG E ENCERRANDO NAVEGADOR.
                    salvar_dataframe_excel(df=df_dados_assembleia, caminho=f'{path['currentExecution_directory']}\\DadosAssembleia [{sigla}] BUILD ({execution_code}).xlsx', nome_arquivo=f'DadosAssembleia [{sigla}] BUILD ({execution_code})', logger=logger)
                    consorcio.EndBrowser()
                    # 7.5.2 - TENTA ADICIONAR DADOS AO 'DF_TODOS_SEGMENTOS'.
                    try: df_todos_segmentos = df_dados_assembleia if df_todos_segmentos is None else pd.concat([df_todos_segmentos, df_dados_assembleia], ignore_index=True)
                    except Exception as err: logger.error(f"[Executar] Falha ao concatenar dados [{sigla}] ao  DF_TODOS_SEGMENTOS.")
                
        # 7.2 - SALVANDO ARQUIVO COM TODOS OS SEGMENTOS AGRUPADOS.
        salvar_dataframe_excel(df=df_todos_segmentos, caminho=f'{path['currentExecution_directory']}\\TodosSegmentos {categoria_processadas} BUILD ({execution_code}).xlsx', nome_arquivo=f'TodosSegmentos {categoria_processadas} BUILD ({execution_code})', logger=logger)
        
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
