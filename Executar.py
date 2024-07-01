# -- ANOTAÇÕES -- #
# 1.0 Testes Realizados com Chrome Driver v 125.0.6422.141 - https://googlechromelabs.github.io/chrome-for-testing/#stable
# 1.1 Testes Realizados com Selenium v 4.21.0
#
#   @Autor = Pedro Camara / GitHub: https://github.com/pedrokabh/ - Versão beta 1.0.3
#
# FALTA ESTRUTURAR UMA FORMA DE IDENTIFICAR AS 3 ULTIMAS DATAS DE ASSEMBLEIA PARA NÃO PRECISAR
# INSERIR O NOME DAS COLUNAS NA MÃO E FICAR TROCANDO AS DATAS DAS ASSEMBLEIAS TODOS OS MESES.

try:
    import sys, os
    from ConsorcioBB import ConsorcioBB
    import pandas as pd
except Exception as err:
    print(f"[Executar] FALHA AO CARREGAR IMPORTAÇÕES \n{err}")
    sys.exit(1)

try:
    # Parametros de Execucao #
    numero_execucao = 0
    # categoria_processadas = ["TC","AI","AU","MO","EE","IM240","IMP"]
    categoria_processadas = ["IMP"]
    endereco_chrome_driver = r"C:\Users\pedro\OneDrive\Área de Trabalho\AUTOMAÇÕES ISF\chromedriver-win64\chromedriver.exe"
    data_assembleia_mais_recente = "25/06/2024"
    data_assembleia_passada = "27/05/2024" 
    data_assembleia_retrasada = "25/04/2024"
    #

    # 1.0 # INICIALIZANDO CLASSE E PASSANDO PARÂMETROS 
    consorcio = ConsorcioBB(
                            login="PEDRO.CAMARA", 
                            senha="415143", 
                            endereco_chrome_driver = endereco_chrome_driver,
                            data_assembleia_mais_recente = data_assembleia_mais_recente,
                            data_assembleia_passada = data_assembleia_passada, 
                            data_assembleia_retrasada = data_assembleia_retrasada
                )
    
    # 1.1 # CRIAÇÃO ARQUIVO COM DADOS DOS GRUPOS ATIVOS.
    df_gruposAtivos = []
    for categoria in categoria_processadas:
        df = consorcio.ReturnsDataFrameWithActiveGroups(sigla=categoria)
        df_gruposAtivos.append(df)
    df_todosGruposAtivos = pd.concat(df_gruposAtivos, ignore_index=True)
    del df_gruposAtivos
    df_todosGruposAtivos.to_excel("df_todosGruposAtivos.xlsx", index=False)
    print("[Executar] Arquivo 'df_todosGruposAtivos.xlsx' criado com sucesso.")
    #

    # 1.2 # CRIAÇÃO ARQUIVO COM DADOS DA ASSEMBLEIA.
    lista_codigos = [codigo for codigo in df_todosGruposAtivos['Grupo'].tolist() if codigo != 0] # Ignora Grupos em Formação código é = 0
    df_assembleia = consorcio.ReturnsDataFrameAssemblyData(lista_grupos=lista_codigos)
    df_assembleia.to_excel("df_dados_assembleia.xlsx", index=False)
    print("[Executar] Arquivo 'df_dados_assembleia.xlsx' criado com sucesso.\n")
    #

    # 1.3 # MESCLANDO INFORMAÇÕES PARA CRIAR EXCEL #
    df_mediaContemplacao = consorcio.ReturnsDataFrameGroupsMedia(df_dadosAssembleia=df_assembleia, lista_grupos=lista_codigos)
    df_mediaContemplacao.to_excel("df_medias_assembleias.xlsx", index=False)
    print("[Executar] Arquivo 'df_medias_assembleias.xlsx' criado com sucesso.\n")
    #

    # 1.2.1 # MESCLANDO INFORMAÇÕES PARA CRIAR EXCEL #
    df_excelFinal = pd.merge(df_todosGruposAtivos, df_assembleia, on="Grupo", how="left")
    df_excelFinal = pd.merge(df_excelFinal, df_mediaContemplacao, on="Grupo", how="left")
    #
    
    # 1.2.3 # Adicionando media.
    df_excelFinal = df_excelFinal.drop(columns=["N° Assembleia M4", "N° Assembleia M5"])
    colunas = [
               "Modalidade", "Grupo", "Prazo", "Vagas", "Taxas", "Calculo TxAdm", "Calculo FR",
               "Calculo Total", "CartasCredito", "Liquidez", f"N° Assembleia {consorcio.sigla_assembleia_mais_recente}",
               f"Contemplados {consorcio.sigla_assembleia_mais_recente}", f"Media Lance {consorcio.sigla_assembleia_mais_recente}", f"Menor Lance {consorcio.sigla_assembleia_mais_recente}",
               f"Contemplados {consorcio.sigla_assembleia_passada}", f"Media Lance {consorcio.sigla_assembleia_passada}", f"Menor Lance {consorcio.sigla_assembleia_passada}", 
               f"Contemplados {consorcio.sigla_assembleia_retrasada}", f"Media Lance {consorcio.sigla_assembleia_retrasada}", f"Menor Lance {consorcio.sigla_assembleia_retrasada}" 
            ]
    df_excelFinal = df_excelFinal[colunas]
    df_excelFinal.to_excel("TodosGruposAtivos.xlsx", index=False)
    print("[Executar] Arquivo 'TodosGruposAtivos.xlsx' criado com sucesso.")

    # 1.3 # ENCERRA PROGRAMA.
    consorcio.EndBrowser()
except Exception as err:
    print(f"\n[Executar] Falha ao executar programa.\nERROR-> {err}")
    sys.exit(1)