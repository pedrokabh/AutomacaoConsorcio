# -- ANOTAÇÕES -- #
# 1.0 Testes Realizados com Chrome Driver v 125.0.6422.141 - https://googlechromelabs.github.io/chrome-for-testing/#stable
# 1.1 Testes Realizados com Selenium v 4.21.0
#
#   @Autor = Pedro Camara / GitHub: https://github.com/pedrokabh/ - Versão beta 1.0.0
#
# FALTA ESTRUTURAR UMA FORMA DE IDENTIFICAR AS 3 ULTIMAS DATAS DE ASSEMBLEIA PARA NÃO PRECISAR
# INSERIR O NOME DAS COLUNAS NA MÃO E FICAR TROCANDO AS DATAS DAS ASSEMBLEIAS TODOS OS MESES.

try:
    import sys
    from ConsorcioBB import ConsorcioBB
    import pandas as pd
except Exception as err:
    print(f"[Executar] FALHA AO CARREGAR IMPORTAÇÕES \n{err}")
    sys.exit(1)

try:
    # 1.0 # INICIALIZANDO CLASSE E PASSANDO PARÂMETROS 
    consorcio = ConsorcioBB(
                    login="PEDRO.CAMARA", 
                    senha="415143", 
                    endereçoChromeDriver=r"C:\Users\pedro\OneDrive\Área de Trabalho\AUTOMAÇÕES ISF\chromedriver-win64\chromedriver.exe"
                )
    siglas = ["TC","IM240","AI","AU","MO","EE", "IMP"]
    # 1.0 #
    
    # 1.1 # CRIAÇÃO ARQUIVO COM DADOS DOS GRUPOS ATIVOS.
    df_gruposAtivos = []
    for sigla in siglas:
        df = consorcio.generate_dataFrame_gruposAtivos(sigla=sigla)
        df_gruposAtivos.append(df)
        print(f"[Executar] Data frame com Grupos Ativos '{sigla}' extraído. ")
    df_todosGruposAtivos = pd.concat(df_gruposAtivos, ignore_index=True)
    df_todosGruposAtivos.to_excel("df_todosGruposAtivos.xlsx", index=False)
    print("[Executar] Arquivo 'df_todosGruposAtivos.xlsx' criado com sucesso.")
    # 1.1 #

    # 1.1.2 # CRIAÇÃO ARQUIVO COM DADOS DA ASSEMBLEIA.
    lista_codigos = [codigo for codigo in df_todosGruposAtivos['Codigo'].tolist() if codigo != 0]
    df_assembleia = consorcio.generate_dataFrame_dadosAssembleia(lista_grupos=lista_codigos)
    df_assembleia.to_excel("df_assembleia.xlsx", index=False)
    print("[Executar] Arquivo 'df_assembleia.xlsx' criado com sucesso.")
    # 1.1.2 #

    # 1.2 # MESCLANDO INFORMAÇÕES PARA CRIAR EXCEL #
    df_excelFinal = pd.merge(df_todosGruposAtivos, df_assembleia, on="Codigo", how="left")
    colunas = ["TipoGrupo","Codigo", "Prazo", "Vagas", "TxAdm+FR", "CalcTxAdm", "CalcFR", "CalcTotal", "CartasCredito", "VolumeTotal", "A 05/24", "QtdCont 05/24", "MEL 05/24","A 04/24", "QtdCont 04/24", "MEL 04/24", "A 03/24", "QtdCont 03/24","MEL 03/24"]
    df_excelFinal = df_excelFinal[colunas]
    df_excelFinal = df_excelFinal.drop(columns=["A 04/24", "A 03/24"])
    df_excelFinal.to_excel("TodosGruposAtivos.xlsx", index=False)
    print("[Executar] Arquivo 'TodosGruposAtivos.xlsx' criado com sucesso.")
    # 1.2 #

    # 1.3 # ENCERRA PROGRAMA.
    consorcio.encerrar_navegador()
    # 1.3 #
except Exception as err:
    print(f"\n[Executar] Falha ao executar programa.\nERROR-> {err}")
    sys.exit(1)