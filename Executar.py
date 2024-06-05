from ConsorcioBB import ConsorcioBB
import pandas as pd

try:
    # 1.0 # INICIALIZANDO CLASSE #
    consorcio = ConsorcioBB(login="*", senha="*")
    siglas = ["TC","IM240","AI","AU","MO","EE", "IMP"]
    # siglas = ["AU"]
    # 1.0 #
    
    # 1.1 # Gera Grupos Ativos
    df_gruposAtivos = []
    for sigla in siglas:
        df = consorcio.generate_dataFrame_gruposAtivos(sigla=sigla)
        print('\n')
        if df is not None:  # Verificar se a função retornou um DataFrame válido
            df_gruposAtivos.append(df)
            print(f"Dataframe Grupos ({sigla}) exportados com sucesso.")
    df_todosGruposAtivos = pd.concat(df_gruposAtivos, ignore_index=True) # Concatenar todos os DataFrames em um único DataFrame
    df_todosGruposAtivos.to_excel("df_todosGruposAtivos.xlsx", index=False) # Salvar o DataFrame final grupos ativos em excel
    print("\nEXCEL 'Todos Grupos Ativos.xlsx' gerado com sucesso.")
    # 1.1 #

    # 1.1.2 # Gera dados da Assembleia em cima dos grupos Ativos
    lista_codigos = [codigo for codigo in df_todosGruposAtivos['Codigo'].tolist() if codigo != 0] # Lista para extrair dados de assembleia.
    df_assembleia = consorcio.generate_dataFrame_dadosAssembleia(lista_grupos=lista_codigos)
    df_assembleia.to_excel("df_assembleia.xlsx", index=False)
    print("\nEXCEL 'Dados Assembleia Grupos Ativos.xlsx' gerado com sucesso.")
    # 1.1.2 #

    # 1.2 # MESCLANDO INFORMAÇÕES PARA CRIAR EXCEL #
    # Mescle os dataframes usando o código do grupo como chave de junção
    df_excelFinal = pd.merge(df_todosGruposAtivos, df_assembleia, on="Codigo", how="left")
    colunas = ["Codigo", "TipoGrupo", "Prazo", "Vagas", "TaxaAdm", "FReserva", "CartasCredito", "VolumeTotal", "A 05/24", "QtdCont 05/24", "MEL 05/24","A 04/24", "QtdCont 04/24", "MEL 04/24", "A 03/24", "QtdCont 03/24","MEL 03/24"]
    df_excelFinal = df_excelFinal[colunas]
    df_excelFinal.to_excel("Grupos Ativos.xlsx", index=False) # Salvar o DataFrame final grupos ativos em excel
    print("\nEXCEL 'Grupos Ativos.xlsx' gerado com sucesso.")
    # 1.2 #

    # 1.3 # ENCERRA PROGRAMA.
    consorcio.encerrar_navegador()
    # 1.3 #
except Exception as err:
    print(f"\n[ERROR] Falha ao executar o script.\n{err}")


# FALTA ESTRUTURAR UMA FORMA DE IDENTIFICAR AS 3 ULTIMAS DATAS DE ASSEMBLEIA PARA NÃO PRECISAR
# INSERIR O NOME DAS COLUNAS NA MÃO E FICAR TROCANDO AS DATAS DAS ASSEMBLEIAS TODOS OS MESES.