# RPA PARA EXTRAÇÃO DE RELATÓRIO #
Automação acessa o portal venda externa da bb consórcio e extraí relatórios com grupos ativos na plataforma e dados de assembleia referentes aos grupos de consórcio.

## REQUIREMENTS ##
Para execução do projeto, é necessário estar com as seguintes bibliotecas Python instaladas.

```markdown
[Veja o arquivo de texto](./requirements.txt)

## VARIAVEIS DE EXECUÇÃO ##
    Antes de executar o script, primeiro verifique os parâmetros de execução necessários.
    
    * 1.0 - Para extrair a informação da média de contemplação da assembleia, precisa-se que os dois parâmetros `dados_assembleia, media_assembleia` sejam `True`.

    * 1.1 - Para extrair somente a informação de todos os grupos ativos, precisa-se que os dois parâmetros `dados_assembleia, media_assembleia` sejam `False`.

    * 1.2 - Para extrair a informação de todos os grupos ativos e dados da assembleia, sem a média de contemplação, precisa-se que o parâmetro `dados_assembleia` = `True` e `media_assembleia` = `False`.


```Python
    # 1.2 - VARIAVEIS DE EXECUÇÃO.
    dados_assembleia, media_assembleia = False, False
    
    # 1.3 - VARIAVEIS PARA EXECUTAR CLASSE CONSORCIO BB.
    categoria_processadas = ["TC", "AI", "AU", "MO", "EE", "IM240", "IMP"] # REFERENTE AOS TIPOS DE GRUPOS A SEREM EXTRAÍDOS.
    login = input("Digite o seu login: ") # LOGIN E SENHA DO PORTAL PARCEIROS
    senha = input("Digite a sua senha: ") 
    data_assembleia_passada, data_assembleia_retrasada, data_assembleia_mais_recente = None, None, None
    if dados_assembleia or media_assembleia: # PARAMETROS PARA IDENTIFICAR TRÊS ÚLTIMAS ASSEMBLEIAS (VERIFICAR NO PORTAL).
        data_assembleia_mais_recente = input("Digite a data da assembleia mais recente: ")    
        data_assembleia_passada = input("Digite a data da assembleia passada: ")
        data_assembleia_retrasada = input("Digite a data da assembleia retrasada: ")

