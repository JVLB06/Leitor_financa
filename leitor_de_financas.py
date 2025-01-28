# pyinstaller --noconsole --onefile --icon="logo.ico" --add-data="arq.xlsx;." --add-data="logo.png;." leitor_de_financas.py

import PySimpleGUI as gui
import pandas as pd
from scrapping import get_data
from scrapping import manage_data
from scrapping import abre_plan
# Declaração de variáveis essenciais
ciclo = True
arquivo = "arq.xlsx"
nome = "Leitor de finanças"
selic = 12.5
logo = "logo.png"
gui.theme("PythonPlus")
# Tela de saída
def sure():
    fechar = True
    while True:
        exit = [[gui.Text("Tem certeza que deseja sair?", font=("Arial", 15))],
                [gui.Text("Qualquer alteração não salva será perdida", font=("Arial", 15))],
                [gui.Button("Sair", key="sair", font=("Arial", 13)),
                 gui.Button("Cancelar", key="cancel", font=("Arial", 13))]]
        fecho = gui.Window("Tem certeza?", exit, element_justification='c', size=(450, 150))
        event, values = fecho.read()
        if event == "sair":
            fechar = False
            break
        if event == "cancel" or event == gui.WIN_CLOSED:
            break
    return fechar
def como_usar(): 
    usar = "A planilha excel, serve como referência para o programa localizar as informações desejadas.\n\nDessa forma, a única parte da planilha que pode ser editada (para o bom funcionamento do programa), é a pasta de nome 'referências', sendo mantido o cabeçalho (ticker & tipo), e nas linhas a seguir deve conter, o nome do ticker (em sua bolsa de origem, por exemplo AMZO34 = AMZN ticker na nasdaq), e o tipo de ativo, FII para fundos, ação e bdr para ações e ativos no exterior.\nA partir disso, o programa irá gerar quatro (4) novas pastas, sendo duas delas para armazenar os dados gerais (Ação & FII) e duas delas para os dados processados e classificados, para a sua respctiva análise (make_acao & make_fii).\n\nÉ importante rodar o programa de extração de dados com a planilha fechada, afim de evitar erros\n\nTendo o programa concluído todos os processos, você terá acesso a essa planilha, completa e devidamente classificada para uso. Tendo seus cálculos, base em uma análise fundamentalista conservadora de todas as ações, tendo como referência os seus indicadores fundamentalistas.\nCom isso, o programa avalia com notas de 0 a 10 e define uma média das notas dos quatro âmbitos mais relevantes financeiramente, na empresa. A partir dessas informações, classifica e entrega devidamente ordenadas todos os ativos em ordem decrescente."
    while True:
        como = [[gui.Text(f"Como usar a planilha de {nome}?")],
                [gui.Multiline(usar, key="sobre", size=(50, 15), disabled=True, no_scrollbar=False)],
                [gui.VSeparator(pad=(0,10))],
                [gui.Button("Entendido", key="got it", size=(250, 10))]]
        window = gui.Window("Como usar?", como, element_justification="c", size=(300, 350))
        event, values = window.read()
        if event == "got it":
            window.close()
            break
# Ciclo principal
while ciclo == True:
# Tela inicial de boas vindas
    # Tela principal 
    welcome = [[gui.Text(f"Bem vindo ao {nome}!", font=("Arial", 15))],
               [gui.Image(filename=logo)],
            [gui.Button("Sair", key="exit", font=("Arial", 13)),
            gui.Button("Ajuda", key="help", font=("Arial", 13)),
            gui.Button("Iniciar coleta", key="start", font=("Arial", 13)),
            gui.Button("Atualizar informações", key="update", font=("Arial", 13))],
            [gui.Button("Add. Ativos para leitura", key="add", font=("Arial", 13))]] 
    principal = gui.Window("Bem vindo", welcome, element_justification="c", size=(500,430))
    event, values = principal.read()
    if event == "add":
        principal.close()
        cia = []
        head = ["Ticker", "Tipo"]
        dados = pd.read_excel(arquivo, sheet_name=None)
        data = dados["Referências"]
        for _, row in data.iterrows():
            dt = [] 
            dt.append(str(row["ticker"]).upper())
            dt.append(str(row["tipo"]).upper())
            cia.append(dt.copy())
        while True:
            adiciona_info = [[gui.Text("Lista de tickers cadastrados:", font=("Arial", 15))],
                             [gui.Table(values=cia, headings=head, max_col_width=25, auto_size_columns=True, justification="center", alternating_row_color="lightblue", enable_events=True)],
                             [gui.Text("Insira uma ação:", font=("Arial", 13)), gui.InputText(key="ticker", size=(20,100))],
                             [gui.Text("Insira o tipo de ação:", font=("Arial", 13))],
                             [gui.Radio("FII", "1", key="fii"), gui.Radio("AÇÃO", "1", key="acao", default=True), gui.Radio("BDR", "1", key="bdr")],
                             [gui.Button("Confirma", key="conf", font=("Arial", 13)), gui.Button("Cancelar", key="canc", font=("Arial", 13)), gui.Button("Excluir", key="del", font=("Arial", 13))]]
            screen = gui.Window("Adicione informações", adiciona_info, size=(350,400), element_justification="c")
            event, values = screen.read()
            if event == "conf":
                # Adiciona o novo dado sem duplicação
                novo_ticker = str(values["ticker"]).upper()
                if values["fii"]:
                    novo_tipo = "FII"
                if values["acao"]:
                    novo_tipo = "AÇÃO"
                if values["bdr"]:
                    novo_tipo = "BDR"
                if [novo_ticker, novo_tipo] not in cia:  # Evita duplicados
                    cia.append([novo_ticker, novo_tipo])
                screen.close()
            if event == "canc" or event == gui.WIN_CLOSED:
                refer = pd.DataFrame(cia, columns=["ticker", "tipo"])
                with pd.ExcelWriter(arquivo, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                    refer.to_excel(writer, sheet_name="Referências", index=False)
                screen.close()
                break      
            if event == "del":
                screen.close()
                while True:
                    exclui_info = [[gui.Text("Lista de tickers cadastrados:", font=("Arial", 15))],
                             [gui.Table(values=cia, headings=head, max_col_width=25, auto_size_columns=True, justification="center", alternating_row_color="lightblue", enable_events=True)],
                             [gui.Text("Insira o item a ser excluído:", font=("Arial", 13)), gui.InputText(key="ticker", size=(20,100))],
                             [gui.Button("Confirma", key="conf", font=("Arial", 13)), gui.Button("Cancelar", key="canc", font=("Arial", 13))]]
                    visor = gui.Window("Remova informações", exclui_info, size=(350,400), element_justification="c")
                    event, values = visor.read()
                    if event == "canc" or event == gui.WIN_CLOSED:
                        refer = pd.DataFrame(cia, columns=["ticker", "tipo"])
                        with pd.ExcelWriter(arquivo, mode="a", engine="openpyxl", if_sheet_exists="replace") as writer:
                            refer.to_excel(writer, sheet_name="Referências", index=False)
                        visor.close()
                        break   
                    if event == "conf":
                        visor.close()
                        for a in range(len(cia)):
                            if str(values["ticker"]).upper() == cia[a][0]:
                                cia.pop(a)
                                break
    if event == "exit" or event == gui.WIN_CLOSED:
        # Sair
        principal.close()
        # Tem certeza que deseja sair?
        ciclo = sure()      
    if event == "update":
        # Atualizar Selic
        principal.close()
        while True:
            update = [[gui.Text("Deseja atualizar dados?", font=("Arial", 15))],
                      [gui.Text("Atualizar taxa Selic de referência:", font=("Arial", 15)),
                      gui.InputText(key="selic", size=(15,20))],
                      [gui.Text("Atualizar informações da planilha:", font=("Arial", 15)),
                       gui.Button("Planilha", key="excel", font=("Arial", 13))],
                      [gui.Button("Confirma", key="confirma", font=("Arial", 13)),
                       gui.Button("Cancelar", key="c", font=("Arial", 13))]]
            atualiza = gui.Window("Atualizar", update, element_justification='c', size=(500,160))
            event, values = atualiza.read()
            # Atualizar planilha
            if event == "excel":
                atualiza.close()
                como_usar()
                # Abre planilha pra edição
                abre_plan(arquivo)
                break
            # Atualizar Selic
            if event == "confirma":
                selic = str(values["selic"])
                if selic.find(","):
                   selic.replace(",", ".")
                selic = float(selic) 
                atualiza.close()
                # A concluir
            if event == "c" or event == gui.WIN_CLOSED:
                atualiza.close()
                break
    if event == "start":
        principal.close()
        while True:
            # Iniciar a coleta de dados
            start = [[gui.Text("Iniciar coleta de dados?", font=("Arial", 15))],
                     [gui.Button("Iniciar", key="init", font=("Arial", 13)),
                      gui.Button("Cancelar", key="cancelar", font=("Arial", 13))]]
            comeca = gui.Window("Capturar dados", start, element_justification='c', size=(300, 100))
            event, values = comeca.read()
            if event == "init":
                def verifica_uso():
                    try:
                        # Tenta abrir o arquivo em modo leitura/escrita exclusivo
                        with open(arquivo, "r+"):
                            pass
                        return False  # Arquivo não está em uso
                    except IOError:
                        return True  # Arquivo está em uso
                testa = verifica_uso()
                if testa == False:
                    # Botão de start
                    comeca.close()
                    # Começa o scrapping
                    get_data(arquivo)
                    # Organiza a planilha
                    manage_data(selic, arquivo)
                    # Orientação sobre planilha
                    como_usar()
                    # Abre a planilha pra visualização
                    abre_plan(arquivo)
                else:
                    while True:
                        layout = [[gui.Text("A planilha está aberta no momento.", font=("Arial", 15))],
                        [gui.Text("Feche-a e tente realizar o processo novamente", font=("Arial", 15))],
                        [gui.Button("Ok", key="ok", font=("Arial", 13))]]   
                        tela = gui.Window("Erro", layout, size=(450,150), element_justification="c")             
                        event, values = tela.read()
                        if event == "ok" or event == "cancel" or event == gui.WIN_CLOSED:
                            tela.close()
                            break
            if event == "cancelar" or event == gui.WIN_CLOSED:
                # Sair
                comeca.close()
                break
    if event == "help":
        principal.close()
        while True:
            # Ajuda
            sobre_txt = f"O {nome}, vem com o intuito de te ajudar com suas análises em renda variável!\nAdotando sempre uma postura mais conservadora, prezamos por analisar de forma automatizada os indicadores de ações e fundos imobiliários, afim de poupar o trabalho de horas de pesquisa. Retornando para você, uma planilha com as análises fundamentalistas, prontas para uso. \nComo funciona? \n\n-Acessando a tela inicial você pode alterar alguns parâmetros ou realizar a busca de dados direto! \n-Para alterar os parâmetros ou visualizar a planilha, basta clicar em ATUALIZAR INFORMAÇÕES e alterar conforme a nossa planilha padrão. \n-Para realizar a busca de dados, basta clicar em INICIAR COLETA e rodar o código para que o programa faça a sua parte e te entregue a planilha pronta para uso. \n.\n.\n.\n-OBS: O {nome}, não se responsabiliza pelas operações realizadas por quaisquer usuários."
            help = [[gui.Text(f"Como funciona o {nome}?")],
                    [gui.Multiline(sobre_txt, key="sobre", size=(50, 15), disabled=True, no_scrollbar=False)],
                    [gui.VSeparator(pad=(0,10))],
                    [gui.Button("Sair", key="sair", size=(250, 10))]]
            ajuda = gui.Window("Ajuda", help, element_justification='c', size=(300, 350))
            eventos, valores = ajuda.read()
            if eventos == "sair" or eventos == gui.WIN_CLOSED:
                ajuda.close()
                break 