import pyautogui as g
import time as t
import datetime
import re
import os
import pandas as pd


from playwright.sync_api import sync_playwright, Page, Error, expect
from dotenv import load_dotenv

load_dotenv()

usuario = os.getenv("UNO_USER")
senha = os.getenv("UNO_PASSWORD")
URL = os.getenv("URL")


g.press("WIN")
t.sleep(0.5)
g.write("OpenVPN GUI")
t.sleep(0.5)
g.press("ENTER")
t.sleep(0.5)
g.hotkey('win', 'b')
t.sleep(0.5)
g.press("ENTER")


t.sleep(0.5)

# Região aproximada da bandeja


def vpn_conectada():

    regiao_de_cenexao_da_vpn_loc = _loc = (1156, 727, 684, 289)

    try:

        imagem_vpn_conectada = g.locateOnScreen(
            'PC_OpenVPN_ja_conectado_.png', region=regiao_de_cenexao_da_vpn_loc, grayscale=True, confidence=0.9)

        if imagem_vpn_conectada:

            print("VPN já está conectada.")
            raspagem_de_dados()

        else:

            print("VPN não está conectada. Iniciando processo de conexão.")

    except g.ImageNotFoundException as e:

        print(f"Imagem de VPN conectada não encontrada. Verifique se a VPN está conectada ou se a imagem de referência é adequada.")

    finally:
        t.sleep(1)


def get_bandeja_region():

    regiaoDaBandeja_loc = (1507, 902, 339, 100)

    try:

        img_PC_OpenVPN = g.locateOnScreen(
            "PC_OpenVPN.png", grayscale=True, region=regiaoDaBandeja_loc, confidence=0.9)
        t.sleep(1)
        center = g.center(img_PC_OpenVPN)
        g.rightClick(center)
        t.sleep(1)

    except g.ImageNotFoundException as e:

        print(f"Ícone do OpenVPN não encontrado na bandeja. Verifique se o OpenVPN GUI está aberto e o ícone é visível.")

        if vpn_conectada():

            print("VPN já está conectada. Prosseguindo com a raspagem de dados.")

    finally:
        t.sleep(1)


# ______________________________________________________________________________________________

# Região aproximada do botão "Escolher Backup"
def clicar_em_escolher_backup():

    try:

        regiaoDoBotao_loc = (1156, 727, 684, 289)
        PC_OpenVPN_Escolhendo_Backup = g.locateOnScreen(
            'PC_OpenVPN_Escolhendo_Backup.png', region=regiaoDoBotao_loc, grayscale=True, confidence=0.9)
        center = g.center(PC_OpenVPN_Escolhendo_Backup)
        g.leftClick(center)
        t.sleep(1)

    except g.ImageNotFoundException as e:

        print(f"Botão 'Escolher Backup' não encontrado. Verifique se a janela de seleção de backup está aberta e o botão é visível.")
        t.sleep(7)
        get_bandeja_region()
        regiaoDoBotao_loc = (1156, 727, 684, 289)
        PC_OpenVPN_Escolhendo_Backup = g.locateOnScreen(
            'PC_OpenVPN_Escolhendo_Backup.png', region=regiaoDoBotao_loc, grayscale=True, confidence=0.9)
        center = g.center(PC_OpenVPN_Escolhendo_Backup)
        g.leftClick(center)
        t.sleep(1)
        clicar_em_conectar()

    finally:
        t.sleep(1)


# ______________________________________________________________________________________________
# Região aproximada do botão "Conectar"
def clicar_em_conectar():

    try:

        botaoConectar_loc = (1156, 727, 684, 289)
        botaoConectar = g.locateOnScreen(
            'Conectar.png', grayscale=True, confidence=0.9)
        center = g.center(botaoConectar)
        g.leftClick(center)
        t.sleep(1)

    except g.ImageNotFoundException as e:

        print(f"Botão 'Conectar' não encontrado. Verifique se a janela de conexão está aberta e o botão é visível.")

    finally:
        t.sleep(1)


def acessar_pagina(page, tentativas=3, espera=8):

    for tentativa in range(1, tentativas + 1):

        try:

            print(f"Tentativa {tentativa} de {tentativas}...")
            page.goto(URL, timeout=10000)  # milesegundos = 10s
            return True

        except Error as e:
            print(f"Erro de conexão: {e}")

            if tentativa < tentativas:
                print(f"Aguardando {espera}s antes de tentar novamente...")
                t.sleep(espera)
            else:
                print("Falha após múltiplas tentativas.")
                return False


def raspagem_de_dados():

    with sync_playwright() as p:

        browser = p.chromium.launch(headless=False)
        context = browser.new_context(accept_downloads=True,)
        t.sleep(1)
        page = context.new_page()
        t.sleep(6)
        acessar_pagina(page)
        t.sleep(2)
        page.get_by_role("textbox", name="Login").fill("usuario")
        t.sleep(1)
        page.get_by_role("textbox", name="Senha").fill("senha")
        t.sleep(1)
        page.get_by_role("button", name="Entrar").click()
        print(page.title())
        page.locator("iframe[name=\"UCommerceManagerSession\"]").content_frame.locator(
            "#frameMenu").content_frame.get_by_text("SERVIÇOS").click()
        t.sleep(1)
        page.locator("iframe[name=\"UCommerceManagerSession\"]").content_frame.locator(
            "#frameMenu").content_frame.get_by_text("FUP - Ordem de Serviço").click()
        t.sleep(1)
        page.locator("iframe[name=\"UCommerceManagerSession\"]").content_frame.locator(
            "#frameCentral").content_frame.get_by_role("checkbox", name="Selecionar Todos").check()
        t.sleep(1)
        page.locator("iframe[name=\"UCommerceManagerSession\"]").content_frame.locator(
            "#frameCentral").content_frame.get_by_role("checkbox", name="Selecionar Todos").uncheck()
        t.sleep(1)
        page.locator("iframe[name=\"UCommerceManagerSession\"]").content_frame.locator(
            "#frameCentral").content_frame.locator("tr:nth-child(4) > td:nth-child(3) > #situacoes").check()
        t.sleep(1)
        page.locator("iframe[name=\"UCommerceManagerSession\"]").content_frame.locator(
            "#frameCentral").content_frame.get_by_role("button", name="Buscar").click()

        with page.expect_popup() as page1_info:
            page.locator("iframe[name=\"UCommerceManagerSession\"]")\
                .content_frame.locator("#frameCentral")\
                .content_frame.get_by_role("button", name="Exportar TXT").click()

        page1 = page1_info.value

        with page1.expect_download() as download_info:
            page1.get_by_role(
                "link",
                name=re.compile(r"^FUPOrdemDeServicos_")
            ).click()

        download = download_info.value
        caminho_arquivo = os.path.join(os.getcwd(), "ordem_de_servico.csv")
        print(caminho_arquivo)
        download.save_as(f"{caminho_arquivo}")

        t.sleep(1)
        browser.close()


get_bandeja_region()
clicar_em_escolher_backup()
clicar_em_conectar()
raspagem_de_dados()


t.sleep(1)
g.press("ENTER")

t.sleep(3)
# ______________________________________________________________________________________________

# screenWidth, screenHeight = g.size()

# ______________________________________________________________________


LOGRADOUROS = r"(RUA|AVENIDA|ALAMEDA|ESTRADA|RODOVIA|TRAVESSA|PRAÇA|VIADUTO|PARQUE|VILA|PONTE|CALÇADA)"

PADRAO_ENDERECO = re.compile(
    rf"\b{LOGRADOUROS}\b\s+.*?\s+(?:\d{{1,6}}|S/?N)\b",
    re.IGNORECASE
)


# Dicionário com os padrões e substituições
substituicoes = {

    # =========================
    # 1️⃣ LOGRADOUROS
    # =========================
    r'\bEND[.]?\b': 'ENDEREÇO',
    r'\bR[.]?\b': 'RUA',
    r'\bAV[.]?\b': 'AVENIDA',
    r'\bPC[.]?\b': 'PRAÇA',
    r'\bROD[.]?\b': 'RODOVIA',
    r'\bEST[.]?\b|\bESTR[.]?\b': 'ESTRADA',
    r'\b(AL|ALA)[.]?\b': 'ALAMEDA',
    r'\b(TRAV|TV)[.]?\b': 'TRAVESSA',
    r'\bVD[.]?\b': 'VIADUTO',
    r'\b(LGO|LGR)[.]?\b': 'LARGO',
    r'\bPQ[.]?\b': 'PARQUE',
    r'\bVL[.]?\b': 'VILA',
    r'\bPTE[.]?\b': 'PONTE',
    r'\bSTA[.]?\b': 'SANTA',
    r'\bSTO[.]?\b': 'SANTO',
    r'\bDR[.]?\b': 'DOUTOR',
    r'\bPROF[.]?\b': 'PROFESSOR',
    r'\bCEL[.]?\b': 'CORONEL',
    r'\bCAP[.]?\b': 'CAPITÃO',
    r'\bBRIG[.]?\b': 'BRIGADEIRO',
    r'\bDES[.]?\b': 'DESEMBARGADOR',
    r'\bENG[.]?\b': 'ENGENHEIRO',
    r'\bTEN[.]?\b': 'TENENTE',
    r'\bGEN[.]?\b': 'GENERAL',
    r'\b(MAL|MAR)[.]?\b': 'MARECHAL',
    r'\bDEP[.]?\b': 'DEPUTADO',
    r'\bSRA[.]?\b': 'SENHORA',
    r'\bSR[.]?\b': 'SENHOR',
    r'\bALBUQ[.]?\b': 'ALBUQUERQUE',
    r'\b(Nº[.]?|N[.]?|N-|N,|N\.O|S/N)\b': '',
    r'[.,°ºª_\*]': '',
    r'\bEND[.]?\b': 'ENDEREÇ',
    r'\bR[.]?\b': 'RUA',
    r'\bAV[.]?\b': 'AVENIDA',
    r'\bPC[.]?\b': 'PRAÇA',
    r'\bPÇA[.]?\b': 'PRAÇA',
    r'\bROD[.]?\b': 'RODOVIA',
    r'\bEST[.]?\b|\bESTR[.]?\b': 'ESTRADA',
    r'\b(AL|ALA)[.]?\b': 'ALAMEDA',
    r'\b(TRAV|TV)[.]?\b': 'TRAVESSA',
    r'\bVD[.]?\b': 'VIADUTO',
    r'\b(LGO|LGR)[.]?\b': 'LARGO',
    r'\bPQ[.]?\b': 'PARQUE',
    r'\bVL[.]?\b': 'VILA',
    r'\bPTE[.]?\b': 'PONTE',

    # =========================
    # 2️⃣ NOMES HONORÍFICOS
    # =========================
    r'\bSTA[.]?\b': 'SANTA',
    r'\bSTO[.]?\b': 'SANTO',
    r'\bDR[.]?\b': 'DOUTOR',
    r'\bPROF[.]?\b': 'PROFESSOR',
    r'\bCEL[.]?\b': 'CORONEL',
    r'\bCAP[.]?\b': 'CAPITÃO',
    r'\bBRIG[.]?\b': 'BRIGADEIRO',
    r'\bDES[.]?\b': 'DESEMBARGADOR',
    r'\bENG[.]?\b': 'ENGENHEIRO',
    r'\bTEN[.]?\b': 'TENENTE',
    r'\bGEN[.]?\b': 'GENERAL',
    r'\b(MAL|MAR)[.]?\b': 'MARECHAL',
    r'\bDEP[.]?\b': 'DEPUTADO',
    r'\bSRA[.]?\b': 'SENHORA',
    r'\bSR[.]?\b': 'SENHOR',
    r'\bALBUQ[.]?\b': 'ALBUQUERQUE',

    # =========================
    # 3️PRESIDENTE (corrigindo variações)
    # =========================
    r'\bPRESI?[.]?\b': 'PRESIDENTE',

    # =========================
    # REMOÇÃO DE MARCADORES DE NÚMERO
    # =========================
    r'\b(Nº|N°|N[.]?|N-|N,|N\.O)\b': '',

    # IMPORTANTE:
    # NÃO remover S/N aqui se quiser capturar endereço com S/N
    # Se quiser remover depois, trate separadamente

    # =========================
    # CARACTERES ESPECIAIS
    # =========================
    r'[.,°ºª_\*":\\\-]': ' '
}


def substituir_abreviacoes(texto):
    for padrao, substituto in substituicoes.items():
        texto = re.sub(padrao, substituto, texto)
    return texto


def inserir_virgula(texto):
    return re.sub(r'(\D)\s+(\d{1,6}\b)', r'\1, \2', str(texto))


# criação de colunas para meses
mesespt = {
    'January': 'Janeiro',
    'February': 'Fevereiro',
    'March': 'Março',
    'April': 'Abril',
    'May': 'Maio',
    'June': 'Junho',
    'July': 'Julho',
    'August': 'Agosto',
    'September': 'Setembro',
    'October': 'Outubro',
    'November': 'Novembro',
    'December': 'Dezembro',
}


# Criando uma chave com valores, para alteração da lingua conforme datas, meses e dias
diaspt = {
    'Sunday': 'Domingo',
    'Monday': 'Segunda-feira',
    'Tuesday': 'Terça-feira',
    'Wednesday': 'Quarta-feira',
    'Thursday': 'Quinta-feira',
    'Friday': 'Sexta-feira',
    'Saturday': 'Sábado'
}

'''#################################################################################################################################################'''
####################################################################################################################################################
'''#################################################################################################################################################'''


nome_do_arquivo = r'ordem_de_servico.csv'  # Nome do arquivo de origem

######################################################################################################################################################################################################
##################################################################


# Criando a coluna Descrição separadamente, pois o tratamento do regex se dá somente com string
coluna_descricao = []

dados = []  # criando novos dados a partir da leitura e quebras de linhas, exceto a coluna Descrição
# fazendo a leitura dos dados do arquivo de origem, para que seja tratado


# Leitura do arquivo
with open(file=f'{nome_do_arquivo}', mode='r', encoding='1252') as arquivo:
    linha = arquivo.readline()  # Cabeçalho
    linha = arquivo.readline()  # Primeira linha
    while linha.upper():
        quebra_linha = linha.strip().split(sep=';')
        nova_linha = quebra_linha[12].upper()
        nova_linha = substituir_abreviacoes(nova_linha)  # Aplica os replaces
        dados.append(nova_linha)
        coluna_descricao.append(nova_linha)
        linha = arquivo.readline()


# Regex para capturar padrões de endereço comuns em São Paulo que estejam entre possiveis nomes de endereço

regex_endereco_bloco = re.compile(r'ENDEREÇO[:]?\s*(.*)', re.IGNORECASE)

# ##### MEU REGEXXXXXXx para capturar padrões de endereço comuns em São Paulo
regex_endereco = re.compile(r'-? RUA (.*?)-|\.? RUA (.*?)-|:? R (.*?)-|:? R: (.*?)-|:? Rua: (.*?)-|:? Rua (.*?)-|:? \. Rua (.*?)-|:? \.Rua (.*?)-|:?\.Rua (.*?)-|:? R\. (.*?)-|\
-? AVENIDA (.*?)-|:? AV (.*?)(?= ?\?)|\.? AVENIDA (.*?)-|:? AV (.*?)-|:? AV: (.*?)-|:? Avenida: (.*?)-|:? Avenida (.*?)-|:? \. Avenida (.*?)-|:? \.Avenida (.*?)-|:?\.Avenida (.*?)-|:? AV\. (.*?)-|\
-? ESTRADA (.*?)-|\.? ESTRADA (.*?)-|:? E (.*?)-|:? E: (.*?)-|:? Estrada: (.*?)-|:? Estrada (.*?)-|:? \. Estrada (.*?)-|:? \.Estrada (.*?)-|:?\.Estrada (.*?)-|:? E\. (.*?)-|\
-? RODOVIA (.*?)-|\.? RODOVIA (.*?)-|:? RD (.*?)-|:? RD: (.*?)-|:? Rodovia: (.*?)-|:? Rodovia (.*?)-|:? \. Rodovia (.*?)-|:? \.Rodovia (.*?)-|:?\.Rodovia (.*?)-|:? RD\. (.*?)-|\
-? PRAÇA (.*?)-|\.? PRAÇA (.*?)-|:? P (.*?)-|:? P: (.*?)-|:? Praça: (.*?)-|:? Praça (.*?)-|:? \. Praça (.*?)-|:? \.Praça (.*?)-|:?\.Praça (.*?)-|:? P\. (.*?)-|\
-? TRAVESSA (.*?)-|\.? TRAVESSA (.*?)-|:? T (.*?)-|:? T: (.*?)-|:? Travessa: (.*?)-|:? Travessa (.*?)-|:? \. Travessa (.*?)-|:? \.Travessa (.*?)-|:?\.Travessa (.*?)-|:? T\. (.*?)-|\
-? ALAMEDA (.*?)-|\.? ALAMEDA (.*?)-|:? AL (.*?)-|:? AL: (.*?)-|:? Alameda: (.*?)-|:? Alameda (.*?)-|:? \. Alameda (.*?)-|:? \.Alameda (.*?)-|:?\.Alameda (.*?)-|:? AL\. (.*?)-|\
-? VIADUTO (.*?)-|\.? VIADUTO (.*?)-|:? VD (.*?)-|:? VD: (.*?)-|:? Viaduto|\
', re.IGNORECASE)


entregar_regex = re.compile(r'ENTREGAR|ENTREGA|DEIXAR|LEVAR', re.IGNORECASE)

# lista dos dados extraidos após identificação no padrão regex( Expressão regular)
lista = []


for texto in coluna_descricao:  # OBS: iteração na nossa lista extraída do documento de origem e NÃO DO DATAFRAME

    # encontra o padrão dentro da coluna de interesse
    match_bloco = regex_endereco_bloco.search(texto)

    if match_bloco:
        # print("Endereço encontrado:", match_bloco.group()) # print serve como validação dos dados encontrados somente referencia
        # salva os dados na variavel, o '.group()' faz com que apenas os valores sejam salvos e não o objeto
        data = match_bloco.group()
        # salva os dados encontrados no padrão regex na nossa lista lista
        lista.append(data)

    else:  # Using a simple else since the second regex is checked below
        match_endereco = regex_endereco.search(texto)

        if match_endereco:
            # print("Endereço encontrado:", match_endereco.group()) # print serve como validação dos dados encontrados somente referencia
            # salva os dados na variavel, o '.group()' faz com que apenas os valores sejam salvos e não o objeto
            data = match_endereco.group()
            # salva os dados encontrados no padrão regex na nossa lista lista
            lista.append(data)

        else:
            # print("Nenhum endereço encontrado em:", texto) # print serve como validação dos dados encontrados somente referencia
            erro = texto  # caso não seja encontrado os valores pelo regex, essa etapa retorna o mesmo valor do campo na coluna alvo
            # salva os dados na nossa lista de valores NÃO encontrados, ou seja, o mesmo valor de origem.
            lista.append(erro)


lista_str = str(lista).split(',')


lista_ = []


for texto in lista_str:  # OBS: iteração na nossa lista extraída do documento de origem e NÃO DO DATAFRAME

    # encontra o padrão dentro da coluna de interesse
    match = regex_endereco.search(texto)

    if match:
        # print("Endereço encontrado:", match.group()) # print serve como validação dos dados encontrados somente referencia
        # salva os dados na variavel, o '.group()' faz com que apenas os valores sejam salvos e não o objeto
        teste = match.group()
        # salva os dados encontrados no padrão regex na nossa lista lista
        lista_.append(teste)

    else:
        # print("Nenhum endereço encontrado em:", texto) # print serve como validação dos dados encontrados somente referencia
        erro = texto  # caso não seja encontrado os valores pelo regex, essa etapa retorna o mesmo valor do campo na coluna alvo
        # salva os dados na nossa lista de valores NÃO encontrados, ou seja, o mesmo valor de origem.
        lista_.append(erro)


dataframe = pd.read_csv(nome_do_arquivo, sep=';',
                        encoding='1252', index_col=False, quotechar='"')

df_original = pd.DataFrame(dataframe)


df = df_original

# recebendo na coluna 'Descrição' os valores de lista já tratados com regex.
df['Endereço'] = lista_

df = df.drop(columns=['Nr. Série', 'Cod Oportunidade', 'Tipo OS', 'Desc Comercial', 'Cod Produto', 'Hora Final', 'Hora Inicial', 'Hrs Apont',
                      'Hrs Prev', 'Defeito Constatado', 'Cod Colaborador', 'Cod Serviço', 'Centro de Custo', 'C. Custo'])


old_list = df["Endereço"].astype(str)

new_list = []

# SEGUNDO VALIDAÇÃO NOS TEXTOD
logradouros = r"(RUA|AVENIDA|ALAMEDA|ESTRADA|RODOVIA|TRAVESSA|PRAÇA|VIADUTO|PARQUE|VILA|PONTE|CALÇADA)"

padrao = re.compile(
    rf"\b{logradouros}\b(?: [^\d\n]+)*? (?:\d{{1,6}}|S[./-]?N)\b",
    re.IGNORECASE
)

for texto in old_list:
    texto = texto.strip()

    match = padrao.search(texto)

    if match:
        # print("Endereço encontrado:", match.group()) # print serve como validação dos dados encontrados somente referencia
        # salva os dados na variavel, o '.group()' faz com que apenas os valores sejam salvos e não o objeto
        data = match.group()

        # salva os dados encontrados no padrão regex na nossa lista lista
        new_list.append(data)

    else:
        # print("Nenhum endereço encontrado em:", texto) # print serve como validação dos dados encontrados somente referencia
        erro = texto  # caso não seja encontrado os valores pelo regex, essa etapa retorna o mesmo valor do campo na coluna alvo
        # salva os dados na nossa lista de valores NÃO encontrados, ou seja, o mesmo valor de origem.
        new_list.append(erro)


# Esse trexo de codigo é usado apenas como conderencia dos dados de endereço
###
###
df['Endereço_'] = new_list
df['Endereço_'] = df['Endereço_'].apply(inserir_virgula)
df["São Paulo"] = " - SÃO PAULO"
df["Endereço_"] = df["Endereço_"] + df["São Paulo"]
###
###
# _____________________________


# Converter para datetime, ignorando erros

df['base_Data_da_visita'] = pd.to_datetime(
    df['Dt Comprometida'],
    dayfirst=True,
    errors='coerce'
)


df['Dia da semana'] = df['base_Data_da_visita'].dt.day_name().replace(diaspt)
df['Mes'] = df['base_Data_da_visita'].dt.month_name().replace(mesespt)
df['Dia da semana'] = df['base_Data_da_visita'].dt.day_name().replace(diaspt)


# Os dados abaixo seram enviados ao telegram
df_telegtam = df[['Dia da semana', 'Dt Comprometida', 'Cliente',
                  'Endereço_', 'Atendente', 'Defeito Relatado', 'Descrição', 'Modalidade']]

df_telegtam.to_csv('dados_para_telegram.csv',
                   index=False, encoding='utf-8-sig')
