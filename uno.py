# importando as bibliotecas necessárias para a automação, manipulação de dados e expressões regulares

# biblioteca para automação de ações do mouse e teclado na maquina local
import pyautogui as g
# biblioteca para controle de tempo e pausas entre as ações, executando no test mode
import time as t
import datetime       # biblioteca para manipulação de datas e horas, responsavel por tratar as datas e extrair informações como dia da semana e mês
import re             # biblioteca para expressões regulares, utilizada para criar padrões de busca e extração de informações específicas dos textos, como endereços
import os             # biblioteca para manipulação de arquivos e diretórios, utilizada para salvar os arquivos baixados e acessar variáveis de ambiente
import pandas as pd   # biblioteca para manipulação de dados, utilizada para criar e manipular dataframes, facilitando a organização e análise dos dados extraídos

import csv
import telebot
import requests as r

from playwright.sync_api import sync_playwright, Page, Error, expect
from dotenv import load_dotenv

load_dotenv()  # Carrega as variáveis de ambiente do arquivo .env para o ambiente de execução do script


# Variável de ambiente para o nome de usuário, armazenada no arquivo .env
usuario = os.getenv("UNO_USER")
# import da variável de ambiente para a senha, armazenada no arquivo .env
senha = os.getenv("UNO_PASSWORD")
# import da variável de ambiente para a URL, armazenada no arquivo .env
URL = os.getenv("URL")

TOKEN_TELEGRAM = os.getenv("TOKEN")
# ID do usuário  que vai receber
CHAT_ID_USER_EDER = os.getenv("CHAT_ID_PESSOAL")
CHAT_ID_CORPORATIVO = os.getenv(
    "CHAT_ID_CORPORATIVO")  # ID do grupo corporativo


print(os.path.exists("PC_OpenVPN.png"))
g.press("WIN")
t.sleep(0.5)
# Digitando "OpenVPN GUI" para buscar o aplicativo no menu iniciar
g.write("OpenVPN GUI")
t.sleep(0.5)
g.press("ENTER")
t.sleep(0.5)
# Atalho para acessar a barra de tarefas, onde o ícone do OpenVPN GUI deve estar localizado
g.hotkey('win', 'b')
t.sleep(0.5)
g.press("ENTER")


t.sleep(0.5)

"""
Função para verificar se a VPN já está conectada, utilizando uma imagem de referência para identificar o status da conexão.
Se a VPN estiver conectada, a função chama a função de raspagem de dados.
Caso contrário, ela inicia o processo de conexão. Caso a imagem de VPN conectada não seja encontrada, a função exibe uma mensagem de erro e aguarda um momento antes de prosseguir para a próxima etapa. A função é projetada para ser robusta, lidando com possíveis exceções relacionadas à localização da imagem na tela.
"""


def vpn_conectada():

    regiao_de_cenexao_da_vpn_loc = _loc = (1156, 727, 684, 289)

    try:

        imagem_vpn_conectada = g.locateOnScreen('PC_OpenVPN_ja_conectado_.png', region=regiao_de_cenexao_da_vpn_loc,
                                                grayscale=True, confidence=0.9)  # Localiza a imagem da VPN "conectada"

        if imagem_vpn_conectada:

            print("VPN já está conectada.")
            raspagem_de_dados()

        else:

            print("VPN não está conectada. Iniciando processo de conexão.")

    except g.ImageNotFoundException as e:

        # Exibe uma mensagem de erro caso a imagem de VPN conectada não seja encontrada, sugerindo verificar o status da conexão ou a adequação da imagem de referência.
        print(f"Imagem de VPN conectada não encontrada. Verifique se a VPN está conectada ou se a imagem de referência é adequada.")

    finally:
        t.sleep(1)


# Encontra o icone do OpenVPN na bandeja e clica com o botão direito

def get_bandeja_region():

    regiaoDaBandeja_loc = (1507, 902, 339, 100)
    # Localiza a imagem do ícone do OpenVPN na bandeja do sistema.
    try:

        img_PC_OpenVPN = g.locateOnScreen(
            "PC_OpenVPN.png", grayscale=True, region=regiaoDaBandeja_loc, confidence=0.9)
        t.sleep(1)
        center = g.center(img_PC_OpenVPN)
        g.rightClick(center)
        t.sleep(1)

    except g.ImageNotFoundException as e:  # Caso seja exibida a mensagem de erro

        print(f"Ícone do OpenVPN não encontrado na bandeja. Verifique se o OpenVPN GUI está aberto e o ícone é visível.")

        if vpn_conectada():  # Verifica se a VPN já está conectada, caso esteja, exibe a mensagem de que a VPN já está conectada e prossegue com a raspagem de dados.

            print("VPN já está conectada. Prosseguindo com a raspagem de dados.")

    finally:
        t.sleep(1)


# ______________________________________________________________________________________________

# Região aproximada do botão "Escolher Backup" e clica para abri a janela de opções a seguir
def clicar_em_escolher_backup():

    try:

        regiaoDoBotao_loc = (1156, 727, 684, 289)
        PC_OpenVPN_Escolhendo_Backup = g.locateOnScreen(
            # Localiza o icone de "Escolher o Backup""
            'PC_OpenVPN_Escolhendo_Backup.png', region=regiaoDoBotao_loc, grayscale=True, confidence=0.9)
        center = g.center(PC_OpenVPN_Escolhendo_Backup)
        g.leftClick(center)
        t.sleep(1)

    except g.ImageNotFoundException as e:

        # Caso de erro, onde o icone não foi encontrado, é exibido a mensagem abaixo
        print(f"Botão 'Escolher Backup' não encontrado. Verifique se a janela de seleção de backup está aberta e o botão é visível.")
        t.sleep(7)  # aguardamos 7 segundos para nova verificação
        get_bandeja_region()  # É exevutado a função que localiza o ícone do OpenVPN na bandeja, para tentar localizar o botão "Escolher Backup" novamente, caso seja encontrado, a função clicar_em_escolher_backup() é chamada para clicar no botão "Escolher Backup" e prosseguir com o processo de conexão.
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
"""
Região aproximada do botão "Conectar"
Função que por fim conecta a VPN
"""


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


# ______________________________________________________________________________________________

""" 
Função para acessar a pagina, usando o Playwright para automação de navegador, com tratamento de erros e tentativas de reconexão em caso de falha na conexão. 
A função tenta acessar a URL especificada, e em caso de erro de conexão, ela aguarda um tempo definido antes de tentar novamente, até atingir o número máximo de tentativas. 
Se a conexão for bem-sucedida, a função retorna True; caso contrário, após todas as tentativas, retorna False.
"""


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


# ______________________________________________________________________________________________
"""
Função para realizar a raspagem de dados, utilizando o Playwright para automação de navegador, com etapas para acessar a página, fazer login, navegar pelos frames e realizar ações para buscar e exportar os dados necessários. 
A função inclui esperas entre as ações para garantir que os elementos estejam carregados antes de interagir com eles, e utiliza tratamento de pop-ups e downloads para salvar os arquivos exportados.
"""


def raspagem_de_dados():

    with sync_playwright() as p:

        # Inicia o navegador Chromium em modo não headless para permitir a visualização das ações realizadas durante a raspagem de dados.
        browser = p.chromium.launch(headless=False)
        # Cria um novo contexto de navegador com a opção de aceitar downloads habilitada, permitindo que os arquivos exportados sejam baixados automaticamente durante a raspagem de dados.
        context = browser.new_context(accept_downloads=True,)
        t.sleep(1)
        page = context.new_page()
        t.sleep(6)
        acessar_pagina(page)
        t.sleep(2)
        page.get_by_role("textbox", name="Login").fill(
            usuario)  # Preenchendo o campo de login com a variável de ambiente "usuario" da minha .env\
        t.sleep(1)
        # Preenchendo o campo de senha com a variável de ambiente "senha" da minha .env
        page.get_by_role("textbox", name="Senha").fill(senha)
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

        # Espera o pop-up de exportação ser aberto após clicar no botão "Exportar TXT"
        with page.expect_popup() as page1_info:
            page.locator("iframe[name=\"UCommerceManagerSession\"]")\
                .content_frame.locator("#frameCentral")\
                .content_frame.get_by_role("button", name="Exportar TXT").click()

        page1 = page1_info.value  # Acessa o pop-up de exportação

        # Aguarda o download ser iniciado após clicar no link de exportação, e salva o arquivo baixado com um nome específico na pasta atual do script.
        with page1.expect_download() as download_info:
            page1.get_by_role(
                "link",
                # O nome do arquivo de exportação pode variar, então usamos uma expressão regular para corresponder ao padrão do nome do arquivo, que começa com "FUPOrdemDeServicos_" seguido por outros caracteres.
                name=re.compile(r"^FUPOrdemDeServicos_")
            ).click()

        download = download_info.value
        # Define o caminho completo para salvar o arquivo baixado, usando o diretório atual do script e o nome "ordem_de_servico.csv"
        caminho_arquivo = os.path.join(os.getcwd(), "ordem_de_servico.csv")
        print(caminho_arquivo)
        download.save_as(f"{caminho_arquivo}")

        t.sleep(1)
        browser.close()  # Fecha o navegador após a conclusão da raspagem de dados


try:

    # Localiza o ícone do OpenVPN na bandeja e clica com o botão direito para acessar as opções de conexão
    get_bandeja_region()
except g.ImageNotFoundException as e:
    pass


try:

    # Clica no botão "Escolher Backup" para abrir a janela de seleção de backup, e em caso de erro, tenta localizar o botão novamente após aguardar um tempo.
    clicar_em_escolher_backup()
except g.ImageNotFoundException as e:
    pass


try:

    # Clica no botão "Conectar" para iniciar a conexão VPN, e em caso de erro, exibe uma mensagem indicando que o botão não foi encontrado.
    clicar_em_conectar()
except g.ImageNotFoundException as e:
    pass


try:

    # Realiza a raspagem de dados utilizando o Playwright para automação de navegador, com etapas para acessar a página, fazer login, navegar pelos frames e realizar ações para buscar e exportar os dados necessários, incluindo tratamento de pop-ups e downloads.
    raspagem_de_dados()
except Error as e:
    pass


t.sleep(1)
g.press("ENTER")

t.sleep(3)


"""
# ______________________________________________________________________________________________
#####
##### ETAPA DE TRATAMENTO DE DADOS EXTRAÍDOS, COM EXPRESSÕES REGULARES 
##### PARA CAPTURA DE ENDEREÇOS E MANIPULAÇÃO DE DATAS PARA EXTRAÇÃO DE INFORMAÇÕES COMO DIA DA SEMANA E MÊS, 
##### ALÉM DE PREPARAR OS DADOS PARA ENVIO AO TELEGRAM. #####
#####
# ______________________________________________________________________
"""
"""
# Padrão regex para capturar endereços comuns em São Paulo, considerando variações na forma como os endereços são escritos, 
# incluindo possíveis abreviações e diferentes formatações. O padrão é projetado para ser flexível, 
# permitindo a captura de endereços mesmo que estejam embutidos em outros textos ou tenham formatações inconsistentes.
"""

LOGRADOUROS = r"(RUA|AVENIDA|ALAMEDA|ESTRADA|RODOVIA|TRAVESSA|PRAÇA|VIADUTO|PARQUE|VILA|PONTE|CALÇADA)"


PADRAO_ENDERECO = re.compile(
    rf"\b{LOGRADOUROS}\b\s+.*?\s+(?:\d{{1,6}}|S/?N)\b",
    re.IGNORECASE
)

# ______________________________________________________________________________________

# Dicionário com os padrões e substituições
substituicoes = {

    # =========================
    # LOGRADOUROS
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
    # NOMES HONORÍFICOS
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
    # PRESIDENTE (corrigindo variações)
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

# Função para substituir as abreviações e caracteres especiais no texto,
# utilizando o dicionário de substituições definido acima.
# A função percorre cada padrão e substituição no dicionário,
# aplicando as substituições ao texto usando expressões regulares, e retorna o texto modificado.


def substituir_abreviacoes(texto):
    for padrao, substituto in substituicoes.items():
        texto = re.sub(padrao, substituto, texto)
    return texto


# Função para inserir uma vírgula entre o nome do logradouro e o número do endereço,
# utilizando uma expressão regular para identificar o padrão de logradouro seguido por um número, e substit
def inserir_virgula(texto):
    return re.sub(r'(\D)\s+(\d{1,6}\b)', r'\1, \2', str(texto))


# criação de colunas para meses e dias da semana em português,
# utilizando dicionários para mapear os nomes dos meses e dias da semana em inglês para suas equivalentes em português.
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

# IMPORTANDO A BASE DE DADOS EXTRAÍDA PARA TRATAMENTO, UTILIZANDO O PANDAS PARA CRIAR UM DATAFRAME A PARTIR DO ARQUIVO CSV EXPORTADO,
nome_do_arquivo = r'ordem_de_servico.csv'  # Nome do arquivo de origem

######################################################################################################################################################################################################
##################################################################


# Criando a coluna Descrição separadamente, pois o tratamento do regex se dá somente com string
coluna_descricao = []

dados = []  # criando novos dados a partir da leitura e quebras de linhas, exceto a coluna Descrição
# fazendo a leitura dos dados do arquivo de origem, para que seja tratado


# Leitura do arquivo
# Abrindo o arquivo CSV exportado, utilizando a codificação '1252' para garantir a leitura correta dos caracteres acentuados e especiais presentes no arquivo, e criando um objeto de arquivo para leitura.
with open(file=f'{nome_do_arquivo}', mode='r', encoding='1252') as arquivo:
    # Lê a primeira linha do arquivo, que geralmente contém os cabeçalhos das colunas, e armazena na variável 'linha' para iniciar o processo de leitura dos dados.
    linha = arquivo.readline()
    # Lê a segunda linha do arquivo, que é a primeira linha de dados, e armazena na variável 'linha' para iniciar o processo de leitura dos dados. A leitura continua dentro do loop while até que todas as linhas sejam processadas.
    linha = arquivo.readline()
    while linha.upper():  # Enquanto a linha lida for diferente de uma string vazia (indicando que ainda há dados para ler), o loop continua a processar cada linha do arquivo e coloca as letras em maiúsculo para garantir a consistência no tratamento dos dados, especialmente para a aplicação de expressões regulares e substituições.
        # Remove os espaços em branco no início e no final da linha, e depois divide a linha em uma lista de valores usando o ponto e vírgula como separador, armazenando o resultado na variável 'quebra_linha' para acessar os dados de cada coluna individualmente.
        quebra_linha = linha.strip().split(sep=';')
        # Acessa o valor da coluna de interesse (neste caso, a coluna 12, que é a "Descrição") e converte o texto para maiúsculo para garantir a consistência no tratamento dos dados, especialmente para a aplicação de expressões regulares e substituições.
        nova_linha = quebra_linha[12].upper()
        nova_linha = substituir_abreviacoes(nova_linha)  # Aplica os replaces
        # Adiciona a nova linha processada à lista de dados, que será utilizada posteriormente para criar o dataframe e realizar análises adicionais.
        dados.append(nova_linha)
        # Adiciona a nova linha processada à lista 'coluna_descricao', que é uma lista separada para armazenar os valores da coluna de descrição, facilitando o tratamento específico dessa coluna com expressões regulares para extração de endereços e outras informações relevantes.
        coluna_descricao.append(nova_linha)
        # Lê a próxima linha do arquivo para continuar o processo de leitura e tratamento dos dados, repetindo o loop até que todas as linhas sejam processadas.
        linha = arquivo.readline()

# ____________________________________________________________________________________________________________________
"""
# Regex para capturar padrões de endereço comuns em São Paulo que estejam entre possiveis nomes de endereço
# O regex é projetado para ser flexível, permitindo variações na forma como os endereços são escritos, 
# incluindo possíveis abreviações e diferentes formatações, 
# garantindo que os endereços sejam extraídos corretamente,
# mesmo que estejam embutidos em outros textos ou tenham formatações inconsistentes.
"""

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

        # salva os dados na variavel, o '.group()' faz com que apenas os valores sejam salvos e não o objeto
        data = match_bloco.group()
        # salva os dados encontrados no padrão regex na nossa lista lista
        lista.append(data)

    else:  # Using a simple else since the second regex is checked below
        match_endereco = regex_endereco.search(texto)

        if match_endereco:

            # salva os dados na variavel, o '.group()' faz com que apenas os valores sejam salvos e não o objeto
            data = match_endereco.group()
            # salva os dados encontrados no padrão regex na nossa lista lista
            lista.append(data)

        else:

            erro = texto  # caso não seja encontrado os valores pelo regex, essa etapa retorna o mesmo valor do campo na coluna alvo
            # salva os dados na nossa lista de valores NÃO encontrados, ou seja, o mesmo valor de origem.
            lista.append(erro)

# Convertendo a lista de dados extraídos para uma string e depois dividindo em uma nova lista usando a vírgula como separador, para facilitar o tratamento dos dados posteriormente.
lista_str = str(lista).split(',')

# Criando uma nova lista para armazenar os dados tratados com o regex mais abrangente, que captura padrões de endereço comuns em São Paulo, mesmo que estejam entre possíveis nomes de endereço ou outros elementos de texto. O regex foi projetado para ser flexível, permitindo variações na forma como os endereços são escritos, incluindo possíveis abreviações e diferentes formatações, garantindo que os endereços sejam extraídos corretamente, mesmo que estejam embutidos em outros textos ou tenham formatações inconsistentes.
lista_ = []

# Iterando sobre a lista de strings extraídas do documento de origem, aplicando o regex para capturar os padrões de endereço e armazenando os resultados na nova lista 'lista_'.
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

# _______________________________________________________________________________

"""
# Criando um dataframe a partir do arquivo CSV exportado, 
# utilizando o pandas para facilitar a manipulação e análise dos dados extraídos, 
# e armazenando o resultado em um novo dataframe chamado 'df_original' para preservar os dados originais antes de realizar as transformações e tratamentos necessários.
"""

dataframe = pd.read_csv(nome_do_arquivo, sep=';',
                        encoding='1252', index_col=False, quotechar='"')

df_original = pd.DataFrame(dataframe)


df = df_original

# recebendo na coluna 'Descrição' os valores de lista já tratados com regex.
df['Endereço'] = lista_

# removendo as colunas que "não são necessárias" para a análise e envio ao Telegram, mantendo apenas as colunas relevantes para o contexto do trabalho, como 'Endereço', 'Dt Comprometida', 'Cliente', 'Atendente', 'Defeito Relatado', 'Descrição' e 'Modalidade'.
df = df.drop(columns=['Nr. Série', 'Cod Oportunidade', 'Tipo OS', 'Desc Comercial', 'Cod Produto', 'Hora Final', 'Hora Inicial', 'Hrs Apont',
                      'Hrs Prev', 'Defeito Constatado', 'Cod Colaborador', 'Cod Serviço', 'Centro de Custo', 'C. Custo'])

# Criando uma nova coluna 'Endereço_' a partir da coluna 'Endereço', aplicando o regex mais abrangente para capturar padrões de endereço comuns em São Paulo, mesmo que estejam entre possíveis nomes de endereço ou outros elementos de texto. O regex foi projetado para ser flexível, permitindo variações na forma como os endereços são escritos, incluindo possíveis abreviações e diferentes formatações, garantindo que os endereços sejam extraídos corretamente, mesmo que estejam embutidos em outros textos ou tenham formatações inconsistentes.
old_list = df["Endereço"].astype(str)

# Criando uma nova lista para armazenar os dados tratados com o regex mais abrangente, que captura padrões de endereço comuns em São Paulo, mesmo que estejam entre possíveis nomes de endereço ou outros elementos de texto. O regex foi projetado para ser flexível, permitindo variações na forma como os endereços são escritos, incluindo possíveis abreviações e diferentes formatações, garantindo que os endereços sejam extraídos corretamente, mesmo que estejam embutidos em outros textos ou tenham formatações inconsistentes.
new_list = []

# biblioteca para manipulação de dados, utilizada para criar e manipular dataframes, facilitando a organização e análise dos dados extraídos
texto = "SEGUNDO VALIDAÇÃO NOS TEXTO, APLICANDO UM REGEX MAIS ABRANGENTE PARA CAPTURAR PADRÕES DE ENDEREÇO COMUM EM SÃO PAULO, MESMO QUE ESTEJAM ENTRE POSSÍVEIS NOMES DE ENDEREÇO OU OUTROS ELEMENTOS DE TEXTO. O REGEX FOI PROJETADO PARA SER FLEXÍVEL, PERMITINDO VARIAÇÕES NA FORMA COMO OS ENDEREÇOS SÃO ESCRITOS, INCLUINDO POSSÍVEIS ABREVIAÇÕES E DIFERENTES FORMATAÇÕES. O OBJETIVO É GARANTIR QUE OS ENDEREÇOS SEJAM EXTRAÍDOS CORRETAMENTE, MESMO QUE ESTEJAM EMBUTIDOS EM OUTROS TEXTOS OU TENHAM FORMATAÇÕES INCONSISTENTES."
texto = texto.islower()
print(texto)
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


# Esse trecho de codigo pode-se ser usado para validar os dados de endereço, porem ele não extrai os dados de bairro, apenas faz um merge com "SP", por tanto, eu recomendo usar o campo "Endereço".
###
###
df['Endereço_'] = new_list
# Aplicando a função para inserir uma vírgula entre o nome do logradouro e o número do endereço, utilizando uma expressão regular para identificar o padrão de logradouro seguido por um número, e substituindo por uma vírgula para facilitar a leitura e análise dos endereços extraídos.
df['Endereço_'] = df['Endereço_'].apply(inserir_virgula)
# Criando uma nova coluna "São Paulo" com o valor fixo " - SÃO PAULO" para ser concatenada posteriormente com a coluna "Endereço_", adicionando a informação de localização aos endereços extraídos, facilitando a identificação de que os endereços estão localizados em São Paulo, e melhorando a clareza dos dados para análises futuras.
df["São Paulo"] = " - SÃO PAULO"
# Concatenando a coluna "Endereço_" com a string " - SÃO PAULO" para adicionar a informação de localização aos endereços extraídos, facilitando a identificação de que os endereços estão localizados em São Paulo, e melhorando a clareza dos dados para análises futuras.
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

# Criando colunas para dia da semana e mês em português, utilizando os dicionários de mapeamento definidos anteriormente para traduzir os nomes dos dias da semana e meses do inglês para o português, facilitando a análise temporal dos dados extraídos.
# Criando uma nova coluna "Dia da semana" a partir da coluna "base_Data_da_visita", utilizando o método dt.day_name() para extrair o nome do dia da semana em inglês, e depois aplicando o método replace() com o dicionário de mapeamento "diaspt" para traduzir os nomes dos dias da semana do inglês para o português, facilitando a análise temporal dos dados extraídos.
df['Dia da semana'] = df['base_Data_da_visita'].dt.day_name().replace(diaspt)
# Criando uma nova coluna "Mes" a partir da coluna "base_Data_da_visita", utilizando o método dt.month_name() para extrair o nome do mês em inglês, e depois aplicando o método replace() com o dicionário de mapeamento "mesespt" para traduzir os nomes dos meses do inglês para o português, facilitando a análise temporal dos dados extraídos.
df['Mes'] = df['base_Data_da_visita'].dt.month_name().replace(mesespt)
# Criando uma nova coluna "Dia da semana" a partir da coluna "base_Data_da_visita", utilizando o método dt.day_name() para extrair o nome do dia da semana em inglês, e depois aplicando o método replace() com o dicionário de mapeamento "diaspt" para traduzir os nomes dos dias da semana do inglês para o português, facilitando a análise temporal dos dados extraídos.
df['Dia da semana'] = df['base_Data_da_visita'].dt.day_name().replace(diaspt)

# capturando a data de hoje
hoje = pd.to_datetime('today')

# Criando uma nova  variavel, método dt.weekday() para extrair o número do dia da semana (0 para segunda-feira, 1 para terça-feira, etc.)
my_day_number = hoje.weekday()

# Criando uma nova variavel, já aplicando o método de calculo do dia+1
amanha = hoje + pd.Timedelta(days=1)

# Criando uma nova variavel, já aplicando o método de calculo do dia+2
depois_de_amanha = hoje + pd.Timedelta(days=2)

# Criando uma nova variavel, já aplicando o método de calculo do dia+3, esse metodo nós aplicaremos nas sextas-feiras
quatro_dias_a_frente = hoje + pd.Timedelta(days=3)


# formatando com .strftime('%d/%m/%Y') para data no modelo brasileiro dd/mm/aaaa
hoje = hoje.strftime('%d/%m/%Y')
amanha = amanha.strftime('%d/%m/%Y')
depois_de_amanha = depois_de_amanha.strftime('%d/%m/%Y')
quatro_dias_a_frente = quatro_dias_a_frente.strftime('%d/%m/%Y')


# aplicando o filtro para seleção dos dados com data de visita igual a data atual ou futura
df_telegram = df.loc[df['Dt Comprometida'] >= hoje]

if my_day_number == 4:  # 4 = Sexta-feira

    visitas = df_telegram.loc[
        (df_telegram['Dt Comprometida'] >= hoje) &
        (df_telegram['Dt Comprometida'] <= quatro_dias_a_frente)]
    print("ESTAMOS ENCAMINHANDO AS VISITAS DE HOJE E DE SEGUNDA-FEIRA")

else:

    visitas = df_telegram.loc[
        (df_telegram['Dt Comprometida'] >= hoje) &
        (df_telegram['Dt Comprometida'] <= depois_de_amanha)
    ]
    print("ESTAMOS ENCAMINHANDO AS VISITAS DE HOJE E AMANHÃ")


# Os dados abaixo seram enviados ao telegram
df_telegram = visitas[['Dia da semana', 'Dt Comprometida', 'Cliente',
                       'Endereço_', 'Atendente', 'Defeito Relatado', 'Descrição', 'Modalidade']]


# criando uma nova variavel, com token do telegram, para envio das mensagens
bot = telebot.TeleBot(TOKEN_TELEGRAM)

# criando uma nova variavel, com caminho do ditorio dos dados de agendamento
caminho_csv = os.getenv('caminho_do_arquivo_dos_dados_de_agendamento')


def ler_csv(caminho_csv):
    tarefas = []

    with open(caminho_csv, mode="r", encoding="utf-8", newline="") as arquivo:
        leitor = csv.DictReader(arquivo, delimiter=";")

        for row in leitor:
            tarefas.append(row)

    return tarefas


def ler_csv(caminho_csv):
    tarefas = []

    with open(caminho_csv, mode="r", encoding="utf-8", newline="") as arquivo:
        leitor = csv.DictReader(arquivo, delimiter=";")

        for row in leitor:
            tarefas.append(row)

    return tarefas


def enviar_mensagem(texto):

    mensagens = []
    for tx in texto:

        msg = (
            f"*DIA*: _{tx['\ufeffDia da semana']}_\n\n"
            f"*DATA DA VISITA*: _{tx['Dt Comprometida']}_\n"
            f"*CLIENTE*: _{tx['Cliente']}_\n\n"
            f"*ENDEREÇO*: _{tx['Endereço_']}_\n\n"
            f"*ANALISTA*: _{tx['Atendente']}_\n\n"
            f"*OBS*: _{tx['Defeito Relatado']}_\n\n"
            f"*DESCRIÇÃO*: _{tx['Descrição']}_\n\n"
            f"*TIPO*: _{tx['Modalidade']}_\n"

        )
        print(msg)
        t.sleep(0.5)

        URL

        CHAT_IDS = [
            CHAT_ID_USER_EDER,
            CHAT_ID_CORPORATIVO,
        ]

        for chat_id in CHAT_IDS:
            bot.send_message(chat_id, msg, parse_mode="Markdown")

    return "\n".join(mensagens)


tarefas = ler_csv(caminho_csv)
enviar_mensagem(tarefas)

tarefas = ler_csv(caminho_csv)
enviar_mensagem(tarefas)
