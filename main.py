# ###################################################################################################################
# Os dados dos alunos estão armazenados em uma planilha,
# para que os alunos sejam identificados é necessário acessar esta planilha.
# Crie uma planilha em excel com os campos: 'TAXA', 'NOME DO ALUNO', 'TELEFONE', 'CPF' e 'SITUAÇÃO'.
# Só serão identificados os alunos que estiverem com o campo 'TAXA' marcado como "PG" e 'SITUAÇÃO' como "PENDENTE".
# É importante que a planilha seja preenchida sempre com Uppercase.
# Após a execução do script, altere a situação na planilha para que o sistema não tente marcar o exame novamente.
# A planilha deverá ser compartilhada através de link e estar disponível no OneDrive do usuário.
# Para que a planilha fosse acessada, foi utilizada a API REST do OneDrive.
# # https://docs.microsoft.com/pt-br/onedrive/developer/rest-api/api/shares_get?view=odsp-graph-online
# É necessário uma assinatura do Office 365. Se for uma versão business, é necessário o contexto de autenticação.
# Verificar documentação API REST do SharePoint.
# A notificação é feita via WhatsApp utilizando API do ZapFácil (requer assinatura / verificar documentação ZapFácil).
# ####################################################################################################################


# Bibliotecas utilizadas

import base64
import os
import requests
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from subprocess import CREATE_NO_WINDOW
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep, strftime
from sys import exit



# Acessa a Planilha

def create_onedrive_direct_download(onedrive_link_file):
    # Acessa a a API REST do one drive
    data_bytes64 = base64.b64encode(bytes(onedrive_link_file, 'utf-8'))
    data_bytes64_string = data_bytes64.decode('utf-8') \
        .replace('/', '_').replace('+', '-').rstrip("=")
    result_url = f"https://api.onedrive.com/v1.0/shares/u!{data_bytes64_string}/root/content"
    return result_url


# Faz o download da planilha compartilhada através do link, na memória.

# Compartilhe a planilha com link e defina a variável onerivelink_link = "link"
onedrive_link = "your link here"
direct_link = create_onedrive_direct_download(onedrive_link)
data = pd.read_excel(direct_link, index_col=False)
data_df = data.loc[(data['TAXA'] == "PG") & (data['SITUAÇÃO'] == "PENDENTE"), ["NOME DO ALUNO", "TELEFONE", "CPF"]]

# Inicia o WebDriver

# Em caso de erro, verificar se a versão do Chrome e do WebDriver são compatíveis
# Caso o código seja compilado para um executável, o webdriver.exe deve ser copiado para a pasta do programa.

# Faz a instalação da extensão do plugin Signa - Prodemge 1.4.0.0
executable_path = "./"
os.environ["webdriver.chrome.driver"] = executable_path
chrome_options = webdriver.ChromeOptions()
chrome_options.add_extension('Signa - Prodemge 1.4.0.0.crx')

# Remove Prompt do Driver
chrome_service = ChromeService("./chromedriver")
chrome_service.creationflags = CREATE_NO_WINDOW

driver = webdriver.Chrome(options=chrome_options, service=chrome_service)


# É necessário instalar o plugin Signa da Prodemge para que o sistema Detrannet possa ser acessado.
# Verificar se na máquina já tem instalado o executável Signa Prodemge,
# disponível em: "https://wwws.prodemge.gov.br/images/Aplicativos/Signa-2.2.00-Prodemge.exe"


# início
driver.get(
    # Página de login do Detrannet
    "https://empresas.detran.mg.gov.br/sdaf/paginas/ssc/login.asp"
    )

# Defina o tempo necessário para a validação do certificado digital
sleep(15) # tempo em segundos

driver.get(
    # Página de marcação de legislação
    "https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp"
    )

# Sem uso no momento
def verifica_autenticacao():
    return


# local para onde será feita a notificação do agendamento
def send_message():
    requests.post("https://api.painel.zapfacil.com/api/webhooks/id", json=aluno)


# Variáveis de objetos do Site
# Verifique a unidade de agendamento em sua cidade
textbox_cpf = '/html/body/center/form/table/tbody/tr[2]/td[2]/input'
listbox_unidade = '//*[@id="localExame"]'
unidade = "unidade de atendimento"  # Alterar para a unidade de agendamento da sua cidade
listbox_turno = '//*[@id="turno"]'
listbox_data = '//*[@id="ajaxInput"]/select'
voltar = '/html/body/center/form/input[2]'
confirma = '/html/body/center/form/input[2]'
confirma_exame = '/html/body/center/form/input[3]'
limpar = '/html/body/center/form/input[3]'
voltarSenha = '/html/body/center/form/div/table/tbody/tr/td/table/tbody/tr[2]/td/input[2]'
date = strftime("%d %b %Y")


def print_screen():
    # print screen
    driver.find_element(
        By.CSS_SELECTOR, 'body').screenshot(f'{my_path}/{nome}_{cpf}.png')
    sleep(1)


def make_dir():
    # Cria um diretório no caminho especificado
    if not os.path.isdir(my_path):
        os.makedirs(my_path)


def resolve_erro():
    # Apaga os dados inseridos no sistema
    try:
        driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp")
        driver.find_element(By.XPATH, limpar).click()
    except Exception:
        # Se houver uma exceção o sistema entrará em loop.
        # Exemplo: Se o sistema for iniciado enquanto a marcação estiver fechada,
        # o sistema continuará tentando marcar até que ela esteja aberta ou cumpra a cota máxima de excessões.
        global turno
        turno = 'manhã'  # Redefine o turno
        mudar_turno()  # Os dados serão inseridos no período da manhã novamente


def mudar_turno():
    # Insere os dados novamente, porém, no próximo turno disponível.
    try:
        # Caso <try> seja atendido, vai para o próximo aluno
        driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp")
        inserir_dados()
    except Exception:
        # O sistema retorna e se mantém em loop até a condição ser atendida.
        global my_path
        my_path = f'./marcacaoLegislacao_logFile/{date}/'  # Define o caminho que a pasta de log será criada
        make_dir()  # Cria a pasta de log
        print_screen()  # Cria o log file
        resolve_erro()


def inserir_dados():
    # insere os dados na página de marcação, tenta marcar o exame e retorna
    driver.find_element(By.XPATH, textbox_cpf).send_keys(f"{cpf}", Keys.TAB)
    driver.find_element(By.XPATH, listbox_unidade).send_keys(unidade)
    if turno == 'manhã':
        driver.find_element(By.XPATH, listbox_turno).send_keys(Keys.DOWN, Keys.TAB)
        sleep(1)
        send_message()
    else:
        # <Keys.DOWN> agora é pressionado duas vezes.
        driver.find_element(By.XPATH, listbox_turno).send_keys(Keys.DOWN, Keys.DOWN, Keys.TAB)
        sleep(1)
        driver.find_element(By.XPATH, listbox_data).send_keys(Keys.DOWN)
        driver.find_element(By.XPATH, confirma).click()
        driver.find_element(By.XPATH, confirma_exame).click()
        driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp")
        sleep(1)
        send_message()


# Marca os exames

for i, nome in enumerate(data_df['NOME DO ALUNO']):
    telefone_df: str = data_df['TELEFONE'].iloc[i]
    cpf_df: str = data_df['CPF'].iloc[i]
    # O sistema lê a coluna CPF como um INT e elimina os zeros iniciais do CPF
    cpf: str = '%011d' % int(cpf_df)  # Completa o cpf com os zeros elimidados
    turno = 'manhã'
    #json aluno para POST
    aluno = {
        "nome":f"{nome}",
        "telefone":f"{telefone_df}",
        "cpf":f"{cpf}"
        }
    try:
        inserir_dados()
    except Exception:
        turno = 'tarde'
        try:
            mudar_turno()
        except Exception:
            exit()  # Caso ocorra uma excessão dentro de uma excessão o sistema encerra.


# Imprime as senhas

for i, nome in enumerate(data_df['NOME DO ALUNO']):
    cpf_df: str = data_df['CPF'].iloc[i]
    # A sistema lê a coluna CPF como um INT e elimina os zeros iniciais do CPF
    cpf: str = '%011d' % int(cpf_df)  # Completa o cpf com os zeros elimidados
    # Vai para a página de impressão de senhas
    driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0327.asp")
    try:
        # Tenta imprimir as senhas. Caso não agendado, será impresso o motivo e o sistema retornará um erro.
        driver.find_element(By.XPATH, textbox_cpf).send_keys(f"{cpf}" + Keys.ENTER)
        my_path = f'./senhas/{date}/'  # Define o caminho que a pasta de senhas será criada
        make_dir()
        sleep(2)
        print_screen()
        driver.find_element(By.XPATH, voltarSenha).click()
    except Exception:
        # Caso erro, volta e pula para o próximo aluno
        driver.find_element(By.XPATH, '/html/body/center/input').click()
        driver.find_element(By.XPATH, voltar).click()
        sleep(1)

# Fecha o WebDriver e encerra o programa sem erros
driver.quit()
exit()
