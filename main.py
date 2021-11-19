import base64
import pandas as pd
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from sys import exit


# Os dados dos alunos estão armazenados em uma planilha,
# para que os alunos sejam identificados é necessário acessar esta planilha.
# Crie uma planilha em excel com os campos: 'TAXA', 'NOME DO ALUNO', 'CPF' e 'SITUAÇÃO'.
# Só serão identificados os alunos que estiverem com o campo 'TAXA' marcado como "PG" e 'SITUAÇÃO' como "PENDENTE".
# É importante que a planilha seja preenchida sempre com Uppercase.
# Após a execução do script, altere a situação para que o sistema não tente marcar o exame novamente.
# A planilha deverá ser compartilhada através de link e está disponível no OneDrive do usuário.
# Para que a planilha fosse acessada, foi utilizada a API REST do OneDrive.
# É necessário uma assinatura do Office 365. Se for uma versão business, é necessário o contexto de autenticação.
# Verificar documentação API REST do SharePoint.

# Acessa a Planilha

def create_onedrive_direct_download(onedrive_link_file):
    # Acessa a a API REST do one drive
    data_bytes64 = base64.b64encode(bytes(onedrive_link_file, 'utf-8'))
    data_bytes64_string = data_bytes64.decode('utf-8') \
        .replace('/', '_').replace('+', '-').rstrip("=")
    result_url = f"https://api.onedrive.com/v1.0/shares/u!{data_bytes64_string}/root/content"
    return result_url


# Faz o download da planilha compartilhada através do link, na memória.
# Compartilhe a planilha com link e defina a variável onerivelink_link = "<link>"

onedrive_link = "<link>"
direct_link = create_onedrive_direct_download(onedrive_link)
data = pd.read_excel(direct_link, index_col=False)
data_df = data.loc[(data['TAXA'] == "PG") & (data['SITUAÇÃO'] == "PENDENTE"), ["NOME DO ALUNO", "CPF"]]

# Inicia o WebDriver

# Em caso de erro, verificar se a versão do Chrome e do WebDriver são compatíveis
# Caso o código seja compilado para um executável, o webdriver.exe deve ser copiado para a pasta do programa.
driver = webdriver.Chrome()

# É necessário instalar o plugin Signa da Prodemge para que o sistema Detrannet possa ser acessado.
# Verificar se na máquina já tem instalado o executável Signa Prodemge,
# disponível em: "https://wwws.prodemge.gov.br/images/Aplicativos/Signa-2.2.00-Prodemge.exe"
driver.get(
    "https://chrome.google.com/webstore/detail/signa-prodemge/idbpfpeogbhifooiagnbbdbffplkfcke?hl=pt-BR")

# Defina o tempo necessário para que o código aguarde a instalação do plugin
time.sleep(6)

# indica qual página deve ser acessada
driver.get(
    # Página de login do Detrannet
    "https://empresas.detran.mg.gov.br/sdaf/paginas/ssc/login.asp")
# Defina o tempo necessário para a validação do certificado digital
time.sleep(10)

driver.get(
    # Página de marcação de legislação
    "https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp")

# Variáveis de objetos do Site
# Verifique a unidade de agendamento em sua cidade
textbox_cpf = '/html/body/center/form/table/tbody/tr[2]/td[2]/input'
listbox_unidade = '//*[@id="localExame"]'
unidade = "<Unidade>"  # Alterar para a unidade de agendamento da sua cidade
listbox_turno = '//*[@id="turno"]'
listbox_data = '//*[@id="ajaxInput"]/select'
voltar = '/html/body/center/form/input[2]'
confirma = '/html/body/center/form/input[2]'
confirma_exame = '/html/body/center/form/input[3]'
limpar = '/html/body/center/form/input[3]'
voltarSenha = '/html/body/center/form/div/table/tbody/tr/td/table/tbody/tr[2]/td/input[2]'
date = time.strftime("%d %b %Y")


def print_screen():
    # print screen
    element = driver.find_element(By.CSS_SELECTOR, 'body')
    element.screenshot(f'{my_path}/{nome}_{cpf}.png')
    time.sleep(1)


def make_dir():
    # Cria um diretório no caminho especificado
    if not os.path.isdir(mypath):
        os.makedirs(mypath)


def resolve_erro():
    # Apaga os dados inseridos no sistema
    try:
        driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp")
        driver.find_element(By.XPATH, limpar).click()
    except Exception:
        # Se <try> == FALSE o sistema irá para a próxima etapa,
        # caso a próxima etapa também seja == FALSE o sistema entrará em loop.
        print_screen()
        mudar_turno()


def mudar_turno():
    # Insere os dados novamente, porém, no período da tarde.
    try:
        # Se <try> == TRUE, vai para o próximo aluno
        driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp")
        driver.find_element(By.XPATH, textbox_cpf)
        driver.find_element(By.XPATH, textbox_cpf).send_keys(f"{cpf}", Keys.TAB)
        driver.find_element(By.XPATH, listbox_unidade).send_keys(unidade)
        # <Keys.DOWN> agora é pressionado duas vezes.
        driver.find_element(By.XPATH, listbox_turno).send_keys(Keys.DOWN, Keys.DOWN, Keys.TAB)
        time.sleep(1)
        driver.find_element(By.XPATH, listbox_data).send_keys(Keys.DOWN)
        driver.find_element(By.XPATH, confirma).click()
        driver.find_element(By.XPATH, confirma_exame).click()
        driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp")
        time.sleep(1)
    except Exception:
        # O sistema retorna e se mantém em loop até o primeiro <try> ser atendido.
        # Exemplo: Se o sistema for iniciado enquanto a marcação estiver fechada,
        # o sistema continuará tentando marcar até que ela esteja aberta ou cumpra a cota máxima de excessões.
        make_dir()
        print_screen()
        resolve_erro()


def inserir_dados():
    # insere os dados na página de marcação, tenta marcar o exame e retorna
    driver.find_element(By.XPATH, textbox_cpf).send_keys(f"{cpf}", Keys.TAB)
    driver.find_element(By.XPATH, listbox_unidade).send_keys(unidade)
    driver.find_element(By.XPATH, listbox_turno).send_keys(Keys.DOWN, Keys.TAB)
    time.sleep(1)
    driver.find_element(By.XPATH, listbox_data).send_keys(Keys.DOWN)
    driver.find_element(By.XPATH, confirma).click()
    driver.find_element(By.XPATH, confirma_exame).click()
    driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp")
    time.sleep(1)


# Marca os exames

for i, nome in enumerate(data_df['NOME DO ALUNO']):
    cpf_df: str = data_df['CPF'].iloc[0 + i]
    # A sistema lê a coluna CPF como um INT e elimina os zeros iniciais do CPF
    num_cpf_char = 11  # quantidade de números em um CPF
    if len(f'{cpf_df}') < num_cpf_char:  # Caso o cpf possua menos de 11 caracteres
        cpf: str = '%011d' % int(cpf_df)  # A quantidade é completada com zeros
    else:
        cpf: str = cpf_df  # O cpf não inicia com zeros
    try:
        inserir_dados()
    except Exception:
        try:
            mypath = f'./marcacaoLegislacao_logFile/{date}/'  # Define o caminho que a pasta de log será criada
            mudar_turno()
            resolve_erro()
        except Exception:
            exit()  # Caso ocorra uma excessão dentro de uma excessão o sistema encerra.


# Imprime as senhas

for i, nome in enumerate(data_df['NOME DO ALUNO']):
    cpf_df: str = data_df['CPF'].iloc[0 + i]
    # A sistema lê a coluna CPF como um INT e elimina os zeros iniciais do CPF
    num_cpf_char = 11  # quantidade de números em um CPF
    if len(f'{cpf_df}') < num_cpf_char:  # Caso o cpf possua menos de 11 caracteres
        cpf: str = '%011d' % int(cpf_df)  # A quantidade é completada com zeros
    else:
        cpf: str = cpf_df  # O cpf não inicia com zeros
    # Vai para a página de impressão de senhas
    driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0327.asp")
    try:
        # Tenta imprimir as senhas
        driver.find_element(By.XPATH, textbox_cpf).send_keys(f"{cpf}" + Keys.ENTER)
        my_path = f'./senhas/{date}/'  # Define o caminho que a pasta de senhas será criada
        make_dir()
        time.sleep(2)
        print_screen()
        driver.find_element(By.XPATH, voltarSenha).click()
    except Exception:
        # Caso não consiga, imprime o motivo e pula para o próximo aluno
        driver.find_element(By.XPATH, '/html/body/center/input').click()
        driver.find_element(By.XPATH, voltar).click()
        time.sleep(1)

# Fecha o WebDriver e encerra o programa sem erros
driver.quit()
exit()
