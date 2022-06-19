# ###################################################################################################################
# Os dados dos alunos estão armazenados em uma planilha de Excel,
# para que os alunos sejam identificados é necessário acessar esta planilha.
# Crie uma planilha em Excel com os campos: 'TAXA', 'NOME DO ALUNO', 'TELEFONE', 'CPF' e 'SITUAÇÃO'.
# Só serão identificados os alunos que estiverem com o campo 'TAXA' marcado como "PG" e 'SITUAÇÃO' como "PENDENTE".
# É recomendado que a planilha seja preenchida sempre com Uppercase.
# Após a execução do script, altere a situação na planilha para que o sistema não tente marcar o exame novamente.
# A planilha deverá ser compartilhada por link e estar disponível no OneDrive do usuário.
# Para que a planilha fosse acessada, foi utilizada a API REST do OneDrive.
# # https://docs.microsoft.com/pt-br/onedrive/developer/rest-api/api/shares_get?view=odsp-graph-online
# É necessária uma assinatura do Office 365. Se for uma versão business, é necessário o contexto de autenticação.
# Verificar documentação API REST do SharePoint.
# A notificação é feita via WhatsApp utilizando API do ZapFácil (requer assinatura / verificar documentação ZapFácil).
# ####################################################################################################################

import base64
import os
import sys
import credentials
import requests
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait
from subprocess import CREATE_NO_WINDOW
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep, strftime


def create_onedrive_direct_download(onedrive_link_file):
    # Acessa a API REST do one drive
    data_bytes64 = base64.b64encode(bytes(onedrive_link_file, 'utf-8'))
    data_bytes64_string = data_bytes64.decode('utf-8') \
        .replace('/', '_').replace('+', '-').rstrip("=")
    result_url = f"https://api.onedrive.com/v1.0/shares/u!{data_bytes64_string}/root/content"
    return result_url


# Acessando a Planilha

# <Faz o download da planilha compartilhada através do link, na memória.>

# Compartilhe a planilha com link e defina a variável onedrive_link = "link"
onedrive_link: str = credentials.MyCredentials.onedrive_link
direct_link = create_onedrive_direct_download(onedrive_link)
# noinspection PyArgumentList
data = pd.read_excel(direct_link)
data_df = data.loc[(data['TAXA'] == "PG") & (data['SITUAÇÃO'] == "PENDENTE"), ["NOME DO ALUNO", "TELEFONE", "CPF"]]

# Variáveis de objetos do Site
# Verifique a unidade de agendamento em sua cidade
title = 'Sistema de Agendamento de Exames'
textbox_cpf = '/html/body/center/form/table/tbody/tr[2]/td[2]/input'
listbox_unidade = '//*[@id="localExame"]'
unidade = credentials.MyCredentials.unidade  # Alterar para a unidade de agendamento da sua cidade
listbox_turno = '//*[@id="turno"]'
confirma_dados = '/html/body/center/form/input[2]'
confirma_exame = '/html/body/center/form/input[3]'
agendamento_disponivel = textbox_cpf
date = strftime("%d %b %Y")


class Aluno:
    @staticmethod
    def create_aluno():
        # <Cria uma lista de alunos>
        lista_alunos = []
        for i, nome in enumerate(data_df['NOME DO ALUNO']):
            telefone_df: str = data_df['TELEFONE'].iloc[i]
            # O sistema lê a coluna CPF como um INT e elimina os zeros iniciais do CPF
            # necessário formatar para preencher os zeros eliminados
            cpf_df: str = '%011d' % int(data_df['CPF'].iloc[i])
            telefone: str = telefone_df
            dados_pessoais = nome, telefone, cpf_df
            lista_alunos.append(dados_pessoais)
        return lista_alunos


class Marcador:
    @staticmethod
    # Webdriver
    def load_webdriver():
        # Em caso de erro, verificar se a versão do Chrome e do WebDriver são compatíveis.
        # Caso o código seja compilado para um executável,
        # o webdriver.exe e o Signa - Prodemge 1.4.0.0.crx deverão ser copiados para a pasta raiz do programa.

        # Necessário já ter instalado o Signa.exe na máquina, disponível em:
        # https://wwws.prodemge.gov.br/images/Aplicativos/Signa-2.2.00-Prodemge.exe

        # <Faz a instalação da extensão do plugin Signa - Prodemge 1.4.0.0>
        executable_path = "./"
        os.environ["webdriver.chrome.driver"] = executable_path
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_extension('Signa - Prodemge 1.4.0.0.crx')

        # <Remove Prompt do Driver>
        chrome_service = ChromeService("./chromedriver")
        chrome_service.creationflags = CREATE_NO_WINDOW
        # noinspection PyGlobalUndefined
        global driver
        driver = webdriver.Chrome(options=chrome_options, service=chrome_service)

    @staticmethod
    def close_driver():
        driver.quit()
        sys.exit()

    @staticmethod
    def login():
        # início
        driver.get(
            # Página de login do Detrannet
            "https://empresas.detran.mg.gov.br/sdaf/paginas/ssc/login.asp"
        )

        # noinspection PyBroadException
        try:
            # <Aguarda a validação do certificado digital.>
            WebDriverWait(driver, timeout=120).until(ec.title_is(title))
        except Exception:
            Marcador.close_driver()

    @staticmethod
    def pagina_inicial():
        driver.get(
            # Página de marcação de legislação
            "https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0501.asp"
        )

        # <Verifica se a marcação está disponível.>
        if driver.find_element(By.XPATH, agendamento_disponivel).is_displayed():
            pass
        else:
            # <Aguarda por 10 minutos o sistema abrir para marcação.>
            WebDriverWait(driver, timeout=600).until(ec.visibility_of(
                driver.find_element(By.XPATH, agendamento_disponivel)))
        sleep(1)

    # noinspection PyGlobalUndefined
    @staticmethod
    def definir_turno():
        global turno
        turno = "Manhã"

    # noinspection PyGlobalUndefined
    @staticmethod
    def mudar_turno():
        global turno
        turno = "Tarde"

    @staticmethod
    def trata_erro():
        # noinspection PyBroadException
        try:
            Marcador.pagina_inicial()
            Marcador.mudar_turno()
            Marcador.inserir_dados(col[2])
            Marcador.send_message(nome=col[0], telefone=col[1], cpf=col[2])
            Senhas.imprime_senhas()
        except Exception:
            # noinspection PyGlobalUndefined
            global my_path
            my_path = f'./logFile/{date}/'  # Define o caminho que a pasta de log será criada
            Senhas.make_dir()
            Senhas.print_screen(path=my_path, nome=col[0], cpf=col[2])

    @staticmethod
    def inserir_dados(cpf):
        # <Insere os dados na página de marcação.>
        # Inserir cpf
        driver.find_element(By.XPATH, textbox_cpf).send_keys(cpf, Keys.TAB)
        # Inserir unidade
        driver.find_element(By.XPATH, listbox_unidade).send_keys(unidade)
        # inserir turno
        driver.find_element(By.XPATH, listbox_turno).send_keys(turno)
        # <Aguarda confirmação dos dados.>
        WebDriverWait(driver, timeout=120).until(ec.invisibility_of_element(
            driver.find_element(By.XPATH, confirma_dados)))
        # <Aguarda confirmação do exame.>
        if driver.find_element(By.XPATH, confirma_exame).is_enabled():
            WebDriverWait(driver, timeout=120).until(ec.invisibility_of_element(
                driver.find_element(By.XPATH, confirma_exame)))
        else:
            return Exception

    @staticmethod
    def send_message(**aluno):
        requests.post(credentials.MyCredentials.zap_facil_id,
                      json=aluno)


class Senhas:
    @staticmethod
    def make_dir():
        # Cria um diretório no caminho especificado
        if not os.path.isdir(my_path):
            os.makedirs(my_path)

    @staticmethod
    def print_screen(path, nome, cpf):
        # print screen
        driver.find_element(
            By.CSS_SELECTOR, 'body').screenshot(f'{path}/{nome}_{cpf}.png')
        sleep(1)

    @staticmethod
    def imprime_senhas():
        # noinspection PyGlobalUndefined
        global my_path
        my_path = f'./Senhas/{date}/'  # Define o caminho que a pasta de log será criada
        Senhas.make_dir()
        driver.get("https://empresas.detran.mg.gov.br/sdaf/paginas/sdaf0327.asp")
        sleep(1)
        Senhas.print_screen(path=my_path, nome=col[0], cpf=col[2])


if __name__ == "__main__":
    # Script
    Marcador.load_webdriver()
    Marcador.login()
    for col in Aluno.create_aluno():
        '''
        col[0] == nome
        col[1] == telefone
        col[2] == cpf
        '''
        Marcador.definir_turno()
        # noinspection PyBroadException
        try:
            Marcador.pagina_inicial()
            Marcador.inserir_dados(col[2])
            Marcador.send_message(nome=col[0], telefone=col[1], cpf=col[2])
            Senhas.imprime_senhas()
        except Exception:
            # noinspection PyBroadException
            try:
                Marcador.trata_erro()
            except Exception:
                Marcador.close_driver()
    Marcador.close_driver()
sys.exit()
