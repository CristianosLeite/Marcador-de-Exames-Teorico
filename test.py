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


def send_message(**json_aluno):
    print(json_aluno)


for aluno in create_aluno():
    send_message(nome=aluno[0], telefone=aluno[1], cpf=aluno[2])
