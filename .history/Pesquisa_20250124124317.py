from dotenv import load_dotenv
from openpyxl import load_workbook
from twocaptcha import TwoCaptcha
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import pyautogui
import requests
import shutil
import psutil
import pdfplumber
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import undetected_chromedriver as uc
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
import os
import re
import pickle
import base64
from PIL import Image

# Cria uma pasta de saída com base no nome da planilha
pasta_baixou = r"C:\Users\stefany\Downloads"

API_KEY = "api_key"

# Cria uma pasta de saída com base no nome da planilha
def criar_pasta_saida(caminho_planilha, pasta_downloads):
    nome_arquivo = os.path.splitext(os.path.basename(caminho_planilha))[0]
    pasta_saida = os.path.join(pasta_downloads, nome_arquivo)

    # Verifica se a pasta existe
    if not os.path.exists(pasta_saida):
        os.makedirs(pasta_saida)
        print(f"Pasta criada: {pasta_saida}")
    else:
        print(f"A pasta já existe: {pasta_saida}")

    return pasta_saida

def caminho_paraBoleto(pasta_baixou, pasta_saida, placa_atual):
    # Verifica o nome de todos os processos abertos
    for process in psutil.process_iter(['pid', 'name']):
        # Acha o Acrobat na lista de processos
        if 'Acrobat' in process.info['name']:
            process.terminate()  # Fecha o Adobe
    arquivos = [f for f in os.listdir(pasta_baixou) if f.endswith('.pdf')]
    arquivo = max(arquivos, key=lambda f: os.path.getmtime(os.path.join(pasta_baixou, f)))
    if arquivo in arquivos:
        caminho_arquivo = os.path.join(pasta_baixou, arquivo)
        novo_nome = os.path.join(pasta_saida, f"{placa_atual}.pdf")
        # Renomeia o arquivo com a placa
        shutil.move(caminho_arquivo, novo_nome)
        print(f"Boleto salvo como {novo_nome}")
        return novo_nome
    


