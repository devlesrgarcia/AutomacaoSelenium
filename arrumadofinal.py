import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import time
import requests
import os
import subprocess
import sys
from pymongo import MongoClient
from datetime import datetime
from bs4 import BeautifulSoup
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import csv

# Conectar ao MongoDB localmente
cliente = MongoClient('mongodb://localhost:27017/')
banco_dados = cliente['logs_db']
colecao = banco_dados['install_logs']

# Função para registrar eventos no MongoDB
def registrar_evento(evento, status, detalhes=""):
    log = {
        "timestamp": datetime.now(),
        "evento": evento,
        "status": status,
        "detalhes": detalhes
    }
    colecao.insert_one(log)
    print(f"Log registrado: {evento}, Status: {status}")

# Função para instalar automaticamente o Python com privilégios elevados
def instalar_python(caminho_instalador):
    try:
        subprocess.run(['runas', '/user:Administrator', caminho_instalador, '/quiet', 'InstallAllUsers=1', 'PrependPath=1'], check=True)
        registrar_evento("Execução da Instalação", "sucesso", "Instalação do Python iniciada com sucesso.")
        print("Instalação do Python concluída com sucesso!")
    except subprocess.CalledProcessError as e:
        registrar_evento("Execução da Instalação", "erro", str(e))
        print(f"Erro na execução da instalação: {e}")

# Função para validar a instalação do Python
def validar_instalacao_python():
    try:
        resultado = subprocess.run([sys.executable, "--version"], capture_output=True, text=True)
        versao_python = resultado.stdout.strip()
        registrar_evento("Verificação da Versão Instalado", "sucesso", f"Versão instalada: {versao_python}")
        print(f"Versão instalada do Python: {versao_python}")
        
        if "Python 3.12.7" in versao_python:
            print("Instalação validada com sucesso!")
        else:
            print("A versão instalada não é a esperada.")
    except FileNotFoundError:
        registrar_evento("Verificação da Versão Instalado", "erro", "Python não encontrado.")
        print("Python não foi encontrado. A instalação pode ter falhado.")

# Instala e configura automaticamente o ChromeDriver
chromedriver_autoinstaller.install()

# Configurações do navegador
opcoes_chrome = Options()
opcoes_chrome.add_argument("--start-maximized")
driver = webdriver.Chrome(options=opcoes_chrome)

# Acessa o Google e realiza a pesquisa para o download do Python
driver.get("https://www.google.com")
driver.find_element(By.NAME, "q").send_keys("download python", Keys.RETURN)
time.sleep(5)
driver.find_elements(By.XPATH, '//a[contains(@href, "www.python.org")]')[0].click()
time.sleep(5)
driver.find_element(By.XPATH, '//*[@id="downloads"]/a').click()
time.sleep(5)
driver.find_element(By.XPATH, '//a[contains(text(), "Python 3.12.7")]').click()
time.sleep(10)

# Obtém o link de download do instalador do Python e faz o download
link_download = driver.find_element(By.XPATH, '//a[contains(text(), "Windows installer (64-bit)")]').get_attribute("href")
resposta = requests.get(link_download)
diretorio_download = os.path.join(os.environ["USERPROFILE"], "Downloads")
caminho_arquivo = os.path.join(diretorio_download, "python_installer.exe")

with open(caminho_arquivo, "wb") as arquivo:
    arquivo.write(resposta.content)

registrar_evento("Baixa do Executável", "sucesso", f"Arquivo salvo em: {caminho_arquivo}")
print(f"Download iniciado! Arquivo salvo em: {caminho_arquivo}")

# Espera para garantir que o download foi concluído
time.sleep(5)

# Instala o Python com permissões elevadas
instalar_python(caminho_arquivo)

# Valida a instalação do Python
validar_instalacao_python()

# Fecha o navegador
driver.quit()

# Função para extrair dados da página web
def extrair_dados_da_web():
    url = 'https://books.toscrape.com'
    resposta = requests.get(url)
    sopa = BeautifulSoup(resposta.text, 'html.parser')
    
    dados_livros = []
    
    # Extraindo informações de cada livro na página
    for livro in sopa.find_all('article', class_='product_pod'):
        titulo = livro.find('h3').find('a')['title']
        preco = livro.find('p', class_='price_color').text
        avaliacao = livro.find('p', class_='star-rating')['class'][1]
        disponibilidade = livro.find('p', class_='instock availability').text.strip()
        
        dados_livros.append([titulo, preco, avaliacao, disponibilidade])
    
    salvar_em_csv(dados_livros)

# Função para salvar os dados extraídos em CSV
def salvar_em_csv(dados):
    caminho_arquivo = 'dados_livros.csv'
    
    with open(caminho_arquivo, 'w', newline='', encoding='utf-8') as arquivo:
        escritor = csv.writer(arquivo)
        escritor.writerow(['Título', 'Preço', 'Avaliação', 'Disponibilidade'])
        escritor.writerows(dados)
    
    print(f'Dados salvos em: {caminho_arquivo}')

# Função para manipular os dados (filtros, agregações)
def manipular_dados():
    df = pd.read_csv('dados_livros.csv')
    
    dados_filtrados = df[df['Avaliação'] == 'Five']
    
    contagem_disponibilidade = df['Disponibilidade'].value_counts()
    
    print(f'Dados filtrados por avaliação "Five":\n{dados_filtrados}')
    print(f'Contagem de livros por disponibilidade:\n{contagem_disponibilidade}')
    
    gerar_relatorio_pdf(dados_filtrados, contagem_disponibilidade)

# Função para gerar o relatório em PDF
def gerar_relatorio_pdf(dados_filtrados, contagem_disponibilidade):
    # Criar o arquivo PDF
    caminho_arquivo = 'relatorio_livros.pdf'
    c = canvas.Canvas(caminho_arquivo, pagesize=letter)
    c.setFont("Helvetica", 10)
    
    # Título do relatório
    c.setFont("Helvetica-Bold", 14)
    c.drawString(100, 750, "Relatório de Livros - Books to Scrape")
    c.setFont("Helvetica", 10)
    c.drawString(100, 735, f"Data: {time.strftime('%Y-%m-%d')}")
    
    # Tabela com os dados filtrados (livros com avaliação 'Five')
    y_position = 700
    c.setFont("Helvetica-Bold", 12)
    c.drawString(100, y_position, "Livros com avaliação 'Five':")
    y_position -= 15
    
    # Exibe os dados dos livros filtrados com avaliação 'Five'
    for i, linha in dados_filtrados.iterrows():
        c.setFont("Helvetica", 10)
        c.drawString(100, y_position, f"Título: {linha['Título']}, Preço: {linha['Preço']}, Avaliação: {linha['Avaliação']}")
        y_position -= 15
    
    # Evita que ultrapasse o limite da página
    if y_position < 100:
        c.showPage()
        y_position = 750
        c.setFont("Helvetica-Bold", 12)
        c.drawString(100, y_position, "Contagem de livros por disponibilidade:")
        y_position -= 15
    
    # Adicionar a contagem de disponibilidade
    c.setFont("Helvetica", 10)
    for index, count in contagem_disponibilidade.items():
        c.drawString(100, y_position, f"{index}: {count}")
        y_position -= 15
    
    # Salvar o PDF
    c.save()
    print(f'Relatório gerado: {caminho_arquivo}')

# Função principal para orquestrar o processo
def principal():
    print("Iniciando a extração de dados...")
    extrair_dados_da_web()
    print("Manipulando dados e gerando relatório...")
    manipular_dados()

# Chama a função principal
if __name__ == "__main__":
    principal()
