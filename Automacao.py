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
import unicodedata

#################### conexao mongodb
client = MongoClient('mongodb://localhost:27017/')
bd = client['logs_db']
colecao = bd['install_logs']

#################### função log 
def logs(evento, status, detalhes=""):
    log = {
        "timestamp": datetime.now(),
        "evento": evento,
        "status": status,
        "detalhes": detalhes
    }
    colecao.insert_one(log)
    print(f"Log registrado: {evento}, Status: {status}")

#################### install python
def instalar_python(caminho_instalador):
    try:
        subprocess.run(['runas', '/user:Administrator', caminho_instalador, '/quiet', 'InstallAllUsers=1', 'PrependPath=1'], check=True)
        logs("Execução da Instalação", "sucesso", "Instalação iniciada.")
        print("Instalação do Python concluída!")
    except subprocess.CalledProcessError as e:
        logs("Execução da Instalação", "erro", str(e))
        print(f"Erro na execução da instalação: {e}")

#################### validador da versao
def valida_versao():
    try:
        resultado = subprocess.run([sys.executable, "--version"], capture_output=True, text=True)
        versao_python = resultado.stdout.strip()
        logs("Verificação da Versão", "sucesso", f"Versão instalada: {versao_python}")
        print(f"Versão instalada do Python: {versao_python}")
        
        if "Python 3.12.7" in versao_python:
            print("Instalação validada")
        else:
            print("A versão instalada não é a 3.12.7")
    except FileNotFoundError:
        logs("Verificação da Versão", "erro", "Python não encontrado.")
        print("Python não foi encontrado.")

#################### detecta o chrome driver
chromedriver_autoinstaller.install()

#################### configuraçoes para abrir na janela| o headless é para executar sem abrir janeja
opcoes_chrome = Options()
opcoes_chrome.add_argument("--start-maximized")
opcoes_chrome.add_argument("--headless")
driver = webdriver.Chrome(options=opcoes_chrome)

#################### acessar o google/python.org e fazer o download
driver.get("https://www.google.com")
driver.find_element(By.NAME, "q").send_keys("download python", Keys.RETURN)
time.sleep(5)
driver.find_elements(By.XPATH, '//a[contains(@href, "www.python.org")]')[0].click()
time.sleep(5)
driver.find_element(By.XPATH, '//*[@id="downloads"]/a').click()
time.sleep(5)
driver.find_element(By.XPATH, '//a[contains(text(), "Python 3.12.7")]').click()
time.sleep(10)

#################### obtem o link e usa o request para baixar
link_download = driver.find_element(By.XPATH, '//a[contains(text(), "Windows installer (64-bit)")]').get_attribute("href")
resposta = requests.get(link_download)
diretorio_download = os.path.join(os.environ["USERPROFILE"], "Downloads")
caminho_arquivo = os.path.join(diretorio_download, "python_installer.exe")

with open(caminho_arquivo, "wb") as arquivo:
    arquivo.write(resposta.content)

logs("Baixa python", "sucesso", f"Arquivo salvo em: {caminho_arquivo}")
print(f"Download iniciado! Arquivo salvo em: {caminho_arquivo}")

#################### tempo para garantir que sera concluido o download
time.sleep(5)

#################### chama a funçao de instalacao
instalar_python(caminho_arquivo)

#################### funçao validador
valida_versao()

#################### finaliza o navegador
driver.quit()

#################### funçao para webscraping
def extrair_dados():
    soap = BeautifulSoup(requests.get('https://books.toscrape.com').text, 'html.parser')  
    dados = []
    
#################### manipular a pagina html para poder manipular
    for livro in soap.find_all('article', class_='product_pod'):
        titulo = livro.find('h3').find('a')['title']
#################### formatando o dinheiro
        preco = livro.find('p', class_='price_color').text
        preco = unicodedata.normalize('NFKD', preco).strip().replace('AÌ', '').replace('£', 'R$')
        avaliacao = livro.find('p', class_='star-rating')['class'][1]
        disponibilidade = livro.find('p', class_='instock availability').text.strip()
#################### mudando o in stock para         
        if "In stock" in disponibilidade:
            quantidade = disponibilidade.split()[-1]
            disponibilidade = f"Em estoque: {quantidade}"

        dados.append([titulo, preco, avaliacao, disponibilidade])
#################### chama a funçao salvar com a lista dos dados    
    salvar_csv(dados)

#################### funcao de salvamento
def salvar_csv(dados):
    nome = 'livros.csv'
    
    with open(nome, 'w', newline='', encoding='utf-8') as arquivo:
        w = csv.writer(arquivo)
        w.writerow(['Título', 'Preço', 'Avaliação', 'Disponibilidade'])
        w.writerows(dados)
    
    print(f'Dados salvos em: {nome}')

#################### funcao etl (filtros, agregações)
def manipular_dados():
    df = pd.read_csv('livros.csv')
#################### filtrando apenas os livros com nota 5    
    df_modificado = df[df['Avaliação'] == 'Five']
#################### estou limpando os dados da coluna preco
    df_modificado['Preço'] = (df_modificado['Preço'].str.replace('AÌ‚', '', regex=False).str.strip())
    
    contagem_disponibilidade = df['Disponibilidade'].value_counts()
    
    print(f'Filtrados por avaliação "Five":\n{df_modificado}')
    print(f'Contagem de livros por disponiveis:\n{contagem_disponibilidade}')
    
    criar_pdf(df_modificado, contagem_disponibilidade)

#################### funcao para gerar o PDF
def criar_pdf(dados_filtrados, contagem_disponibilidade):
#################### caminho pdf
    nome = 'pdf_livros.pdf'
#################### configuracoes do pdf    
    c = canvas.Canvas(nome, pagesize=letter)
    c.setFont("Helvetica", 10)    
#################### titulo
    c.setFont("Helvetica-Bold", 14)
    c.drawString(100, 750, "Relatório de Livros - Books to Scrape")
    c.setFont("Helvetica", 10)
    c.drawString(100, 735, f"Data: {time.strftime('%Y-%m-%d')}")
    
#################### informacao dos livros
    y_position = 700
    c.setFont("Helvetica-Bold", 12)
    c.drawString(100, y_position, "Livros com avaliação 'Five':")
    y_position -= 15
    
#################### faz um for para pegar todos os livros com avaliacao 5
    for i, linha in dados_filtrados.iterrows():
        c.setFont("Helvetica", 10)
        c.drawString(100, y_position, f"Título: {linha['Título']}, Preço: {linha['Preço']}, Avaliação: {linha['Avaliação']}")
        y_position -= 15
    
#################### trava para nao passar do limite da pagina
    if y_position < 100:
        c.showPage()
        y_position = 750
        c.setFont("Helvetica-Bold", 12)
        c.drawString(100, y_position, "Contagem de livros por disponibilidade:")
        y_position -= 15
    
#################### por ultimo a contagem de livros em estoque
    c.setFont("Helvetica", 10)
    for index, count in contagem_disponibilidade.items():
        c.drawString(100, y_position, f"{index}: {count}")
        y_position -= 15
    
#################### salvar
    c.save()
    print(f'PDF gerado: {nome}')

#################### funcao main
def main():
    print("Iniciando a extração de dados...")
    extrair_dados()
    print("Manipulando dados e gerando relatório...")
    manipular_dados()

#################### chama a funcao main
if __name__ == "__main__":
    main()
