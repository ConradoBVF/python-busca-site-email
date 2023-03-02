import requests
from bs4 import BeautifulSoup
import re
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
import pandas as pd
import yagmail
from datetime import datetime
import math

headers = headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
}

def busca_sites(atividade, regiao):
    """Essa função busca 11 páginas do Google os sites que aparecem como os primeiros resultados de
    pesquisa de uma palavra ou termo chave que o usuário solicita."""
    ser = Service('C:/chromedriver.exe')
    options = webdriver.ChromeOptions()
    options.add_argument('headless')
    driver = webdriver.Chrome(service=ser, options=options)
    google = 'https://www.google.com/search?q='
    driver.get(google + atividade + ' ' + regiao)
    dicionario_pesquisa = {'Site': []}
    try:
        for i in range(9):
            texto = driver.find_elements(by=By.PARTIAL_LINK_TEXT, value='http')
            for elemento in texto:
                site = elemento.text
                if 'http' in site:
                    site = site[site.find('http'):]
                    site2 = site[site.find('//') + len('//'):]
                    site2 = site2[:site2.find('/')]
                    site = site[:site.find(site2) + len(site2)]
                    if ' ' in site:
                        site = site[:site.find(' ')]
                    if site[-2:] == '.b':
                        site = site + 'r'
                    if site[-3:] == '.co':
                        site = site + 'm'
                    dicionario_pesquisa['Site'].append(site)
            time.sleep(0.5)
            driver.find_element(by=By.XPATH, value='//*[@id="pnnext"]/span[1]').click() # Botão de próxima página
    except:
        pass
    dicionario_pesquisa['Site'] = sorted(set(dicionario_pesquisa['Site']))
    print(atividade, ' - ', len(dicionario_pesquisa['Site']))
    print(dicionario_pesquisa['Site'])
    return dicionario_pesquisa

def find_emails(url):
    response = requests.get(url)
    html = response.text
    emails = re.findall(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', html)
    return emails

def busca_email(site):
    # Make a request to the website
    response = requests.get(site, headers=headers, allow_redirects=False, timeout=5)

    # Use BeautifulSoup to parse the HTML content
    soup = BeautifulSoup(response.content, "html.parser")

    # Find all the text on the page
    text = soup.get_text()

    # Use a regular expression to search for email addresses
    email_addresses = re.findall(r'[\w\.-]+@[\w\.-]+', text)
    print(email_addresses)
    emails = set(email_addresses)
    print(emails)
    regex = re.compile(r'([A-Za-z0-9]+[.-_])*[A-Za-z0-9]+@[A-Za-z0-9-]+(\.[A-Z|a-z]{2,})+')
    for email in emails:
        if 'com' in email or 'ind' in email or 'net' in email:
            teste = 0
            if re.fullmatch(regex, email):
                prefixo = email[:email.find('@')]
                for char in prefixo:
                    if char.isnumeric():
                        teste += 1
                if teste > 3:
                    pass
                else:
                    return email

def criar_lista(urls):
    lista_geral = []
    for site in urls['Site']:
        try:
            print(site)
            email = busca_email(site)
            lista_geral.append([site, email])
        except:
            pass  
    return lista_geral

def transformar_xlsx(lista_site_emails):
    df = pd.DataFrame(lista_site_emails)
    df.to_excel('lista-final-emails.xlsx', index=False, header=False)

def abrir_lista(coluna1, coluna2):
    # Buscar cada um dos sites e emails que o bot encontrou
        try:
            math.isnan(float(coluna2))
            pass
        except:
            site = coluna1
            email = coluna2
            return site, email

def enviar_email(site, destinatario):
    try:
        usuario = yagmail.SMTP(user='geral.prospector@gmail.com', password='******************')
        anexo = './CV_Conrado_Ferreira.pdf'
        if datetime.now().hour >= 12:
            cumprimento = 'Boa tarde'
        else:
            cumprimento = 'Bom dia'
        corpo_email = [
            f'{cumprimento}, tudo bem?',
            '\n'
            'Sou programador em Python e estou buscando uma vaga no mercado.\n'
            f'Encontrei o site {site} através de uma busca automática realizada pelo meu próprio bot.\n',
            'Segue link no Github caso queira conhecer melhor: \n'
            'Também tomei a liberdade de anexar meu currículo neste e-mail para sua visualização.\n'
            'Fico à disposição para quaisquer oportunidades.\n'
            '\n'
            'Atenciosamente,\n'
            '\n'
            f'Conrado Botelho'
        ]
        usuario.send(to=destinatario, subject='Planilhas', contents=corpo_email, attachments=anexo)
    except KeyError:
        print(KeyError)

# atividade = input('Digite o termo de busca: ')
# regiao = input('Digite a região: ')
# urls = busca_sites(atividade, regiao)
# lista_site_emails = criar_lista(urls)
# print(lista_site_emails)
# transformar_xlsx(lista_site_emails)
# enviar_email('www.blablabla.com.br', 'conrado@grupohidrica.com.br')
lista = 'lista-final-emails.xlsx'
df = pd.read_excel(lista)

for index, row in df.iterrows():
    try:
        site, email = abrir_lista(row[0], row[1])
        print(site, email)
    except:
        pass
