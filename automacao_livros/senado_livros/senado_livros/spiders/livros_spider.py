import logging
from selenium.webdriver.chrome.options import Options
import scrapy
import pandas as pd
from scrapy import signals
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementNotInteractableException
import time
import os
import random

class LivrosSpiderSpider(scrapy.Spider):
    name = "livros_spider"
    allowed_domains = ["md2127:8901"]
    start_urls = ["http://md2127:8901/login"]
    prateleiras_url = "http://md2127:8901/shelves"

    @classmethod
    def from_crawler(cls, crawler, *args, **kwargs):
        spider = super(LivrosSpiderSpider, cls).from_crawler(crawler, *args, **kwargs)
        crawler.signals.connect(spider.spider_closed, signal=signals.spider_closed)
        return spider

    def __init__(self):
        chrome_options = Options()
        chrome_options.add_argument("--start-maximized")
        self.driver = webdriver.Chrome(options=chrome_options) # Certifique-se de que o ChromeDriver está instalado e no PATH

    def start_requests(self):
        yield scrapy.Request(url=self.start_urls[0], callback=self.login)

    def login(self, response):
        self.driver.get(response.url)
                
        # Fazer login
        self.driver.find_element(By.ID, 'email').send_keys('*************')
        self.driver.find_element(By.ID, 'password').send_keys('*********')
        self.driver.find_element(By.CSS_SELECTOR, 'button.button').click()

        # Substitua o caminho da pasta correto
        dados = pd.read_csv(r'D:\USERS\05056434132\Documents\Demandas\automacao_livros\senado_livros\senado_livros\spiders\livros.csv', delimiter=';')

        # Iterar sobre as linhas do dataset
        for index, row in dados.iterrows():
            if row['STATUS'] != 'publicado':  # Verificar se o status não é "publicado"
                self.processar_acervo(row)
                self.criar_nova_pagina(row)
                # Atualizar o status para "publicado"
                dados.at[index, 'STATUS'] = 'publicado'

        # Salvar as alterações no arquivo CSV com ponto e vírgula como separador
        dados.to_csv(r'D:\USERS\05056434132\Documents\Demandas\automacao_livros\senado_livros\senado_livros\spiders\livros.csv', index=False, sep=';')

    def processar_acervo(self, row):
        # Navegar até a prateleira correta
        prateleira = row['PRATELEIRAS']
        print(prateleira)
        
        if prateleira == 'apresentacoes':
            url = 'http://md2127:8901/shelves/apresentacoes/create-book'
        elif prateleira == 'livros':
            url = 'http://md2127:8901/shelves/livros/create-book'
        elif prateleira == 'artigos':
            url = 'http://md2127:8901/shelves/tese-artigos-e-pesquisas/create-book'
        else:
            self.log(f'Tipo de item desconhecido: {prateleira}', level=logging.ERROR)
            return

        # Navegar para a URL correta
        self.driver.get(url)
        time.sleep(3) 

        try:
            wait = WebDriverWait(self.driver, 10)
            nome_input = wait.until(EC.presence_of_element_located((By.NAME, 'name')))
            nome_input.send_keys(row['NOME'])
            time.sleep(2)  

            # Preencher a descrição
            descricao_texto = (
                f"Autor(es): {row['AUTOR(ES)']}\n"
                f"Editora: {row['EDITORA']}\n"
                f"Nª de páginas: {row['NUMERO DE PAGINAS']}\n"
                f"Edição: {row['EDICAO']}\n"
                f"Idioma: {row['IDIOMA']}\n"
                f"Ano de Publicação: {row['ANO']}"
            )
            descricao_input = self.driver.find_element(By.NAME, 'description')
            descricao_input.send_keys(descricao_texto)
            time.sleep(2) 

            # Adicionar imagem de capa
            self.driver.find_element(By.XPATH, '//button[.//label[text()="Imagem de capa"]]').click()
            time.sleep(2)  

            self.driver.find_element(By.XPATH, '//label[@for="image" and contains(@class, "button outline")]').click()
            time.sleep(2) 

            # Substitua o caminho da pasta correto
            caminho_imagem = os.path.join(
                r'D:\USERS\05056434132\Documents\Demandas\automacao_livros\senado_livros\senado_livros\img',
                f'{prateleira}.png'  
            ).replace('\\', '/')
            self.driver.find_element(By.NAME, 'image').send_keys(caminho_imagem)
            time.sleep(2)  

            # Simular pressionar a tecla "Esc"  para fechar a janela de seleção de arquivo
            nome_input.send_keys(Keys.ESCAPE)
            time.sleep(2)  

            # Adicionar tags
            self.driver.find_element(By.XPATH, '//button[.//label[@for="tag-manager"]]').click()
            time.sleep(2)  

            tags = row['TAGS'].split(',')
            for i, tag in enumerate(tags):
                tag_input_locator = (By.NAME, f'tags[{i}][name]')
                tags_input = wait.until(EC.presence_of_element_located(tag_input_locator))
                wait.until(EC.element_to_be_clickable(tag_input_locator))
                tags_input.send_keys(tag)
                tags_input.send_keys(Keys.ENTER)
                time.sleep(2)  

            self.driver.find_element(By.CSS_SELECTOR, 'button[type="submit"].button').click()
            time.sleep(5) 
            self.driver.find_element(By.CSS_SELECTOR, 'button[type="submit"].button').click()

        except TimeoutException:
            self.log(f"Timeout ao procurar pelo elemento 'name' na prateleira {prateleira}", level=logging.ERROR)
        except NoSuchElementException as e:
            self.log(f"Elemento não encontrado: {e}", level=logging.ERROR)
        except ElementNotInteractableException as e:
            self.log(f"Elemento não interativo: {e}", level=logging.ERROR)

    def criar_nova_pagina(self, row):
        prateleira = row['PRATELEIRAS']
        livro_url = 'http://md2127:8901/books/teste/create-page'
        self.driver.get(livro_url)
        time.sleep(3)
        nome = row['NOME']
        try:
            # Ajustar o título da página de acordo com a prateleira
            if prateleira == 'apresentacoes':
                titulo = f"Apresentação N°: {random.randint(10000000, 99999999)}"
            elif prateleira == 'artigos':
                titulo = f"Artigo N°: {random.randint(10000000, 99999999)}"
            elif prateleira == 'livros':
                titulo = f"Livro N°: {random.randint(10000000, 99999999)}"
            else:
                titulo = f"{prateleira} N°: {random.randint(10000000, 99999999)}"

            # Preencher o título da página
            titulo_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.NAME, 'name'))
            )
            titulo_input.send_keys(titulo)
            time.sleep(2)

            # Clicar no botão de anexos
            self.driver.find_element(By.XPATH, '//button[@title="Anexos"]').click()
            time.sleep(2)

            # Clicar no botão de upload de arquivos
            self.driver.find_element(By.XPATH, '//button[contains(text(), "Upload de Arquivos")]').click()
            time.sleep(2)

            # Clicar para anexar arquivos
            self.driver.find_element(By.XPATH, '//button[@type="button" and contains(@class, "dz-message")]').click()
            time.sleep(4)

            caminho_pdf = os.path.join(
                r'D:\USERS\05056434132\Documents\Demandas\automacao_livros\senado_livros\senado_livros\pdf',
                f'{nome}.pdf'  
            ).replace('\\', '/')
            print(caminho_pdf)

            # Enviar o arquivo PDF
            file_input = self.driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
            file_input.send_keys(caminho_pdf)
            time.sleep(5)  # Aguarde o upload
            
            # Verificar se o PDF foi anexado com sucesso
            anexo_element = WebDriverWait(self.driver, 30).until(
                EC.presence_of_element_located((By.XPATH, '//div[@class="py-s"]/a[contains(@href, "attachments")]'))
            )
            pdf_url = anexo_element.get_attribute('href')

            self.driver.find_element(By.CSS_SELECTOR, 'button[refs="editor-toolbox@toggle"]').click()
            time.sleep(2)
            # Clicar no botão de código fonte
            self.driver.find_element(By.XPATH, '//button[@title="Código fonte"]').click()
            time.sleep(2)

            # Inserir o iframe com o link do PDF
            iframe_html = (
                f'<p id="bkmrk-"><iframe style="width: 100%; height: 950px;" src="{pdf_url}?open=true#toolbar=0&amp;navpanes=0&amp;scrollbar=0"></iframe>'
                f'<a href="{pdf_url}?open=true" target="_blank" rel="noopener">'
                f'<img style="margin-right: 7px;" src="http://md2127:8901/uploads/images/gallery/2023-01/scaled-1680-/popup-link-icon.png" alt="abrir em uma nova guia" width="22px"></a>'
                f'<a href="{pdf_url}" target="_blank" rel="noopener" download="">'
                f'<img src="http://md2127:8901/uploads/images/gallery/2023-01/scaled-1680-/download-file-icon.png" alt="download do acervo em pdf" width="25px"></a></p>'
            )

            code_editor = self.driver.find_element(By.XPATH, '//textarea[@class="tox-textarea"]')
            code_editor.send_keys(iframe_html)
            time.sleep(2)

            # Salvar o código fonte
            self.driver.find_element(By.XPATH, '//button[@title="Salvar"]').click()
            time.sleep(3)

            # Salvar a página
            self.driver.find_element(By.ID, 'save-button').click()
            time.sleep(5)  # Aguarde a página ser salva

        except TimeoutException:
            self.log(f"Timeout ao criar nova página para a prateleira {prateleira}", level=logging.ERROR)
        except NoSuchElementException as e:
            self.log(f"Elemento não encontrado: {e}", level=logging.ERROR)
        except ElementNotInteractableException as e:
            self.log(f"Elemento não interativo: {e}", level=logging.ERROR)

    def spider_closed(self, spider):
        self.driver.quit()