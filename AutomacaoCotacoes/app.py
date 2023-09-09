from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import smtplib
import email.message
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import pandas as pd
from time import sleep

class Scrapy:
    def __init__(self):
        self.navegadorConfigure()
        self.cotarDolar()
        self.cotarEuro()
        self.cotarOuro()
        self.montarDf()
        self.enviarEmail()

    def navegadorConfigure(self):
        self.options = webdriver.ChromeOptions()
        self.options.add_argument("--headless=new")
        self.service = Service(ChromeDriverManager().install())
        self.browser = webdriver.Chrome(service=self.service, options=self.options)
    
    def cotarDolar(self):
        self.browser.get("https://www.google.com/")
        self.browser.find_element(By.XPATH,
                      '//*[@id="APjFqb"]').send_keys('Cotação dolar' + Keys.ENTER)
        self.cotacaoDolar = self.browser.find_element(By.XPATH, 
                                    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
        print("\033[32;1mCotacao do Dolar realizada\033[0;0m")

    def cotarEuro(self):
        self.CotarDolar = self.browser.find_element(By.CSS_SELECTOR,
                    '#knowledge-currency__updatable-data-column > div.ePzRBb > div > div.vLqKYe.egcvbb.q0WxUd > div > select').click()
        sleep(0.2)
        self.browser.find_element(By.CSS_SELECTOR,
                    '#knowledge-currency__updatable-data-column > div.ePzRBb > div > div.vLqKYe.egcvbb.q0WxUd > div > select').send_keys("eur" + Keys.ENTER)
        sleep(0.2)
        self.cotacaoEuro = self.browser.find_element(By.XPATH,
                    '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute("data-value")
        print("\033[32;1mCotacao do Euro realizada\033[0;0m")

    def cotarOuro(self):
        self.browser.get('https://www.melhorcambio.com/ouro-hoje')
        self.cotacaoOuro = self.browser.find_element(By.XPATH,'//*[@id="comercial"]').get_attribute("value")
        self.cotacaoOuro = self.cotacaoOuro.replace(',', '.')
        print("\033[32;1mCotacao do Ouro realizada\033[0;0m")

    def montarDf(self):
        self.df = pd.read_excel("Produtos.xlsx")
        
        # Realizando calculos para atualizar o preço final
        self.df.loc[self.df['Moeda'] == "Dólar", "Cotação"] = float(self.cotacaoDolar)
        self.df.loc[self.df['Moeda'] == "Euro", "Cotação"] = float(self.cotacaoEuro)
        self.df.loc[self.df['Moeda'] == "Ouro", "Cotação"] = float(self.cotacaoOuro)

        # Atualizando preços (Valor original *  cotação)
        self.df['Preço de Compra'] = self.df['Preço Original'] * self.df['Cotação']

        # Preço final (Preço de compra * Margem)
        self.df['Preço de Venda'] = self.df['Preço de Compra'] * self.df['Margem']

        # transformando em R$
        self.df['Preço de Venda'] = self.df['Preço de Venda'].map("R${:.2f}".format)

        # Salvando nova base
        self.df.to_excel("ProdutosNovo.xlsx", index=False)

    def enviarEmail(self):
        self.corpoEmail = '''
        <p>Bom dia Ruan</p>
        <p>Segue em anexo a Planilha com as cotações atualizadas</p>
        <p>A disposição</p>
        <h5>Esta mensagem foi encaminhada por um robô, por favor não responda</h5>
        '''
        
        # Abrindo o arquivo em modo de leitura e binary
        self.caminhoArquivo = 'C:\\Users\\mruan\\OneDrive\\Área de Trabalho\\AutomacaoCotacoes\\ProdutosNovo.xlsx'
        self.anexo = open(self.caminhoArquivo, 'rb')

        # Configurando padrões da lib
        self.msg = MIMEMultipart()
        self.msg['Subject'] = 'Planilha de Compras'
        self.msg['From'] = 'mruan309@gmail.com'
        self.msg['to'] = 'garagem.itech@gmail.com'
        self.password = 'aarojhnbxckuhwyf'
        self.msg.add_header('Content-Type', 'text/html')
        self.msg.attach(email.mime.text.MIMEText(self.corpoEmail, 'html'))

        #Anexando arquivo
        self.produtos = 'ProdutosNovo.xlsx'
        with open(self.produtos, 'rb') as xlsx_file:
            self.xlsx_attachment = MIMEApplication(xlsx_file.read(), Name=self.produtos)
            self.xlsx_attachment['Content-Disposition'] = f'attachment; filename={self.produtos}'
            self.msg.attach(self.xlsx_attachment)

        #Enviado e-mail
        self.s = smtplib.SMTP('smtp.gmail.com: 587')
        self.s.starttls()

        # Login Credentials for sending email
        self.s.login(self.msg['From'], self.password)
        self.s.sendmail(self.msg["From"], [self.msg['to']], self.msg.as_string().encode('utf-8'))
        self.s.quit()
        print("\033[32;1mE-mail enviado com sucesso.\033[0;0m")

Scrapy()