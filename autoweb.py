from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from time import sleep


import pandas as pd

''''
banana = webdriver.Chrome(service=servico)
banana.get("https://www.google.com/")
sleep(10)
'''

chrome_options = Options()
#chrome_options.add_argument('--headless')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_argument('--log-level=3')



servico = Service(ChromeDriverManager().install())
banana=webdriver.Chrome(service=servico,options=chrome_options)
banana.get("https://www.google.com/")
banana.find_element('xpath','/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação do dolar')
banana.find_element('xpath','/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
valor_dolar =  banana.find_element('xpath','//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value').replace(',','.')

banana.get("https://www.google.com/")
banana.find_element('xpath','/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys('cotação do euro')
banana.find_element('xpath','/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/input').send_keys(Keys.ENTER)
valor_euro =  banana.find_element('xpath','//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]').get_attribute('data-value').replace(',','.')


banana.get('https://www.melhorcambio.com/ouro-hoje')#//*[@id="comercial"]
valor_ouro = banana.find_element('xpath','//*[@id="comercial"]').get_attribute('value').replace(',','.')


#banana.quit()

tabela = pd.read_excel(r'C:\\Users\\tf938\Downloads\\Produtos.xlsx')

print(tabela)

tabela.loc[tabela['Moeda'] == 'Dólar','Cotação'] = float(valor_dolar) 
tabela.loc[tabela['Moeda'] == 'Euro','Cotação'] = float(valor_euro) 
tabela.loc[tabela['Moeda'] == 'Ouro','Cotação'] = float(valor_ouro) 

tabela['Preço de Compra'] = tabela['Cotação'] * tabela['Preço Original']

tabela['Preço de Venda'] = tabela['Preço de Compra'] * tabela['Margem']
#tabela['Preço de Venda'] = tabela['Preço de Venda'].map("{:,.2f})". format) #FORMATAÇÃO PRA DEIXAR BONITO 


print(tabela)

tabela.to_excel("Produtos Atualizados.xlsx",index=False)
print('Banco de Dados Atuaizada e extraída')





