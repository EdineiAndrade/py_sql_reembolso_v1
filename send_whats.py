import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import urllib

navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com/")

pasta_reembolso = 'C:\\Users\\inec\\OneDrive - Instituto Nordeste Cidadania\\AGROAMIGO\RELATÓRIOS_2024\\_REEMBOLSO'
contatos_df = pd.read_excel("C:\\send_mail\\Contato_Agentes.xlsx")
total_agentes = contatos_df['Agente'].count()
erro = "não conectado"
while erro !="":
    try:
        while len(navegador.find_elements_by_id("side")) < 1:
            time.sleep(1)
            erro = ""
    except:
        erro = "não conectado"    
    # já estamos com o login feito no whatsapp web
for i, mensagem in enumerate(contatos_df['Menssagem']):
            pessoa = contatos_df.loc[i, ""]
            numero = contatos_df.loc[i, "Número"]
            unidade = contatos_df.loc[i, "Unidade"]
            arquivo_envio = f'{pasta_reembolso} {unidade}{pessoa}'
            texto = urllib.parse.quote(f"Oi {pessoa}! {mensagem}")
            link = f"https://web.whatsapp.com/send?phone={numero}&text={texto}"
            navegador.get(link)
            while len(navegador.find_elements_by_id("side")) < 1:
                time.sleep(1)
            navegador.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]').send_keys(Keys.ENTER)
            time.sleep(10)
                    