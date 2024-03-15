
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
#from PrettyColorPrinter import add_printer
from seleniumbase import Driver
import pandas as pd
import datetime
import urllib
import time
import os

#tempo inicial
tempo_inicio = time.time()

#Endereço pasta arquivos
pasta_reembolso = 'C:\\Users\\inec\\OneDrive - Instituto Nordeste Cidadania\\AGROAMIGO\RELATÓRIOS_2024\\_REEMBOLSO'
contatos_df = pd.read_excel("C:\\send_mail\\Contato_Agentes.xlsx") 
total_agentes = contatos_df['Agente'].count()

data_h = datetime.datetime.now()
#add_printer(1)
navegador = Driver(uc=True)
# Abrir o Whatsapp com o chrome
navegador.get("https://web.whatsapp.com/")
#esperar conectar
time.sleep(10)
while True:
    try:
        # Tente encontrar o elemento usando o XPath
        elemento = navegador.find_element(By.XPATH,'//*[@id="app"]/div/div/div[3]/div[1]/div/div/div[2]/div/canvas')
        print('Aguardando Conectar...')
    except NoSuchElementException:
        # Se o elemento não estiver presente, saia do loop
        print('Conectado!!!')
        break
    # Por exemplo, você pode esperar ou fazer outra coisa
    time.sleep(1)
 
    #____________________________WHATSAPP__________________________________________________________

for i, mensagem in enumerate(contatos_df['Agente']):
            unidade = contatos_df.loc[i, "Unidade"]
            nome_completo = contatos_df.loc[i, "Agente"]
            nome_primeiro = nome_completo.split()
            nome_primeiro = nome_primeiro[0]
            nome_primeiro = nome_primeiro.capitalize()
            numero = contatos_df.loc[i, "Número"]
            msg = contatos_df.loc[i, "Menssagem"]
            link_drive = contatos_df.loc[i, "Link_OneDrive"]
            texto = urllib.parse.quote(f"Oi {nome_primeiro} {msg} {link_drive}")
            arquivo_envio = f'{pasta_reembolso}\\{unidade}\\{nome_completo}.pdf'
            restante = total_agentes - i
            time.sleep(2)
            link = f"https://web.whatsapp.com/send?phone=55{numero}&text={texto}"
            navegador.get(link)
            agora = datetime.datetime.now()
            minuto_inicio = agora.minute
            while True:
                try:
                    # Verifique se o elemento está presente usando um tempo limite
                    bt_enviar_msg = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@class="_2xy_p _3XKXx"]/button'))
                    )
                    erro = WebDriverWait(navegador, 10).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]'))
                    )
                    print(f'Enviando msg {nome_completo}...')
                    # Se o elemento estiver presente, saia do loop
                    break
                except:
                    erro =WebDriverWait(navegador, 240).until(
                        EC.presence_of_element_located((By.XPATH, '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]'))
                    )
                    # Se o elemento não estiver presente, continue aguardando
                    if erro != '':
                        print(erro.text)
                    print(f'Aguardando conversa {nome_completo}...')
                    agora = datetime.datetime.now()
                    minuto_atual = agora.minute
                    minutos = minuto_atual - minuto_inicio
                    if minutos > 2:
                         break
                    continue
            try:    
                time.sleep(2)
                bt2_msg = navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span')
                bt2_msg.click()
                #if os.path.exists(arquivo_envio)==True:
                navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/div/span').click()
                time.sleep(2)
                input_arquivo = navegador.find_element(By.XPATH,'//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/ul/div/div[1]/li/div/input') 
                input_arquivo.send_keys(arquivo_envio)
                time.sleep(2)
                input_legenda = navegador.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[1]/div[3]/div/div/div[1]/div[1]/p')
                input_legenda.send_keys("Reembolso PDF")
                bt_enviar_arquivo = navegador.find_element(By.XPATH,'//*[@id="app"]/div/div[2]/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div/span')
                bt_enviar_arquivo.click() #Clicar em Enviar arquivo  
                time.sleep(2)
                #Clicar em Enviar Menssagem
                time.sleep(2)    
                data_h = datetime.datetime.now()
                tempo_atual = time.time() - tempo_inicio         
            except:
                data_h = datetime.datetime.now()
                tempo_atual = time.time() - tempo_inicio 
                continue
            duracao = f"{int(tempo_atual // 3600):02d}:{int((tempo_atual % 3600) // 60):02d}:{int(tempo_atual % 60):02d}"

            print(f'Tempo: processo: {i +1 } faltam: {restante} Agente: {nome_completo}')
            
print(f'Processo Finalizado. Tempo: processo: {i +1 } faltam: {restante} Agente: {nome_completo}')

