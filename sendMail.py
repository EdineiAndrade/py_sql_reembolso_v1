import win32com.client as win32com
import openpyxl
import time
import base64
import os

n_processo = 0

# Função para ler o conteúdo HTML do arquivo
def ler_corpo_html(arquivo):
    with open(arquivo, 'r', encoding='utf-8') as file:
        return file.read()

# Função para enviar e-mails com cópia (CC), anexos e corpo HTML
def enviar_email(nome, destinatario, copia, unidade,assunto,anexo_path):
    outlook = win32com.Dispatch("Outlook.Application")
    message = outlook.CreateItem(0)
    message.To = destinatario
    message.Subject = assunto
    message.Display()
    
    
    # Adicionar cópia (CC) se houver
    if copia:
        destinatarios_cc = copia.split(';')
        message.CC = ";".join(destinatarios_cc)

    nome = nome.split()
    nome = nome[0]
    nome = nome.capitalize()

    # Configurar o corpo do email como HTML
    corpo_com_variaveis = corpo_html.format(nome=nome, unidade=unidade,imagem_base64=imagem_base64)
    message.HTMLBody = corpo_com_variaveis
    
    # Adicionar anexos, se existirem
    anexos_lista = anexos.split(';')

    # Criar objeto Attachments
    anexos_outlook = message.Attachments

    for anexo_path in anexos_lista:
        if os.path.exists(anexo_path):
            anexos_outlook.Add(anexo_path)
        else:
            print(f"Arquivo de anexo não encontrado: {anexo_path}")
            break
    # Esperar 5 segundos
    time.sleep(1)
    # Enviar o e-mail
    message.Send()
    
    tempo_atual = time.time() - tempo_inicio 
    duracao = f"{int(tempo_atual // 3600):02d}:{int((tempo_atual % 3600) // 60):02d}:{int(tempo_atual % 60):02d}"
    print(f"Tempo: {duracao} .Email enviado para {destinatario} com sucesso!")

# Ler dados do Excel
excel_path = 'C:\send_mail\dados.xlsx'
wb = openpyxl.load_workbook(excel_path)
sheet = wb.active
tempo_inicio = time.time()
#converte para base64
with open('C:\send_mail\logo_mail.png', 'rb') as img_file:
    imagem_base64 = base64.b64encode(img_file.read()).decode('utf-8')
# caminho HTML
corpo_html = ler_corpo_html('C:\send_mail\corpo_mail.html')

# Iterar sobre as linhas do Excel (assumindo que a primeira linha é o cabeçalho)
for row in sheet.iter_rows(min_row=4, values_only=True):
    nome, destinatario, copia, unidade,assunto,anexos = row
    enviar_email(nome, destinatario, copia, unidade,assunto,anexos)
    
# Fechar o arquivo Excel após a conclusão
wb.close()
