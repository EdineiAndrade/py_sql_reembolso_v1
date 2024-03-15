import pandas as pd
from win32com.client import Dispatch
import win32com.client as win32
from openpyxl import load_workbook
from openpyxl import Workbook
from datetime import datetime
import time
import os
import shutil

df = pd.read_excel('C:\send_mail\dados_envio_agente.xlsx')
df2 = pd.read_excel("C:\\Reembolso\\Relatorio_Geral\\Reembolso_Gerencial_Atualizado.xlsx")
quantidade_linhas = df.shape[0]
# Obtém a data atual
data_atual = datetime.now()
# Formata a data e hora no formato desejado (dd-mm-yy_h-m-s)
data_formatada = data_atual.strftime("%d-%m-%y")
# Formata a data no formato desejado
pasta_reembolso= 'C:\\Users\\inec\\OneDrive - Instituto Nordeste Cidadania\\AGROAMIGO\RELATÓRIOS_2024\\_REEMBOLSO'

#tempo inicial
tempo_inicio = time.time()

if os.path.exists(pasta_reembolso):
        
        # Se a pasta existir, exclua-a
        shutil.rmtree(pasta_reembolso)
       
#Cria uma nova pasta
os.mkdir(pasta_reembolso)

for indice, linha in df.iterrows():
    
    unidade = linha['UNIDADE']
    total_agentes = df['AGENTE'].count()
    if indice == total_agentes:
        excel_app.Quit()
        print(f'Processo Finalizado. Tempo: {duracao} Processo: {indice + 1} Agente: {Agente}')
        break
    Agente = linha['AGENTE'].strip()
    unidade = unidade.strip()
    pasta_unidade = os.path.join(pasta_reembolso, unidade)
    
    if not os.path.exists(pasta_unidade):
        # Cria uma nova pasta se ela não existir
        os.mkdir(pasta_unidade)
    caminho_arquivo_unidade = os.path.join(pasta_unidade, f'_Reembolso_{unidade}.xlsx')
    if not os.path.exists(caminho_arquivo_unidade):
        dados_unidade = df2.loc[df2['Unidade Ativa']==unidade]
        dados_unidade.to_excel(caminho_arquivo_unidade, index=False)

        # Iniciar uma instância do Excel
        excel_app = Dispatch("Excel.Application")
        excel_app.Visible = False  # Se desejar que o Excel seja visível durante o processo
        # Abrir a planilha existente para obter formatação
        existing_workbook = excel_app.Workbooks.Open('C:\\Reembolso\\Bases\\base_teste.xlsx')
        existing_worksheet = existing_workbook.Sheets(1)
        #Formatar Planilha UNIDADE

        # Copiar a formatação da planilha existente
        existing_worksheet.UsedRange.Rows[1].EntireRow.Copy()
        #existing_worksheet.UsedRange.Copy()
        # Abrir o arquivo Excel AGENTE
        workbook_unidade = excel_app.Workbooks.Open(caminho_arquivo_unidade)
        worksheet_unidade = workbook_unidade.Sheets(1)
        #existing_formatting
        worksheet_unidade.Range("A:AD").PasteSpecial(Paste=-4122)
        excel_app.CutCopyMode = False

        # Ajustar automaticamente a largura das colunas
        worksheet_unidade.Columns.AutoFit()

        #Salvar e sair da planilha Unidade
        workbook_unidade.Save()
        workbook_unidade.Close()
        existing_workbook.Close()

    #Exportar po agente
    dados_agente = df2.loc[df2['Agente de Crédito']==Agente] #Incluir o filtro de data
    linhas_dados_agente = dados_agente.shape[0]
    if linhas_dados_agente <= 1:
         continue
    caminho_arquivo = os.path.join(pasta_unidade, f'{Agente}.xlsx')    
    dados_agente.to_excel(caminho_arquivo, index=False)
    #dados_agente = df2.loc[(df2['Agente de Crédito']==Agente) & (df2['Dias']<=30) & (df2['Status']!='PAGO' )] #Incluir o filtro de data
    caminho_destino_pdf = os.path.join(pasta_unidade, f'{Agente}.pdf') 
    
    # Copiar a formatação da planilha existente
    existing_workbook = excel_app.Workbooks.Open('C:\\Reembolso\\Bases\\base_teste.xlsx')
    existing_worksheet = existing_workbook.Sheets(1)
    existing_worksheet.UsedRange.Rows[1].EntireRow.Copy()
    #existing_worksheet.UsedRange.Copy()
    # Abrir o arquivo Excel AGENTE
    workbook = excel_app.Workbooks.Open(caminho_arquivo)
    worksheet = workbook.Sheets['Sheet1']
    #existing_formatting
    worksheet.Range("A:AD").PasteSpecial(Paste=-4122)
    excel_app.CutCopyMode = False

    #Filtar Dias e Status
    worksheet.Range("A:AD").AutoFilter(Field=2, Criteria1="<=40")
    worksheet.Range("A:AD").AutoFilter(Field=5, Criteria1="<>PAGO")
    # Ajustar automaticamente a largura das colunas
    worksheet.Columns.AutoFit()
    #Colunas por página

    # Encontrar a última coluna com informações na primeira linha
    last_column = worksheet.Cells(1, worksheet.Columns.Count).End(-4159).Column  # -4159 para xlToLeft

    # Configurar a área de impressão para todas as colunas
    worksheet.PageSetup.PrintArea = f"A1:AD{worksheet.Rows.Count}"

    # Configurar layout da página para ajustar uma página por largura e em paisagem
    worksheet.PageSetup.Orientation = 2
    #worksheet.PageSetup.FitToPagesWide = 1
    worksheet.PageSetup.FitToPagesTall = False
    worksheet.PageSetup.Zoom = 40
    existing_workbook.Close()
    # Configurar as margens de impressão (em polegadas)
    worksheet.PageSetup.LeftMargin = 10
    worksheet.PageSetup.RightMargin = 10
    worksheet.PageSetup.TopMargin = 10
    worksheet.PageSetup.BottomMargin = 10
    workbook.Save()
    #Configurar area de impressão
    
    # Encontrar a última linha com informações na coluna "A"
    last_row = worksheet.Cells(worksheet.Rows.Count, "A").End(-4162).Row  # -4162 para xlUp
    sheet = workbook.Sheets(1)
    # Loop através de todas as linhas
    
    # Configurar a área de impressão usando Dispatch
    worksheet.PageSetup.PrintArea = f"A1:AD{last_row}"
    workbook.Sheets(1).Columns("F:G").Hidden = True
    workbook.Sheets(1).Columns("J").Hidden = True
    workbook.Sheets(1).Columns("L:M").Hidden = True
    workbook.Sheets(1).Columns("Q").Hidden = True
    workbook.Sheets(1).Columns("U:W").Hidden = True
    workbook.Sheets(1).Columns("AC:AD").Hidden = True

    # largura da
    workbook.Sheets(1).Columns("Y").ColumnWidth = 30    

    # Salvar como PDF    
    workbook.ExportAsFixedFormat(0, caminho_destino_pdf)
    
    workbook.Close(SaveChanges=False)
    #time.sleep(.01)

# Fechar a aplicação do Excel

    tempo_atual = time.time() - tempo_inicio 
    duracao = f"{int(tempo_atual // 3600):02d}:{int((tempo_atual % 3600) // 60):02d}:{int(tempo_atual % 60):02d}"

    print(f'Tempo: {duracao} Processo: {indice + 1} Agente: {Agente}')
excel_app.Quit()
print(f'Processo Finalizado. Tempo: {duracao} Processo: {indice + 1} Agente: {Agente}')