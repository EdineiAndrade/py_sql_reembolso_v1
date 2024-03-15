import pandas as pd
import sqlite3
from datetime import datetime
import time
import os

# Leitura das planilhas Excel
planilha2 = pd.read_excel("C:\\Reembolso\\Bases\\Lista_Propostas.xlsx")
agora = datetime.now()

#tempo inicial
tempo_inicio = time.time()
# Loop sobre as linhas da tabela1
for i, linha_tabela2 in planilha2.iterrows():
    nome_S476 = linha_tabela2['Cliente']
    sicad_S476 = linha_tabela2['SICAD']
    sicad_S476 = sicad_S476.replace('-','')
    cpf_S476 = linha_tabela2['CPF'].replace('.','')
    planilha2.at[i,'SICAD'] = sicad_S476
    planilha2.at[i,'CPF'] = cpf_S476
    
planilha2.to_excel("C:\\Reembolso\\Bases\\Lista_Propostas.xlsx", index=False)

planilha1 = pd.read_excel("C:\\Reembolso\\Relatorio_Geral\\Reembolso_Gerencial_Atualizado.xlsx")
planilha2 = pd.read_excel("C:\\Reembolso\\Bases\\Lista_Propostas.xlsx")
planilha3 = pd.read_excel("C:\Reembolso\Bases\BASE_DAP.xlsx")


quantidade_linhas = planilha1.shape[0]

# Criar conexão com um banco de dados temporário SQLite VERIFICAR PROPOSTAS NO S476
conexao1 = sqlite3.connect(':memory:')

# Salvar os DataFrames como tabelas temporárias no banco de dados
planilha1.to_sql('tabela1_temp', conexao1, index=False, if_exists='replace')
planilha2.to_sql('tabela2_temp', conexao1, index=False, if_exists='replace')

# Criar conexão com um banco de dados temporário SQLite VERIFICAR PROPOSTAS NO S476
conexao2 = sqlite3.connect(':memory:')

# Salvar os DataFrames como tabelas temporárias no banco de dados
planilha1.to_sql('tabela3_temp', conexao2, index=False, if_exists='replace')
planilha3.to_sql('tabela4_temp', conexao2, index=False, if_exists='replace')


# Loop sobre as linhas da tabela1
for indice, linha_tabela1 in planilha1.iterrows():
    nome = linha_tabela1['Cliente']
    sicad = linha_tabela1['Código do Cliente']
    cpf = linha_tabela1['CPF_CNPJ']
    
    restante = quantidade_linhas - indice
    
    if nome == 'JACKSON KLEVERSON SANTANA':
       print(f'nome:{nome} sicad:{sicad} cpf:{cpf}')
       print(f'nome:{nome} sicad:{sicad} cpf:{cpf}')
 # Consulta SQL para a linha atual da tabela1 - LITA DE PROPOSTAS S476
    consulta_sql1 = f'''
        SELECT STATUS_PROPOSTA
        FROM tabela2_temp
        WHERE tabela2_temp.SICAD = '{sicad}' AND tabela2_temp.CPF = '{cpf}';
    '''
     # Consulta SQL para a linha atual da tabela1 - LITA DE PROPOSTAS S476
    consulta_sql2 = f'''
        SELECT STATUS_DAP,VALIDADE
        FROM tabela4_temp
        WHERE tabela4_temp.SICAD = '{sicad}' AND tabela4_temp.CPF = '{cpf}';
    '''
    resultado1 = pd.read_sql_query(consulta_sql1, conexao1)
    resultado2 = pd.read_sql_query(consulta_sql2, conexao2)
    try:
      resultado_status_dap = resultado2['STATUS_DAP'][0]
      resultado_vencimento_dap = resultado2['VALIDADE'][0]
      planilha1.at[indice,'DAP'] = resultado_status_dap
      resultado_vencimento_dap = resultado_vencimento_dap.replace('00:00:00','')
      resultado_vencimento_dap = resultado_vencimento_dap.replace('VENCIDA:','')
      resultado_vencimento_dap = resultado_vencimento_dap.replace(' ','')
      data_dap_split = resultado_vencimento_dap.split('-')
      ano = int(data_dap_split[0])  
      mes = int(data_dap_split[1])  
      dia = int(data_dap_split[2])  
      data_dap_formatada = datetime(ano,mes,dia) 
      data_atual = datetime.now()
      if data_dap_formatada < data_atual:
          resultado_status_dap = f'Vencida{resultado_status_dap[-2:]}'
      planilha1.at[indice,'DAP'] = resultado_status_dap
      planilha1.at[indice,'Vecimento DAP'] = data_dap_formatada
    except:
        resultado_status_dap = ""
    try:
        resultado_status_proposta = resultado1['STATUS_PROPOSTA'][0]
        if resultado_status_proposta == "Elaborada - em ser":
            resultado_status_proposta = "El.em ser"
        elif resultado_status_proposta == "Encaminhada para Deferimento" or  resultado_status_proposta == "Em Deferimento":
            resultado_status_proposta = "Deferimento"    
        elif resultado_status_proposta == "Aprovada Sicor":
            resultado_status_proposta = "Apr. Sicor"
        elif resultado_status_proposta == "Devolvida da Validação" or resultado_status_proposta == "Devolvida do Deferimento" or resultado_status_proposta == "Devolvida Sicor":
            resultado_status_proposta = "Devolvida"    
        planilha1.at[indice,'S476'] = resultado_status_proposta
    except:
        resultado_status_proposta = ""
    tempo_atual = time.time() - tempo_inicio 
    duracao = f"{int(tempo_atual // 3600):02d}:{int((tempo_atual % 3600) // 60):02d}:{int(tempo_atual % 60):02d}"
    print(f"Tempo: {duracao} consulta: {indice} faltam: {restante} Cliente: {nome}")    
planilha1.to_excel("C:\\Reembolso\\Relatorio_Geral\\Reembolso_Gerencial_Atualizado.xlsx", index=False)
print(f"Tempo: {duracao}. Planilha atualizada. Processo finalizado.")