import pandas as pd
import sqlite3
from  datetime import datetime
import time

# Leitura das planilhas Excel
planilha1 = pd.read_excel("C:\\Reembolso\\Bases\\Base_Reembolso.xlsx")
planilha2 = pd.read_excel("C:\\Reembolso\\Bases\\Valores_Recebidos.xlsx")

quantidade_linhas = planilha1.shape[0]
# Criar conexão com um banco de dados temporário SQLite
conexao = sqlite3.connect(':memory:')

# Salvar os DataFrames como tabelas temporárias no banco de dados
planilha1.to_sql('tabela1_temp', conexao, index=False, if_exists='replace')
planilha2.to_sql('tabela2_temp', conexao, index=False, if_exists='replace')

#tempo inicial
tempo_inicio = time.time()
# Loop sobre as linhas da tabela1
for indice, linha_tabela1 in planilha1.iterrows():
    nome = linha_tabela1['Cliente']
    sicad = linha_tabela1['Código do Cliente']
    sicad = int(str(sicad)[:-1])
    operacao = linha_tabela1['Código do Contrato']
    unidade = linha_tabela1['Unidade Ativa']
    Agente = linha_tabela1['Agente de Crédito']
    data_previsao = linha_tabela1['Data_Previsão']
    Status_Pagamento = linha_tabela1['Status Pagamento']
    percentual_Reembolso = linha_tabela1['%']
    Valor_Previsto = linha_tabela1['Previsto']
    Valor_Efetivo = linha_tabela1['Efetivo']
    #Calcular dias q faltam
    data_previsao = linha_tabela1['Data_Previsão']
    dias_parcela = data_previsao - datetime.today()
    hoje = datetime.now()
    dias_parcela = (data_previsao - hoje).days
    planilha1.at[indice,'Dias'] = dias_parcela + 1
    #Extrair município
    Endereco_Completo = linha_tabela1['Endereço Completo'] 
    nm_municipio = Endereco_Completo.split("-")
    nm_municipio = nm_municipio[len(nm_municipio)-2]
    nm_municipio = nm_municipio[4:]
    nm_municipio = nm_municipio.strip()
    planilha1.at[indice,'Município'] = nm_municipio
    nm_municipio = nm_municipio.strip()
    carteira = linha_tabela1['Programa de Crédito']
    #Correção endereço
    endereco_compact = Endereco_Completo.split("-")
    endereco_compact = endereco_compact[0][:-6]
    planilha1.at[indice,'Endereço Completo'] = endereco_compact
    #Setar o que falta para exibir no print
    restante = quantidade_linhas - indice
    nm_opc_condicoes = "('CLIENTE', 'BONUS DE ADIMPLENCIA-FNE', 'RENEGOCIACAO DE DIVIDAS')"

    # Consulta SQL para a linha atual da tabela1
    consulta_sql = f'''
        SELECT VR_REA_MN, DT_VAL_LNC
        FROM tabela2_temp
        WHERE tabela2_temp.cd_cli = '{sicad}' AND tabela2_temp.CD_CTR = '{operacao}' AND tabela2_temp.NM_OPC IN {nm_opc_condicoes};
    '''
    # Executar a consulta WHERE tabela2_temp.nm_cli = 'FLAVIO GONCALVES DOS SANTOS' AND tabela2_temp.cd_cli = '11641102' AND tabela2_temp.CD_CTR = 'C300005001' AND tabela2_temp.NM_OPC = 'CLIENTE';
    resultado = pd.read_sql_query(consulta_sql, conexao)
    # Pegar valor recebido
    soma_resultado = resultado['VR_REA_MN'].astype(float).sum()
    agora = datetime.now()

    #ataFrame chamado df e uma coluna chamada 'coluna_texto'
    #df_resultado = df[df['coluna_texto'].str.contains('sua_palavra', case=False)]
    #Ajustar Agentes


    if any(item in carteira for item in ['GRUPO "B"', '-GRP.B', 'GRUPO B','PRONAF-B']):
        if nm_municipio == 'MAETINGA' or nm_municipio == 'CARAIBAS':
            planilha1.at[indice,'Agente de Crédito'] = 'ALEXANDRE SANTOS ROCHA'
        elif nm_municipio == 'CORDEIROS' or nm_municipio == 'PIRIPA':
            planilha1.at[indice,'Agente de Crédito'] = 'NATALINE SOARES MALTA'
        elif nm_municipio == 'BOM JESUS DA SERRA' or nm_municipio == 'MIRANTE':
            planilha1.at[indice,'Agente de Crédito'] = 'JORDIMAR MONIZO DE FARIAS'
        elif nm_municipio == 'BELO CAMPO' or nm_municipio == 'CANDIDO SALES':
            planilha1.at[indice,'Agente de Crédito'] = 'ALLAN ALVES SOUSA'
        elif nm_municipio == 'JUSSIAPE' or nm_municipio == 'IBICOARA':
            planilha1.at[indice,'Agente de Crédito'] = 'LUIS CARLOS MOREIRA DA SILVA'
        elif nm_municipio == 'CANAPOLIS':
            planilha1.at[indice,'Agente de Crédito'] = 'GILVAN DOS SANTOS PEREIRA'
        elif nm_municipio == 'SERRA DOURADA':
            planilha1.at[indice,'Agente de Crédito'] = 'DARKSON BATISTA DA SILVA'
        elif nm_municipio == 'COCOS':
            planilha1.at[indice,'Agente de Crédito'] = 'RAFAEL OLIVEIRA DA SILVA'
        elif nm_municipio == 'LIVRAMENTO DE NOSSA SENHORA':
            planilha1.at[indice,'Agente de Crédito'] = 'SIDINEY DE SOUZA SILVA'
        elif nm_municipio == 'LAGOA REAL':
            planilha1.at[indice,'Agente de Crédito'] = 'MAXSUEL NEVES DE AGUIAR'
        elif nm_municipio == 'TANHACU':
            planilha1.at[indice,'Agente de Crédito'] = 'GILSON DA SILVA PEREIRA'
        elif nm_municipio == 'BARRA DA ESTIVA':
            planilha1.at[indice,'Agente de Crédito'] = 'RAFAEL SANTOS PEREIRA'
        elif nm_municipio == 'LAJEDO DO TABOCAL' or nm_municipio == 'ITIRUCU' or nm_municipio == 'LAGEDO DO TABOCAL' or nm_municipio == 'CRAVOLANDIA'  or nm_municipio == 'SANTA INES':
            planilha1.at[indice,'Agente de Crédito'] = 'EVERTON ALMEIDA DE SOUZA'

    #pegar a data da valorização
    #if nome == 'MARIA APARECIDA DANTAS DE OLIVEIRA':
        #print(nome)
    data_valorizacao = resultado['DT_VAL_LNC'].max()
    if soma_resultado > 0:
        data_valorizacao = data_valorizacao[:10]
        data_valorizacao = data_valorizacao.split('-')
        data_valorizacao = f'{data_valorizacao[2]}/{data_valorizacao[1]}/{data_valorizacao[0]}'
        planilha1.at[indice,'Data_Pagamento'] = data_valorizacao
        planilha1.at[indice,'Valor_Recebido'] = soma_resultado

        if soma_resultado < Valor_Previsto * .97:
            planilha1.at[indice,'Status'] = "->VERIFICAR"
        elif soma_resultado >= Valor_Previsto * .97:
            planilha1.at[indice,'Status'] = "PAGO"
            planilha1.at[indice,'Data_Pagamento'] = data_valorizacao
    elif data_previsao < agora and Status_Pagamento != 'PAGO' and percentual_Reembolso < 1:
        planilha1.at[indice,'Status'] = "ATRASO"  

    if  percentual_Reembolso >= 1:
        planilha1.at[indice,'Status'] = 'PAGO'
        planilha1.at[indice,'Valor_Recebido'] = Valor_Efetivo
        planilha1.at[indice,'Data_Pagamento'] = data_valorizacao
    #atualizar dados no excel
    
    tempo_atual = time.time() - tempo_inicio 
    duracao = f"{int(tempo_atual // 3600):02d}:{int((tempo_atual % 3600) // 60):02d}:{int(tempo_atual % 60):02d}"
    if  soma_resultado <= 0:
        print(f"Tempo: {duracao} consulta: {indice} faltam: {restante} Cliente: {nome} resultado: {soma_resultado}")
    else:
        print(f"Tempo: {duracao} consulta: {indice} faltam: {restante} Cliente: {nome} resultado: {soma_resultado}")

#Data hora salvar arquivo
data_hora_formatada = agora.strftime("%d-%m-%Y_%H-%M-%S")       
#salvar planilha
planilha1.to_excel("C:\\Reembolso\\Relatorio_Geral\\Reembolso_Gerencial_Atualizado.xlsx", index=False)

   
# Commit e fechar a conexão
conexao.commit()
conexao.close()

# Imprimir a soma
print("Atualização concluída.")
