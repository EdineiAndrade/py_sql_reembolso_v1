import pandas as pd
import sqlite3

# Carregar dados das planilhas para DataFrames do pandas
planilha1 = pd.read_excel('Reembolso_Gerencial.xlsx')
planilha2 = pd.read_excel('Valores_Recebidos.xlsx')

# Criar conexão com um banco de dados temporário SQLite
con = sqlite3.connect(':memory:')

# Salvar DataFrames como tabelas no banco de dados
planilha1.to_sql('tabela1', con, index=False)
planilha2.to_sql('tabela2', con, index=False)

# Iterar sobre as linhas da Planilha 1
for index, row in planilha1.iterrows():
    # Condições da consulta SQL
    condicoes_sql = f"""
        sicad = {row['Código do Cliente']} AND
        operacao = '{row['Código do Contrato']}' AND
        tipo = 'CLIENTE'
    """
    
    # Executar consulta SQL para verificar condições
    consulta_sql = f"""
        SELECT SUM(valor) as soma_valor
        FROM tabela2
        WHERE {condicoes_sql}
    """
    
    resultado = pd.read_sql_query(consulta_sql, con)
    
    # Obter a soma dos valores
    soma_valor = resultado['soma_valor'].values[0] if not resultado['soma_valor'].isna().any() else 0
    
    # Imprimir resultado
    print(f"Para a linha {index + 1}: Soma dos valores na Planilha 2 = {soma_valor}")

# Fechar a conexão
con.close()
