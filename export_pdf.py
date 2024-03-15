# Use Aspose.Cells for Python via Java
import jpype
import asposecells
jpype.startJVM()
from asposecells.api import Workbook, FileFormatType, PdfSaveOptions
from win32com.client import Dispatch
import glob
import os
import time
#Iterar pastas
# Diretório raiz contendo as pastas com as planilhas
diretorio_raiz = 'C:\\Reembolso\\Relatorio_Unidades'

# Listar todas as pastas no diretório raiz
nomes_pastas = [nome for nome in os.listdir(diretorio_raiz) if os.path.isdir(os.path.join(diretorio_raiz, nome))]

# Iniciar uma instância do Excel
excel_app = Dispatch("Excel.Application")
excel_app.Visible = False  # Defina para True se quiser que o Excel seja visível durante o processo

# Iterar sobre as pastas
for nome_pasta in nomes_pastas:

    # Listar arquivos específicos usando padrão (por exemplo, todos os arquivos .txt)
    arquivos_xlsx = glob.glob(os.path.join(diretorio_raiz, nome_pasta, '*.xlsx'))
    # Imprimir a lista de arquivos
    for arquivo in arquivos_xlsx:
        
        try:
            workbook = Workbook(arquivo)
            nome_pdf = arquivo.replace('xlsx','pdf')
            if os.path.exists(nome_pdf):
                os.remove(nome_pdf)
                print(f'O arquivo {nome_pdf} foi excluído.')
            time.sleep(0.5)
            saveOptions = PdfSaveOptions()
            saveOptions.setOnePagePerSheet(True)
            workbook.save(nome_pdf, saveOptions)
            print(arquivo)
        except Exception as error:
            print(f'ERRO no {error}{nome_pdf}')
jpype.shutdownJVM()




