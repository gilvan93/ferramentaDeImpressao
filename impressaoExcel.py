import os
import win32com.client

# Caminho da pasta com os arquivos Excel
pasta_excel = r'G:\Gilvan\2025\relatorios\Julho\Juliana'

# Inicializa o Excel
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False  # Deixa o Excel invisível (pode mudar para True se quiser ver)

# Percorre todos os arquivos da pasta
for arquivo in os.listdir(pasta_excel):
    if arquivo.endswith('.xlsx') or arquivo.endswith('.xls'):
        caminho_arquivo = os.path.join(pasta_excel, arquivo)
        print(f'Imprimindo: {arquivo}')

        # Abre o arquivo
        wb = excel.Workbooks.Open(caminho_arquivo)

        # Para cada planilha no workbook
        for ws in wb.Worksheets:
            # Define a orientação da página como paisagem
            ws.PageSetup.Orientation = 2  # 2 = xlLandscape

        # Imprime o workbook
        wb.PrintOut()

        # Fecha sem salvar
        wb.Close(SaveChanges=False)

# Encerra o Excel
excel.Quit()
print("Todos os arquivos foram enviados para impressão.")

#testes feitos