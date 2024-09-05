#Ler dados da planilha
#Inserir cada célula de cada linha em um campo do sistema
import openpyxl
import pyautogui

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

for linha in vendas_sheet.iter_rows(min_row=2):
    pyautogui.click(x=1528, y=498)
    pyautogui.write(linha[0].value)
    linha[0].value
    linha[1].value
    linha[2].value
    linha[3].value