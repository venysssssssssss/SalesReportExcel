# SalesReportExcel

import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.chart import BarChart, Reference
import string

def automate_excel(file_name: str) -> None:
    """
    Automatiza a criação de relatórios Excel a partir de um arquivo de dados mensal.

    Parâmetros:
    - file_name (str): Nome do arquivo Excel contendo os dados mensais de vendas.
    """
    # Leitura do arquivo Excel
    excel_file = pd.read_excel(file_name)
    
    # Criação da tabela dinâmica
    report_table = excel_file.pivot_table(index='Gender', columns='Product line', values='Total', aggfunc='sum').round(0)
    
    # Separando o mês e a extensão do nome do arquivo
    month_and_extension = file_name.split('_')[1]
    
    # Salva a tabela dinâmica no arquivo Excel
    report_table.to_excel(f'report_{month_and_extension}', sheet_name='Report', startrow=4)
    
    # Carregamento do workbook e seleção da planilha
    wb = load_workbook(f'report_{month_and_extension}')
    sheet = wb['Report']
    
    # Referências das células (planilha original)
    min_column = wb.active.min_column
    max_column = wb.active.max_column
    min_row = wb.active.min_row
    max_row = wb.active.max_row
    
    # Adição de um gráfico de barras
    barchart = BarChart()
    data = Reference(sheet, min_col=min_column+1, max_col=max_column, min_row=min_row, max_row=max_row) # incluindo cabeçalhos
    categories = Reference(sheet, min_col=min_column, max_col=min_column, min_row=min_row+1, max_row=max_row) # não incluindo cabeçalhos
    barchart.add_data(data, titles_from_data=True)
    barchart.set_categories(categories)
    sheet.add_chart(barchart, "B12")  # localização do gráfico
    barchart.title = 'Vendas por Linha de Produto'
    barchart.style = 2  # escolhe o estilo do gráfico
    
    # Aplicação de fórmulas
    alphabet = list(string.ascii_uppercase)
    excel_alphabet = alphabet[0:max_column]
    
    for i in excel_alphabet:
        if i != 'A':
            sheet[f'{i}{max_row+1}'] = f'=SUM({i}{min_row+1}:{i}{max_row})'
            sheet[f'{i}{max_row+1}'].style = 'Currency'
    
    sheet[f'{excel_alphabet[0]}{max_row+1}'] = 'Total'
    
    # Obtendo o nome do mês
    month_name = month_and_extension.split('.')[0]
    
    # Formatação do relatório
    sheet['A1'] = 'Relatório de Vendas'
    sheet['A2'] = month_name.title()
    sheet['A1'].font = Font('Arial', bold=True, size=20)
    sheet['A2'].font = Font('Arial', bold=True, size=10)
    
    # Salvando o relatório
    wb.save(f'report_{month_and_extension}')
    return

# Exemplo de uso para um ano inteiro
automate_excel('/content/sales_2021.xlsx')

# Exemplo de uso para relatórios mensais individuais
automate_excel('/content/sales_january.xlsx')
automate_excel('/content/sales_february.xlsx')
automate_excel('/content/sales_march.xlsx')

# Opção: Concatenando relatórios mensais e criando um relatório para o ano
excel_file_1 = pd.read_excel('sales_january.xlsx')
excel_file_2 = pd.read_excel('sales_february.xlsx')
excel_file_3 = pd.read_excel('sales_march.xlsx')

new_file = pd.concat([excel_file_1, excel_file_2, excel_file_3], ignore_index=True)
new_file.to_excel('sales_2021.xlsx')
automate_excel('sales_2021.xlsx')
