import os
import openpyxl
import xlwings
from string import ascii_uppercase
import openpyxl.styles
from lib import to_power_symbol
from sympy import interpolate, simplify
from sympy.abc import x

filename = input('Filename: ')

if not os.path.isdir('./excel'):
    os.mkdir('./excel')

data = {
    'x': [1, 2, 3],
    'y': [1, 4, 9]
}

x_table = data.get('x')
y_table = data.get('y')

def get_delta(y: list):
    deltas = []
    while len(y) > 1:
        next_deltas = []
        for i in range(len(y) - 1):
            result = y[i + 1] - y[i]
            next_deltas.append(result)
        deltas.append(next_deltas)
        y = next_deltas
    return deltas

all_deltas = get_delta(y_table)

func = interpolate([(x_table[i], y_table[i]) for i in range(len(x_table))], x)
    
wb = openpyxl.Workbook()
ws = wb.active
ws.title = filename

ws['A1'] = 'x'
for i in range(len(x_table)):
    ws[f'A{i + 2}'] = x_table[i]
ws['B1'] = 'y'
for i in range(len(y_table)):
    ws[f'B{i + 2}'] = x_table[i]

for i in range(len(x_table) - 1):
    column = ascii_uppercase[i + 2]
    cell = f"{column}1"
    ws[cell] = f'Δ{to_power_symbol(i + 1)}y'
    for j in range(len(all_deltas[i])):
        row = j + 2
        column = ascii_uppercase[i + 1]
        if (ws[f'{column}{row + 1}'].value and ws[f'{column}{row}'].value):
            ws[f"{ascii_uppercase[i + 2]}{j + 2}"] = f"={column}{row + 1}-{column}{row}"

for i in ws[f'A1:{ascii_uppercase[len(x_table)]}1']:
    for j in i:
        j.fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='f8ff7a')

ws[f'A{len(x_table) + 3}'] = "f(x) ="
ws[f'B{len(x_table) + 3}'] = str(func)

ws[f'B{len(x_table) + 3}'].fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='58fc84')
ws[f'A{len(x_table) + 3}'].fill = openpyxl.styles.PatternFill(fill_type='solid', start_color='58fc84')

wb.save(f'./excel/{filename}.xlsx')

with xlwings.App(visible=True) as app:
    wb = app.books.open(f'./excel/{filename}.xlsx') 
    ws = wb.sheets[0]

    f = str(func).replace('**', '^')
    vba_code = f'''
    Function F(x As Double) As Double
        F = {f}
    End Function
    '''
    
    vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 - это добавление нового модуля
    vba_module.CodeModule.AddFromString(vba_code)
    
    wb.save(r'./excel/custom_function.xlsm')
    wb.close()