import os
import xlwings
import openpyxl
from lib import get_filename_ext, get_delta, render_lagrange_table, render_newton_table, render_vba_function
from sympy import interpolate
from sympy.abc import x

if not os.path.isdir('./excel'):
    os.mkdir('./excel')

filename = input('Filename: ')

print('1 - Метод Лагранжа')
print('2 - Метод Ньютона')
method = int(input('Method: '))
if method > 2 or method < 1:
    print('Неверно введен метод')
    os._exit(1)

xlsx_filename = get_filename_ext(filename, 'xlsx')
xlsm_filename = get_filename_ext(filename, 'xlsm')

table = {
    'x': [1, 2, 3, 5],
    'y': [1, 4, 9, 25]
}

x_table = table.get('x')
y_table = table.get('y')

all_deltas = get_delta(y_table)

func = interpolate([(x_table[i], y_table[i]) for i in range(len(x_table))], x)

wb = openpyxl.Workbook()
ws = wb.active
ws.title = filename

if method == 1:
    render_lagrange_table(ws, x_table, y_table, func)
elif method == 2:
    render_newton_table(ws, x_table, y_table, all_deltas, func)

wb.save(xlsx_filename)
wb.close()

with xlwings.App(visible=True) as app:
    wb = app.books.open(xlsx_filename) 
    ws = wb.sheets[0]

    render_vba_function(wb, func)
    
    wb.save(xlsm_filename)
    wb.close()
    
if os.path.isfile(os.path.abspath(xlsm_filename)):
    os.remove(os.path.abspath(xlsx_filename))