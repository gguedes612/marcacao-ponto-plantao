import datetime
from openpyxl import Workbook

data_hora_atual =  datetime.datetime.now()

print(data_hora_atual)

if data_hora_atual.month >= 10:
    data_inicio = f'{data_hora_atual.day}/{data_hora_atual.month}/{data_hora_atual.year}'
else:
    data_inicio = f'{data_hora_atual.day}/0{data_hora_atual.month}/{data_hora_atual.year}'

if data_hora_atual.minute >= 10:
    hora_inicio = f'{data_hora_atual.hour}:{data_hora_atual.minute}'
else:
    hora_inicio = f'{data_hora_atual.hour}:0{data_hora_atual.minute}'


print(data_inicio)
print(hora_inicio)

wb = Workbook()

planilha = wb.worksheets[0]

planilha['A1'] = data_inicio
planilha['A2'] = hora_inicio

wb.save('/home/furiosa/teste.xlsx')