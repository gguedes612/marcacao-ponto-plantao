from datetime import datetime
from planilha import Planilha

nome_empresa = ''
numero_matricula = ''
nome_completo = ''
localidade = ''

planilha = Planilha('/home/guilherme/teste.xlsx')

def main():
    print('Deseja sair ou bater entrada?')
    print('Sair(X) Bater entrada(Enter)')
    entrada = input('')
    if entrada == 'X' or entrada == 'x':
        planilha.salvar_planilha()
        exit()
    else:
        bater_entrada()

def bater_entrada():
    data_hora_atual =  datetime.now()
    data_inicio = data_hora_atual.strftime('%d/%m/%Y')
    hora_inicio = data_hora_atual.strftime('%H:%M')
    print(f'Sua hora e data de entrada é {data_inicio} ás {hora_inicio}.\n')
    bater_saida(data_inicio,hora_inicio)

def bater_saida(data_inicio,hora_inicio):
    print('Deseja bater saida ponto?')
    print('Sair(X) Bater saida(Enter)')
    entrada = input()
    if entrada == 'X' or entrada == 'x':
        exit()
    else:
        data_hora_atual =  datetime.now()
        data_final = data_hora_atual.strftime('%d/%m/%Y')
        hora_final = data_hora_atual.strftime('%H:%M')
        planilha.adicionar_valores(nome_empresa,numero_matricula,nome_completo,localidade,data_inicio,hora_inicio,data_final,hora_final)
        print(f'Sua hora e data de saida é {data_inicio} ás {hora_inicio}.\n')
        main()

main()