from datetime import datetime
from planilha import Planilha

nome_empresa = 'Dock'
numero_matricula = '123456'
nome_completo = 'guilherme de lima guedes'
localidade = 'Goiana PE'
diretorio_planilha = '/home/guilherme/teste.xlsx' # ex: '/home/teste.xlsx'

planilha = Planilha(diretorio_planilha)

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
    chamado = input('Digite o Nº do chamado: ')
    observacao = input('Digite uma Observação: ')
    
    print('Deseja bater saida ponto?')
    print('Sair(X) Bater saida(Enter)')
    entrada = input()
    
    if entrada == 'X' or entrada == 'x':
        exit()
    else:
        data_hora_atual =  datetime.now()
        data_final = data_hora_atual.strftime('%d/%m/%Y')
        hora_final = data_hora_atual.strftime('%H:%M')
        planilha.adicionar_valores(nome_empresa,numero_matricula,nome_completo,localidade,data_inicio,hora_inicio,data_final,hora_final,chamado,observacao)
        print(f'Sua hora e data de saida é {data_final} ás {hora_final}.\n')
        main()

main()