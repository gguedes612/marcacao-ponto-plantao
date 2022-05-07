from openpyxl import Workbook, load_workbook

class Planilha:

    def __init__(self,diretorio):
        
        self.diretorio = diretorio
        try: 
            self.wb = load_workbook(self.diretorio)
            self.ws = self.wb['Planilha']
        except FileNotFoundError:
            self.wb = Workbook()
            self.ws = self.wb.create_sheet('Planilha',0)   
       
    def mostra_planilhas(self):
        return self.wb.sheetnames

    def mostrar_celula(self,celula):
        return self.ws[celula].value

    def salvar_planilha(self):
        self.wb.save(self.diretorio)
    
    def adicionar_valores(self,nome_empresa,numero_matricula,nome_completo,localidade,data_inicio,hora_inicio,data_final,hora_final,linha=1):
        if self.ws[f'A{linha}'].value == None:
            self.ws[f'A{linha}'] = nome_empresa
            self.ws[f'B{linha}'] = numero_matricula
            self.ws[f'C{linha}'] = nome_completo
            self.ws[f'D{linha}'] = localidade
            self.ws[f'E{linha}'] = 'Acionamento'
            self.ws[f'F{linha}'] = data_inicio
            self.ws[f'G{linha}'] = f'=IF(F{linha}="";"";WEEKDAY(F{linha}))'
            self.ws[f'H{linha}'] = hora_inicio
            self.ws[f'I{linha}'] = data_final
            self.ws[f'J{linha}'] = f'=IF(I{linha}="";"";WEEKDAY(I{linha}))'
            self.ws[f'K{linha}'] = hora_final
            
        else:
            self.adicionar_valores(nome_empresa,numero_matricula,nome_completo,localidade,data_inicio,hora_inicio,data_final,hora_final,linha+1)