from openpyxl import Workbook, load_workbook


class Planilha:

    def __init__(self,diretorio):
        
        try: 
            self.wb = load_workbook(diretorio)
            self.ws = self.wb['Planilha']
        except FileNotFoundError:
            self.wb = Workbook()
            self.ws = self.wb.create_sheet('Planilha',0)
        
    
    def mostra_planilhas(self):
        return self.wb.sheetnames



planilha = Planilha('/home/guilherme/teste.xlsx')
print(planilha.mostra_planilhas())

