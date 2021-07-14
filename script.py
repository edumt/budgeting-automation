import openpyxl

class Client:
    def __init__(self, name, consumption, state, district = ''):
        self.name = name
        self.consumption = consumption
        self.state = state
        self.district = district
    
test_client = Client('EDUARDO MOURA TAVARES', 500, 'ES', 'PRAIA DE ITAPARICA-VILA VELHA')
print(vars(test_client))
#consumption = 500 # Consumo em kWh/mes
#name = 'EDUARDO MOURA TAVARES' # NOME COMPLETO (Caps Lock)

wb = openpyxl.load_workbook('excel_template.xlsx')

#sheets = wb.sheetnames
#print(sheets)
consumption_sheet = wb['HCONSUMO']
#print(consumption_sheet)
print(consumption_sheet['D18'].value)
