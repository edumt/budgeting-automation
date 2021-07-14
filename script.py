# TO DO: organize script into multiple files
import openpyxl

statesDict = {
    "ES": "ESPÍRITO SANTO",
    "RJ": "RIO DE JANEIRO",
    "BA": "BAHIA",
    "MG": "MINAS GERAIS"
}

class Client:
    def __init__(self, fullName, aConsumption, state, city = ''):
        self.name = fullName.upper()
        self.aConsumption = aConsumption # Average monthly energy consumption [kWh/month]
        self.state = state.upper() # State initials
        self.city = city.upper() # (District-)City
    
test_client = Client('Eduardo Moura Tavares', 500, 'ES', 'Praia de Itaparica-Vila Velha')
#print(vars(test_client))



wb = openpyxl.load_workbook('excel_template.xlsx')
#print(wb.sheetnames)

consumption_sheet = wb['HCONSUMO']
#print(consumption_sheet)
#print(consumption_sheet['D18'].value)
