# https://automatetheboringstuff.com/chapter12/
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
    
def populateSheet(client):
    print(client.name)

def generateReport(client, sheet):
    pass

CONST_CONSUMPTION_CELL = 'D18'
CONST_NAME_CELL = 'C4'
CONST_PERCENTAGE_INCREASE_CELL = 'H44'
CONST_BUDGET_CELL = 'B3:H32'

test_client = Client('Eduardo Moura Tavares', 500, 'ES', 'Praia de Itaparica-Vila Velha')
#print(vars(test_client))
populateSheet(test_client)

wb = openpyxl.load_workbook('../excel_template.xlsx')
#print(wb.sheetnames)

consumption_sheet = wb['HCONSUMO']
#print(consumption_sheet)
print(consumption_sheet[CONST_CONSUMPTION_CELL].value)
