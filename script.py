# https://automatetheboringstuff.com/chapter12/
# TO DO: organize script into multiple files

import os
import datetime
import openpyxl

CONST_SHEET_VERSION = '13-JULHO-2021'
CONST_CONSUMPTION_CELL = 'D18'
CONST_NAME_CELL = 'C4'
CONST_PERCENTAGE_INCREASE_CELL = 'H44'
CONST_BUDGET_CELL = 'B3:H32'

statesDict = {
    "ES": "ESPÍRITO SANTO",
    "RJ": "RIO DE JANEIRO",
    "BA": "BAHIA",
    "MG": "MINAS GERAIS"
}

class Client:
    def __init__(self, fullName, consumption, state, city = ''):
        self.fullName = fullName.upper()
        self.consumption = consumption # Average monthly energy consumption [kWh/month]
        self.state = state.upper() # State initials
        self.city = city.upper() # (District-)City
    
def populateSheet(client):
    wb = openpyxl.load_workbook('../excel_template.xlsx')
    #print(wb.sheetnames)
    consumption_sheet = wb['HCONSUMO']
    consumption_sheet[CONST_CONSUMPTION_CELL].value = client.consumption
    print(consumption_sheet[CONST_CONSUMPTION_CELL].value)
    return wb

def copyBudgetArea(sheet):
    pass

def saveSheet(sheet, client):
    #wb.save('../test.xlsx')
    os.mkdir(('../{name}').format(name=client.fullName))
    sheet.save(('../{name}/GERADORES-{name}-ALDO-{version}-{time}.xlsx').format(name=client.fullName, version=CONST_SHEET_VERSION, time=datetime.time()))

def generateReport(client, sheet):
    pass

test_client = Client('Eduardo Moura Tavares', 500, 'ES', 'Praia de Itaparica-Vila Velha')
#print(vars(test_client))
wb = populateSheet(test_client)
print(wb)