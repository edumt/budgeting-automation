import os
import datetime

# from openpyxl import Workbook
from openpyxl import load_workbook
from docx import Document

SHEET_VERSION = "13-JULHO-2021"
CONSUMPTION_TABLE = "HCONSUMO"
INVERTER_PANEL = "Preço SFCR-GROWATT PHONO 450Wp"
NEAREST_PANELS = "=+VLOOKUP(MIN('GROWATT-PHONO 450Wp'!D:D),'GROWATT-PHONO 450Wp'!D:E,2,0)"

# sheet cells constants
CONSUMPTION = "D18"
NAME = "C4"
PRICE_ADJUSTMENT = "H44"  # TODO: dinamically adjust price
OPTION_1 = "D11"
OPTION_2 = "F11"
OPTION_3 = "H11"
QUOTE_AREA = "B3:H32"

statesDict = {
    "ES": "ESPÍRITO SANTO",
    "RJ": "RIO DE JANEIRO",
    "BA": "BAHIA",
    "MG": "MINAS GERAIS",
}


def getFormulaSheet():
    return load_workbook("../excel_template.xlsx")


def getDocxTemplate():
    return Document("../word_template.docx")


def getCurrentDate():
    return datetime.date.today().strftime("%d/%m/%Y")


# TO DO: change some methods to functions
class Client:
    def __init__(self, name, consumption, state, city=""):
        self.name = name.upper()
        self.consumption = consumption  # Average monthly energy consumption [kWh/month]
        self.state = state.upper()  # State initials, e.g. ES
        self.city = city.upper()  # (District-)City, e.g. PRAIA DE ITAPARICA-VILA VELHA
        self.reference = 1234  # TO DO: get highest reference number on quotes dir and add 1 -> getReferenceNumber()

    def populateSheet(self):
        self.sheet[CONSUMPTION_TABLE][CONSUMPTION].value = self.consumption
        self.sheet[CONSUMPTION_TABLE][NAME].value = "CLIENTE: {name}".format(name=self.name)
        self.setPanelsQuantity()

    def setPanelsQuantity(self):
        if self.consumption <= 550:
            self.sheet[INVERTER_PANEL][OPTION_2].value = NEAREST_PANELS + "+1"
            self.sheet[INVERTER_PANEL][OPTION_3].value = NEAREST_PANELS + "+2"

        elif self.consumption <= 750:
            self.sheet[INVERTER_PANEL][OPTION_1].value = NEAREST_PANELS + "-1"
            self.sheet[INVERTER_PANEL][OPTION_2].value = NEAREST_PANELS
            self.sheet[INVERTER_PANEL][OPTION_3].value = NEAREST_PANELS + "+1"

        elif self.consumption <= 1300:
            self.sheet[INVERTER_PANEL][OPTION_1].value = NEAREST_PANELS + "-2"
            self.sheet[INVERTER_PANEL][OPTION_2].value = NEAREST_PANELS + "-1"
            self.sheet[INVERTER_PANEL][OPTION_3].value = NEAREST_PANELS

        else:
            self.sheet[INVERTER_PANEL][OPTION_1].value = NEAREST_PANELS + "-4"
            self.sheet[INVERTER_PANEL][OPTION_2].value = NEAREST_PANELS + "-2"
            self.sheet[INVERTER_PANEL][OPTION_3].value = NEAREST_PANELS

    def saveSheet(self):
        # TO DO: learn mkdir() and/or save() best practices, what to do when dir/file aready exists, exception handling etc
        # os.mkdir(('../{name}').format(name=client.name))
        # self.sheet.save(('../{name}/GERADORES-{name}-ALDO-{version}.xlsx').format(name=self.name, version=SHEET_VERSION))
        # self.sheet.active = self.sheet[INVERTER_PANEL]
        self.sheet.save("../test.xlsx")

    def getDataSheet(self):
        self.sheet = load_workbook("../test.xlsx", data_only=True)

    def generateSheet(self):
        self.sheet = getFormulaSheet()
        self.populateSheet()
        self.saveSheet()
        print("Success! Sheet generated.")

    def docxSearchAndReplace(self, oldText, newText):
        # reference: https://stackoverflow.com/questions/24805671/how-to-use-python-docx-to-replace-text-in-a-word-document-and-save
        # TO DO: learn about args and kwargs

        # TO DO: function not working properly, missing some replaces
        for p in self.docx.paragraphs:
            if oldText in p.text:
                for r in p.runs:
                    if oldText == r.text:
                        r.text = newText

    def populateDocx(self):
        self.docxSearchAndReplace("fullName", self.name)
        self.docxSearchAndReplace("reference", str(self.reference))  # not replacing all reference tags
        self.docxSearchAndReplace("date", getCurrentDate())
        if self.city == "":
            self.docxSearchAndReplace("location", statesDict[self.state])
        else:
            self.docxSearchAndReplace(
                "location",
                (self.city + "/" + self.state),
            )
        self.docxSearchAndReplace(
            "consumption", str(self.consumption)
        )  # https://stackoverflow.com/questions/1823058/how-to-print-number-with-commas-as-thousands-separators

    def copyBudgetArea(self):
        # TODO: try to copy and paste excel to word
        # if not possible, try to copy excel to clipboard and manually paste special
        # if not possible, probably give up lol
        self.getDataSheet()

    def saveDocx(self):
        self.docx.save("../test.docx")

    def generateQuote(self):
        self.docx = getDocxTemplate()
        self.populateDocx()
        self.saveDocx()
        print("Success! Quote generated.")


test_client = Client(
    "Eduardo Moura Tavares",
    500,
    "ES",
    "Praia de Itaparica-Vila Velha",
)
# test_client = Client('Eduardo Moura Tavares', 5000, 'mg')
test_client.generateSheet()
# test_client.generateQuote()
os.startfile("F:/Google Drive/Projetos/test.xlsx")  # for testing purposes
# os.startfile('F:/Google Drive/Projetos/test.docx') #for testing purposes


""" 
if __name__ == "__main__":
    main()
    
    
    TO DO:  CLI
            GUI
            
    MAYBE TO DO:
                e-mail integration
                whatsapp integration
"""
