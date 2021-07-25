import os
import datetime

# from openpyxl import Workbook
from openpyxl import load_workbook
from docx import Document

SHEET_VERSION = "13-JULHO-2021"
CONSUMPTION = "D18"
NAME = "C4"
PRICE_ADJUSTMENT = "H44"  # TODO: dinamically adjust price
OPTION_1 = "D11"
OPTION_2 = "F11"
OPTION_3 = "H11"
QUOTE_AREA = "B3:H32"


statesDict = {
    # TO DO: learn about enums and determine if they make more sense than a dictionary for this
    "ES": "ESPÍRITO SANTO",
    "RJ": "RIO DE JANEIRO",
    "BA": "BAHIA",
    "MG": "MINAS GERAIS",
}

# TO DO: change some methods to functions


class Client:
    def __init__(self, name, consumption, state, city=""):
        self.name = name.upper()
        self.consumption = consumption  # Average monthly energy consumption [kWh/month]
        self.state = state.upper()  # State initials, e.g. ES
        self.city = city.upper()  # (District-)City, e.g. PRAIA DE ITAPARICA-VILA VELHA
        self.reference = (
            1234  # TO DO: get highest reference number on quotes directory and add 1 -> getReferenceNumber()
        )

    def getFormulaSheet(self):
        self.sheet = load_workbook("../excel_template.xlsx")

    def populateSheet(self):
        # print(sheet.sheetnames)
        self.sheet["HCONSUMO"][CONSUMPTION].value = self.consumption
        self.sheet["HCONSUMO"][NAME].value = ("CLIENTE: {name}").format(name=self.name)
        self.setPanelsQuantity()

    def setPanelsQuantity(self):
        # TODO: nearest panel number const, removing unnecessary cell accessing
        if self.consumption <= 550:
            # set option 2
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_2].value = (
                self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_1].value + "+1"
            )
            # set option 3
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_3].value = (
                self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_2].value + "+1"
            )
        elif self.consumption <= 750:
            # set option 2
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_2].value = self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][
                OPTION_1
            ].value
            # set option 1
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_1].value = (
                self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_2].value + "-1"
            )
            # set option 3
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_3].value = (
                self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_2].value + "+1"
            )
        elif self.consumption <= 1300:
            # set option 3
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_3].value = self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][
                OPTION_1
            ].value
            # set option 2
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_2].value = (
                self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_3].value + "-1"
            )
            # set option 1
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_1].value = (
                self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_2].value + "-1"
            )
        else:
            # set option 3
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_3].value = self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][
                OPTION_1
            ].value
            # set option 2
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_2].value = (
                self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_3].value + "-2"
            )
            # set option 1
            self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_1].value = (
                self.sheet["Preço SFCR-GROWATT PHONO 450Wp"][OPTION_2].value + "-2"
            )

    def saveSheet(self):
        # TO DO: learn mkdir() and/or save() best practices, what to do when dir/file aready exists, exception handling etc
        # os.mkdir(('../{name}').format(name=client.name))
        # self.sheet.save(('../{name}/GERADORES-{name}-ALDO-{version}.xlsx').format(name=self.name, version=SHEET_VERSION))
        self.sheet.active = self.sheet["Preço SFCR-GROWATT PHONO 450Wp"]
        self.sheet.save("../test.xlsx")

    def getDataSheet(self):
        self.sheet = load_workbook("../test.xlsx", data_only=True)

    def generateSheet(self):
        self.getFormulaSheet()
        self.populateSheet()
        self.saveSheet()
        print("Success! Sheet generated.")

    def getDocxTemplate(self):
        self.docx = Document("../word_template.docx")

    def docxSearchAndReplace(self, oldText, newText):
        # reference: https://stackoverflow.com/questions/24805671/how-to-use-python-docx-to-replace-text-in-a-word-document-and-save
        # TO DO: learn about args and kwargs

        # TO DO: function not working properly, missing come replaces
        for p in self.docx.paragraphs:
            if oldText in p.text:
                for r in p.runs:
                    if oldText == r.text:
                        r.text = newText

    def getCurrentDate(self):
        today = datetime.date.today()
        return today.strftime("%d/%m/%Y")

    def populateDocx(self):
        self.docxSearchAndReplace("fullName", self.name)
        self.docxSearchAndReplace("reference", str(self.reference))  # not replacing all reference tags
        self.docxSearchAndReplace("date", self.getCurrentDate())
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
        self.getDocxTemplate()
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


""" TO DO:  CLI
            GUI
            
    MAYBE TO DO:
                e-mail integration
                whatsapp integration
"""
