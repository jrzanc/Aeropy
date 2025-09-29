
import pandas as pd
from openpyxl import load_workbook
from tcrp.cia_aerea.base import CiaAerea


# Classe Cia Aérea Emitidos
class CiaAreaIssued(CiaAerea):
    def __init__(self,address):
        super().__init__(address)
        self.table=self.get_table()   

    def get_sheet_name(self):
        return "Relatório Definitivo"

 

        
    def flow(self):
        self.data_structure()
        self.insert_message_on_tickets_canceled_after_the_billing_period()
        self.insert_message_on_tickets_with_discount_bellow_three_percent()
        self.insert_message_about_72_hours()
        self.general_message()
        self.change_worksheet()