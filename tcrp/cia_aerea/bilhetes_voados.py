import pandas as pd  
from cia_aerea.base import CiaAerea
class CiaAreaCancelsOtherPeriods(CiaAerea):
    def __init__(self,address):
        super().__init__(address)
        self.table=self.get_table() 

    def get_sheet_name(self):
        return None