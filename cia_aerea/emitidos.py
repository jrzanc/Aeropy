
import pandas as pd
from openpyxl import load_workbook
from cia_aerea.base import CiaAerea


# Classe Cia Aérea Emitidos
class CiaAreaIssued(CiaAerea):
    def __init__(self,address):
        super().__init__(address)
        self.table=self.get_table()   

    def get_sheet_name(self):
        return "Emitidos"

    def insert_message_on_tickets_canceled_after_the_billing_period(self):
        df = getattr(self, "table") 
        #tbl=self.tickets_issued.copy()
        filter_billing_period=df[(~df['Data de Cancelamento'].isnull()
                                )&(df['Data de Cancelamento']>=self.metadata.end_billing_period)
                                ]
        new_text="Cancelado no mês seguinte ao de referência. Será reembolsado no próximo faturamento."
        
        df.loc[filter_billing_period.index, "Análise da Fiscalização"] = (df.loc[filter_billing_period.index, "Análise da Fiscalização"]
                                                                           .apply(lambda x: new_text if pd.isna(x) else f"{x}\n{new_text}"
                                                                                 )
                                                                          )
        setattr(self,"table",df)

        
    def flow(self):
        self.data_structure()
        self.insert_message_on_tickets_canceled_after_the_billing_period()
        self.insert_message_on_tickets_with_discount_bellow_three_percent()
        self.insert_message_about_72_hours()
        self.general_message()
        self.change_worksheet()