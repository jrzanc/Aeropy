import pandas as pd  
from cia_aerea.base import CiaAerea

class CiaAreaCancelsOtherPeriods(CiaAerea):
    def __init__(self,address):
        super().__init__(address)
        self.table=self.get_table() 


    def get_sheet_name(self):
        return "Bilhetes de Outros Períodos"
    
    def data_structure(self):
        keys_cols={"string":["Localizador",
                                "Bilhete", "Órgão Solicitante",
                                "Órgão Superior",
                                "PCDP" ,
                                "Evento(s) de Faturamento",
                                "Situação do Bilhete",
                                "Situações de Embarque",
                                "Status do Faturamento",
                                "Tipo(s) de Pendência",
                                "Análise da Fiscalização",
                                "Análise da Companhia Aérea",
                                "Relatório de Faturamento Definitivo da Emissão"],
                    "datetime64[ns]":["Data de emissão",
                                    "Data de Cancelamento"],
                    "float64":['Crédito de Reembolso',
                                'Desconto %',
                                'Desconto Efetivo %',
                                'Desconto Efetivo R$',
                                'Desconto R$',
                                'Glosa',
                                'Multa Calculada',
                                'Multa Efetiva',
                                'Multa de Reembolso Retornada',
                                'Reembolso Calculado',
                                'Reembolso Faturado',
                                'Reembolso Efetivo',
                                'Reembolso Retornado',
                                'Tarifa Comercial Retornada',
                                'Tarifa Embarque Efetiva',
                                'Tarifa Praticada Calculada',
                                'Tarifa Praticada Efetiva',
                                'Tarifa Praticada Retornada',
                                'Tarifa de Embarque Retornada',
                                'Total',
                                'Valor a Faturar']
                    }
        for key in keys_cols.keys():
            self._loc_teste("table",keys_cols[key],key)

    def flow(self):
        self.data_structure()
        self.insert_message_on_tickets_canceled_after_the_billing_period()
        self.change_worksheet()