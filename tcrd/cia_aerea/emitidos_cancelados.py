import pandas as pd  
from tcrp.cia_aerea.base import CiaAerea


# Classe emitidos e Cancelados  
class CiaAreaIssuedCancelled(CiaAerea):
    def __init__(self,address):
        super().__init__(address)
        self.table=self.get_table()   

    def get_sheet_name(self):
        return "Emitidos e Cancelados"
        
    def void_tickets(self):
        df = getattr(self, "table") 
        df.loc[df["Status do Faturamento"] == "Não Faturável", "Análise da Fiscalização"] = "VOID"
        setattr(self,"table",df)

    
    def insert_message_on_tickets_with_discount_bellow_three_percent(self):
        df = getattr(self, "table") 
        filter_discount_below_three_percent= df[df["Tipo(s) de Pendência"].str.contains("Pendente de Desconto", na=False)]
        new_text="Validar o valor do reembolso indicado na coluna 'Reemboso Efetivo' - Não foi aplicado o desconto - bilhete glosado"
        df.loc[filter_discount_below_three_percent.index,"Análise da Fiscalização"]=(df.loc[filter_discount_below_three_percent.index,"Análise da Fiscalização"]
                                                                                        .apply(lambda x: new_text if pd.isna(x) else f"{x}\n{new_text}"
                                                                                                                                                                        )
                                                                                        )
        setattr(self,"table",df)   


     
    def insert_returned_tickets_above_threshold(self):
        """
        Este método compara o reembolso retornado pela Cia com o calculado pelo SCDP.
        Se o valor retornado for maior que o efetivo, adiciona uma observação na coluna
        'Análise da Fiscalização' indicando a necessidade de validação.
        """
        df = getattr(self, "table")

        def append_analysis_returned_tickets_above_threshold(row):
            
            if row['Reembolso Retornado'] > row['Reembolso Efetivo']:
                valor_retornado = row['Reembolso Retornado']
                texto_adicional = (
                    "Validar o valor do reembolso indicado na coluna 'Reemboso Efetivo' - "
                    f"Reembolso retornado pela cia. Valor retornado: R$ {valor_retornado:.2f}"
                )
                if pd.isna(row['Análise da Fiscalização']):
                    return texto_adicional
                else:
                    return f"{row['Análise da Fiscalização']}\n{texto_adicional}"
            return row['Análise da Fiscalização']

        df['Análise da Fiscalização'] = df.apply(append_analysis_returned_tickets_above_threshold, axis=1)
        setattr(self,"table",df) 

 

    def general_message(self):
        df = getattr(self, "table") 
        df.loc[df["Análise da Fiscalização"].isna(),"Análise da Fiscalização"]="Validar o valor do reembolso indicado na coluna 'Reemboso Efetivo'"
        setattr(self,"table",df)    
    
    
    def flow(self):
        self.data_structure()        
        self.insert_message_on_tickets_with_discount_bellow_three_percent()
        self.insert_message_about_72_hours()
        self.insert_returned_tickets_above_threshold()
        self.general_message()
        self.void_tickets()
        self.change_worksheet()