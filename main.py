
#from planilhas.metadata import SpreadsheetMetadata

from cia_aerea.emitidos import CiaAreaIssued
from cia_aerea.emitidos_cancelados import CiaAreaIssuedCancelled
from cia_aerea.bilhetes_outros_periodos import CiaAreaCancelsOtherPeriods
#from cia_aerea.base import CiaAerea
from planilhas.manipulados import planilha

#from planilhas.metadata import SpreadsheetMetadata




# Método Main
if __name__ == "__main__": 
    while True:
        # Instancia a classe Planilha e atribui à variável worksheet_excel
        worksheet_excel = planilha()
        # Chama .flow() para obter o endereço da planilha copiada com novo nome
        address = worksheet_excel.flow()    

        # Cria uma instância de CiaAreaIssued usando o endereço da planilha copiada   
        issued = CiaAreaIssued(address)
        # Chama .flow() para tratar a planilha de bilhetes emitidos
        issued.flow()

        # Cria uma instância de CiaAreaIssuedCancelled usando o endereço da planilha copiada  
        issued_cancelled = CiaAreaIssuedCancelled(address)
        # Chama .flow() para tratar a planilha de bilhetes emitidos
        issued_cancelled.flow()

        # Cria uma instância de CiaAreaCancelsOtherPeriods usando o endereço da planilha copiada  
        #issued_cancels_other_periods = CiaAreaCancelsOtherPeriods(address)
        # Chama .flow() para tratar a planilha de bilhetes cancelados em outroas periodos
        #issued_cancels_other_periods.flow()

        cont=input("Deseja continuar? (0 - Não, 1 - Sim): ")
        if cont=="0":
            break