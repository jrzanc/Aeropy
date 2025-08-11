
#from planilhas.metadata import SpreadsheetMetadata

from gui.tel_inicial import TelaInicial,QApplication
import sys




# Método Main

if __name__ == "__main__":
    app = QApplication.instance()
    if app is None:
        app = QApplication(sys.argv)
    janela = TelaInicial()
    janela.show()
    sys.exit(app.exec())




    '''
    while True:
        # Instancia a classe Planilha e atribui à variável worksheet_excel
        worksheet_excel_planilha_medicao_provisoria= Planilha()
        # Chama .flow() para obter o endereço da planilha copiada com novo nome
        address_planilha_medicao_provisoria = worksheet_excel_planilha_medicao_provisoria.flow("Selecione Planilha de Medição Provisória")    
        worksheet_excel_bilhetes_voados = Planilha()
        address_bilhetes_voados  = worksheet_excel_bilhetes_voados.flow("Selecione Planilha de Bilhetes Voados",copy=False) 



        # Cria uma instância de CiaAreaIssued usando o endereço da planilha copiada   
        issued = CiaAreaIssued(address_planilha_medicao_provisoria)
        # Chama .flow() para tratar a planilha de bilhetes emitidos
        issued.flow()

        # Cria uma instância de CiaAreaIssuedCancelled usando o endereço da planilha copiada  
        issued_cancelled = CiaAreaIssuedCancelled(address_planilha_medicao_provisoria)
        # Chama .flow() para tratar a planilha de bilhetes emitidos
        issued_cancelled.flow()

        #Cria uma instância de CiaAreaCancelsOtherPeriods usando o endereço da planilha copiada  
        issued_cancels_other_periods = CiaAreaCancelsOtherPeriods(address_planilha_medicao_provisoria,address_bilhetes_voados)
        # Chama .flow() para tratar a planilha de bilhetes cancelados em outroas periodos
        issued_cancels_other_periods.flow()

        cont=input("Deseja continuar? (0 - Não, 1 - Sim): ")
        if cont=="0":
            break'''