# Importing Labraries
from tcrp.cia_aerea.emitidos import CiaAreaIssued
from tcrp.cia_aerea.emitidos_cancelados import CiaAreaIssuedCancelled
from tcrp.cia_aerea.bilhetes_outros_periodos import CiaAreaCancelsOtherPeriods
#from cia_aerea.base import CiaAerea
from planilhas.manipulados import Planilha

from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout,
    QHBoxLayout, QGroupBox, QMessageBox, QRadioButton, QButtonGroup
)
from PySide6.QtGui import QPixmap
from PySide6.QtCore import Qt
import os

# Classe Visual
class TelaInicial(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Aeropy - Central de Compras")
        self.setGeometry(100, 100, 300, 200)
        self.setStyleSheet("background-color: #007b8a; color: white; font-family: Arial;")

        # Logo JPG à esquerda
        logo = QLabel()
        caminho_logo = "logo.jpg"
        pixmap = QPixmap(caminho_logo)
        pixmap = pixmap.scaled(100, 100, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        logo.setPixmap(pixmap)
        logo.setFixedSize(100, 100)

        # Título e subtítulo à direita da logo
        titulo = QLabel("Bem-vindo ao Aeropy")
        titulo.setStyleSheet("font-size: 20px; font-weight: bold;")
        subtitulo = QLabel("Criado por Junior A. Zancanaro")
        subtitulo.setStyleSheet("font-size: 12px;")

        titulo_layout = QVBoxLayout()
        titulo_layout.addWidget(titulo, alignment=Qt.AlignCenter)
        titulo_layout.addWidget(subtitulo, alignment=Qt.AlignCenter)

        # Layout horizontal para logo + título
        topo_layout = QHBoxLayout()
        topo_layout.addWidget(logo)
        topo_layout.addLayout(titulo_layout)

        # Instrução
        instrucao = QLabel("Selecione qual parte do processo de medição você está executando:")
        instrucao.setAlignment(Qt.AlignCenter)

        # Radio buttons
        self.radio_group = QButtonGroup(self)
        self.radio_provisoria = QRadioButton("Etapa de medição provisória")
        self.radio_definitiva = QRadioButton("Etapa de medição definitiva")
        for rb in [self.radio_provisoria, self.radio_definitiva]:
            rb.setStyleSheet("font-size: 16px;")
            self.radio_group.addButton(rb)

        # Agrupando os radio buttons
        box = QGroupBox("Etapas disponíveis")
        box.setStyleSheet("QGroupBox { font-size: 18px; font-weight: bold; }")
        box_layout = QVBoxLayout()
        box_layout.addWidget(self.radio_provisoria)
        box_layout.addWidget(self.radio_definitiva)
        box.setLayout(box_layout)

        # Botões
        btn_continuar = QPushButton("Continuar")
        btn_cancelar = QPushButton("Cancelar")
        btn_continuar.clicked.connect(self.continuar)
        btn_cancelar.clicked.connect(self.fechar_aplicacao)

        for btn in [btn_continuar, btn_cancelar]:
            btn.setStyleSheet("background-color: white; color: #007b8a; font-weight: bold; padding: 8px;")

        # Layout principal
        layout = QVBoxLayout()
        layout.addLayout(topo_layout)
        layout.addWidget(instrucao)
        layout.addWidget(box)

        botoes_layout = QHBoxLayout()
        botoes_layout.addWidget(btn_continuar)
        botoes_layout.addWidget(btn_cancelar)
        layout.addLayout(botoes_layout)

        self.setLayout(layout)
        
        # Centraliza a janela na tela
        screen = QApplication.primaryScreen()
        screen_geometry = screen.availableGeometry()
        x = (screen_geometry.width() - self.width()) // 2
        y = (screen_geometry.height() - self.height()) // 2
        self.move(x, y)


    def continuar(self):
        if self.radio_provisoria.isChecked():
            try:
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
                QMessageBox.information(self, "Sucesso", "Processo executado com sucesso!")
            except:
                QMessageBox.critical(self, "Erro", "Erro ao manusear a planilha!")
            finally:
                self.close()
                QApplication.instance().quit()

        elif self.radio_definitiva.isChecked():
            QMessageBox.information(self, "Continuar", "Iniciando a etapa: Etapa de medição definitiva")

        else:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione uma etapa antes de continuar.")

    def fechar_aplicacao(self):
        self.close()
        QApplication.instance().quit()