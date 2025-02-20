import sys
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QMessageBox
from PyQt6.QtCore import Qt
from openpyxl import Workbook
from datetime import datetime, timedelta

class BancoHorasApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Calculadora de Banco de Horas')
        self.setGeometry(100, 100, 600, 400)

        layout = QVBoxLayout()

        # Nome do Trabalhador
        self.nome_label = QLabel('Nome do Trabalhador:')
        self.nome_input = QLineEdit()
        layout.addWidget(self.nome_label)
        layout.addWidget(self.nome_input)

        # Setor
        self.setor_label = QLabel('Setor:')
        self.setor_input = QLineEdit()
        layout.addWidget(self.setor_label)
        layout.addWidget(self.setor_input)

        # Função
        self.funcao_label = QLabel('Função:')
        self.funcao_input = QLineEdit()
        layout.addWidget(self.funcao_label)
        layout.addWidget(self.funcao_input)

        # Horas Trabalhadas
        self.horas_label = QLabel('Horas Trabalhadas (HH:MM):')
        self.horas_input = QLineEdit()
        layout.addWidget(self.horas_label)
        layout.addWidget(self.horas_input)

        # Botão para adicionar horas
        self.adicionar_button = QPushButton('Adicionar Horas')
        self.adicionar_button.clicked.connect(self.adicionar_horas)
        layout.addWidget(self.adicionar_button)

        # Tabela para exibir as horas trabalhadas
        self.tabela = QTableWidget()
        self.tabela.setColumnCount(4)
        self.tabela.setHorizontalHeaderLabels(['Nome', 'Setor', 'Função', 'Horas Trabalhadas'])
        layout.addWidget(self.tabela)

        # Botão para exportar para Excel
        self.exportar_button = QPushButton('Exportar para Excel')
        self.exportar_button.clicked.connect(self.exportar_excel)
        layout.addWidget(self.exportar_button)

        self.setLayout(layout)

    def adicionar_horas(self):
        nome = self.nome_input.text()
        setor = self.setor_input.text()
        funcao = self.funcao_input.text()
        horas = self.horas_input.text()

        if not nome or not setor or not funcao or not horas:
            QMessageBox.warning(self, 'Erro', 'Todos os campos devem ser preenchidos!')
            return

        try:
            horas_trabalhadas = datetime.strptime(horas, '%H:%M').time()
        except ValueError:
            QMessageBox.warning(self, 'Erro', 'Formato de horas inválido! Use HH:MM.')
            return

        row_position = self.tabela.rowCount()
        self.tabela.insertRow(row_position)
        self.tabela.setItem(row_position, 0, QTableWidgetItem(nome))
        self.tabela.setItem(row_position, 1, QTableWidgetItem(setor))
        self.tabela.setItem(row_position, 2, QTableWidgetItem(funcao))
        self.tabela.setItem(row_position, 3, QTableWidgetItem(horas))

        self.nome_input.clear()
        self.setor_input.clear()
        self.funcao_input.clear()
        self.horas_input.clear()

    def exportar_excel(self):
        if self.tabela.rowCount() == 0:
            QMessageBox.warning(self, 'Erro', 'Nenhum dado para exportar!')
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Banco de Horas"

        # Adicionar cabeçalhos
        headers = ['Nome', 'Setor', 'Função', 'Horas Trabalhadas']
        ws.append(headers)

        # Adicionar dados
        for row in range(self.tabela.rowCount()):
            row_data = []
            for column in range(self.tabela.columnCount()):
                item = self.tabela.item(row, column)
                row_data.append(item.text())
            ws.append(row_data)

        # Salvar o arquivo
        filename = 'banco_de_horas.xlsx'
        wb.save(filename)
        QMessageBox.information(self, 'Sucesso', f'Dados exportados para {filename}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = BancoHorasApp()
    ex.show()
    sys.exit(app.exec())