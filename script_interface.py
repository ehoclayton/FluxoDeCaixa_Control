import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QLabel, QLineEdit, QPushButton, QComboBox, QMessageBox, QGridLayout
from PyQt5.QtGui import QFont

import openpyxl
from openpyxl.styles import Alignment
from datetime import datetime

class FluxoDeCaixaApp(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Registro de Fluxo de Caixa")
        self.setGeometry(100, 100, 400, 300)
        self.setStyleSheet("background: qradialgradient(cx:0.5, cy:0.5, fx:0.5, fy:0.5, radius: 1.5, stop:0 #000099, stop:1 #000066); color: white;")

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.central_widget.setStyleSheet("background: transparent; color: white;")

        # Crie um layout de grade para organizar os widgets
        grid = QGridLayout(self.central_widget)

        self.initUI(grid)

    def initUI(self, grid):
        font = QFont("Arial", 12)

        # Rótulos
        self.tipo_label = QLabel("Tipo de Transação:")
        self.tipo_label.setFont(font)

        self.descricao_label = QLabel("Descrição:")
        self.descricao_label.setFont(font)

        self.valor_label = QLabel("Valor:")
        self.valor_label.setFont(font)

        self.categoria_label = QLabel("Categoria:")
        self.categoria_label.setFont(font)

        # Entradas de texto
        self.descricao_entry = QLineEdit(self.central_widget)
        self.descricao_entry.setFont(font)

        self.valor_entry = QLineEdit(self.central_widget)
        self.valor_entry.setFont(font)

        # Dropdown para Tipo de Transação
        self.tipo_dropdown = QComboBox(self.central_widget)
        self.tipo_dropdown.addItems(["Entrada", "Saída"])
        self.tipo_dropdown.setFont(font)

        # Dropdown para Categoria
        self.categoria_dropdown = QComboBox(self.central_widget)
        self.categoria_dropdown.setFont(font)

        # Botão para adicionar nova categoria
        self.nova_categoria_entry = QLineEdit(self.central_widget)
        self.nova_categoria_entry.setFont(font)

        self.adicionar_categoria_button = QPushButton("Adicionar Categoria", self.central_widget)
        self.adicionar_categoria_button.setFont(font)
        self.adicionar_categoria_button.setStyleSheet(
            "background-color: #4CAF50; border: none; color: white; padding: 10px 20px; border-radius: 5px;"
        )
        self.adicionar_categoria_button.clicked.connect(self.adicionar_categoria)

        # Botão para remover categoria
        self.remover_categoria_button = QPushButton("Remover Categoria", self.central_widget)
        self.remover_categoria_button.setFont(font)
        self.remover_categoria_button.setStyleSheet(
            "background-color: #FF3333; border: none; color: white; padding: 10px 20px; border-radius: 5px;"
        )
        self.remover_categoria_button.clicked.connect(self.remover_categoria)

        # Rótulo de status
        self.status_label = QLabel("", self.central_widget)
        self.status_label.setFont(font)
        self.status_label.setStyleSheet("color: green;")

        # Botões
        self.registrar_button = QPushButton("Registrar", self.central_widget)
        self.registrar_button.setFont(font)
        self.registrar_button.setStyleSheet(
            "background-color: #2196F3; border: none; color: white; padding: 10px 20px; border-radius: 5px;"
        )
        self.registrar_button.clicked.connect(self.registrar_transacao)

        self.fechar_button = QPushButton("Fechar", self.central_widget)
        self.fechar_button.setFont(font)
        self.fechar_button.setStyleSheet(
            "background-color: #666; border: none; color: white; padding: 10px 20px; border-radius: 5px;"
        )
        self.fechar_button.clicked.connect(self.close)

        # Posicionamento dos widgets na grade
        grid.addWidget(self.tipo_label, 0, 0)
        grid.addWidget(self.tipo_dropdown, 0, 1)

        grid.addWidget(self.descricao_label, 1, 0)
        grid.addWidget(self.descricao_entry, 1, 1)

        grid.addWidget(self.valor_label, 2, 0)
        grid.addWidget(self.valor_entry, 2, 1)

        grid.addWidget(self.categoria_label, 3, 0)
        grid.addWidget(self.categoria_dropdown, 3, 1)

        grid.addWidget(self.nova_categoria_entry, 4, 1)
        grid.addWidget(self.adicionar_categoria_button, 4, 2)

        grid.addWidget(self.remover_categoria_button, 5, 0, 1, 2)

        grid.addWidget(self.status_label, 6, 0, 1, 2)

        grid.addWidget(self.registrar_button, 7, 0, 1, 2)
        grid.addWidget(self.fechar_button, 8, 0, 1, 2)

        # Carregar categorias
        self.carregar_categorias()

    def carregar_categorias(self):
        try:
            with open('categorias.txt', 'r') as file:
                categorias = file.read().splitlines()
                self.categoria_dropdown.addItems(categorias)
        except FileNotFoundError:
            return []

    def salvar_categorias(self):
        categorias_existentes = [self.categoria_dropdown.itemText(i) for i in range(self.categoria_dropdown.count())]
        with open('categorias.txt', 'w') as file:
            for categoria in categorias_existentes:
                file.write(categoria + '\n')

    def atualizar_categorias(self):
        self.categoria_dropdown.clear()
        self.carregar_categorias()

    def adicionar_categoria(self):
        nova_categoria = self.nova_categoria_entry.text()
        if nova_categoria and nova_categoria not in [self.categoria_dropdown.itemText(i) for i in range(self.categoria_dropdown.count())]:
            self.categoria_dropdown.addItem(nova_categoria)
            self.salvar_categorias()
            self.nova_categoria_entry.clear()
            QMessageBox.information(self, "Sucesso", "Categoria adicionada com sucesso!")

    def remover_categoria(self):
        categoria_selecionada = self.categoria_dropdown.currentText()
        if categoria_selecionada in categorias_existentes:
            categorias_existentes.remove(categoria_selecionada)
            self.salvar_categorias()
            self.atualizar_categorias()
            QMessageBox.information(self, "Sucesso", "Categoria removida com sucesso!")

    def registrar_transacao(self):
        tipo_transacao = self.tipo_dropdown.currentText()
        descricao = self.descricao_entry.text()
        valor_str = self.valor_entry.text().replace(",", ".")
        categoria = self.categoria_dropdown.currentText()

        valor = 0.0

        if not tipo_transacao or not descricao or not valor_str or not categoria:
            QMessageBox.critical(self, "Erro", "Preencha todos os campos antes de registrar!")
            return

        try:
            valor = float(valor_str)
        except ValueError:
            QMessageBox.critical(self, "Erro", "Valor inválido! Certifique-se de usar um número válido.")
            return

        if tipo_transacao == "Saída":
            valor = -valor

        try:
            workbook = openpyxl.load_workbook('fluxo_de_caixa.xlsx')
            sheet = workbook.active
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["Data", "Tipo", "Descrição", "Valor", "Categoria"])

        next_row = sheet.max_row + 1

        data_hora = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        sheet.cell(row=next_row, column=1, value=data_hora)
        sheet.cell(row=next_row, column=2, value=tipo_transacao)
        sheet.cell(row=next_row, column=3, value=descricao)
        sheet.cell(row=next_row, column=4, value=valor)
        sheet.cell(row=next_row, column=5, value=categoria)

        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=4, max_col=4):
            for cell in row:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal='right')

        workbook.save('fluxo_de_caixa.xlsx')
        QMessageBox.information(self, "Sucesso", "Transação registrada com sucesso!")

        self.descricao_entry.clear()
        self.valor_entry.clear()
        self.categoria_dropdown.setCurrentIndex(-1)

def main():
    app = QApplication(sys.argv)
    window = FluxoDeCaixaApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
