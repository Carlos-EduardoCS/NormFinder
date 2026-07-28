import os

from PySide6.QtWidgets import (
    QMainWindow,
    QWidget,
    QPushButton,
    QLabel,
    QFileDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QTableWidget,
    QTableWidgetItem,
    QMessageBox
)

from scanner import Scanner
from export_excel import ExportExcel


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.scanner = Scanner()
        self.exportador = ExportExcel()

        self.resultados = []

        self.pasta = ""

        self.setWindowTitle("NormFinder")
        self.resize(1100, 650)

        central = QWidget()
        self.setCentralWidget(central)

        layout = QVBoxLayout()

        # ==========================
        # Seleção da pasta
        # ==========================

        linha1 = QHBoxLayout()

        self.bt_pasta = QPushButton("Selecionar Pasta")
        self.bt_pasta.clicked.connect(self.selecionar_pasta)

        self.lb_pasta = QLabel("Nenhuma pasta selecionada")

        linha1.addWidget(self.bt_pasta)
        linha1.addWidget(self.lb_pasta)

        layout.addLayout(linha1)

        # ==========================
        # Pesquisa
        # ==========================

        linha2 = QHBoxLayout()

        self.txt_norma = QLineEdit()

        self.txt_norma.setPlaceholderText(
            "Digite uma norma (Ex.: IEC 60335-1)"
        )

        self.bt_buscar = QPushButton("Pesquisar")

        self.bt_buscar.clicked.connect(self.buscar)

        linha2.addWidget(self.txt_norma)
        linha2.addWidget(self.bt_buscar)

        layout.addLayout(linha2)

        # ==========================
        # Tabela
        # ==========================

        self.tabela = QTableWidget()

        self.tabela.setColumnCount(5)

        self.tabela.setHorizontalHeaderLabels([
            "Arquivo",
            "Norma",
            "Ocorrências",
            "Trecho",
            "Caminho"
        ])

        layout.addWidget(self.tabela)

        # ==========================
        # Exportar
        # ==========================

        self.bt_excel = QPushButton(
            "Exportar para Excel"
        )

        self.bt_excel.clicked.connect(
            self.exportar_excel
        )

        layout.addWidget(self.bt_excel)

        central.setLayout(layout)

    # ======================================

    def selecionar_pasta(self):

        pasta = QFileDialog.getExistingDirectory(
            self,
            "Escolha uma pasta"
        )

        if pasta:

            self.pasta = pasta

            self.lb_pasta.setText(pasta)

    # ======================================

    def buscar(self):

        if self.pasta == "":

            QMessageBox.warning(
                self,
                "Aviso",
                "Escolha uma pasta."
            )

            return

        norma = self.txt_norma.text().strip()

        if norma == "":

            QMessageBox.warning(
                self,
                "Aviso",
                "Digite uma norma."
            )

            return

        self.resultados = self.scanner.pesquisar(
            self.pasta,
            norma
        )

        self.preencher_tabela()

    # ======================================

    def preencher_tabela(self):

        self.tabela.setRowCount(0)

        for resultado in self.resultados:

            for trecho in resultado["trechos"]:

                linha = self.tabela.rowCount()

                self.tabela.insertRow(linha)

                self.tabela.setItem(
                    linha,
                    0,
                    QTableWidgetItem(
                        resultado["arquivo"]
                    )
                )

                self.tabela.setItem(
                    linha,
                    1,
                    QTableWidgetItem(
                        resultado["norma"]
                    )
                )

                self.tabela.setItem(
                    linha,
                    2,
                    QTableWidgetItem(
                        str(resultado["quantidade"])
                    )
                )

                self.tabela.setItem(
                    linha,
                    3,
                    QTableWidgetItem(
                        trecho
                    )
                )

                self.tabela.setItem(
                    linha,
                    4,
                    QTableWidgetItem(
                        resultado["caminho"]
                    )
                )

        self.tabela.resizeColumnsToContents()

        QMessageBox.information(
            self,
            "Pesquisa Finalizada",
            f"Foram encontrados {len(self.resultados)} arquivo(s)."
        )

    # ======================================

    def exportar_excel(self):

        if len(self.resultados) == 0:

            QMessageBox.warning(
                self,
                "Aviso",
                "Nenhum resultado para exportar."
            )

            return

        arquivo, _ = QFileDialog.getSaveFileName(

            self,

            "Salvar",

            "Relatorio.xlsx",

            "Excel (*.xlsx)"

        )

        if arquivo == "":
            return

        self.exportador.exportar(
            self.resultados,
            arquivo
        )

        QMessageBox.information(

            self,

            "Concluído",

            "Relatório exportado com sucesso."

        )