import sys
from PySide6.QtWidgets import QApplication
from gui import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)

    janela = MainWindow()
    janela.show()

    sys.exit(app.exec())