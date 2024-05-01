from PyQt6.QtCore import Qt
import sys 
from PyQt6.QtWidgets import QApplication, QWidget
class VentanaVacia(QWidget):
    def __init__(self):
        super().__init__()
        self.iniciarUI()

    def iniciarUI(self):
        self.setGeometry(100,100,250,250)
        self.setWindowTitle("titulo")
        self.show()


if __name__=='__main__':
    app= QApplication(sys.argv)
    ventana=VentanaVacia()
    sys.exit(app.exec())            
    