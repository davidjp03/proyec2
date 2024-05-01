import sys
import test
from PyQt6 import QtCore, QtGui, QtWidgets 
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QApplication, QLabel,QWidget,QLineEdit,
 QPushButton, QMessageBox,QCheckBox, QVBoxLayout)
from PyQt6.QtGui import QFont, QPixmap
from main import MainWindow 
from Estudiante import Estudiante




usuarios = {
    "usuario1": Estudiante("david jimenez","1193285812", ["ing. de sistemas"], 3.5, 4, "david.jimenez11@eia.edu.co", "abc123"),
    "usuario2": Estudiante("paulina medina","123456789", ["ing. biomedica","ing. de sistemas"], 4.5, 4, "paulina.medina@eia.edu.co", "123abc"),
    "usuario3": Estudiante("isaac","987654321", ["ing. administrativa","ing. de sistemas"], 4.4, 4, "isaac.giraldo@eia.edu.co", "321bca")
}

class Login(QWidget):
    def __init__(self):
        super().__init__()

        self.setGeometry(100,100,550,450)
        self.setWindowTitle("Iniciar Sesión con cuenta EIA")
        self.generarFormulario()
        
        self.show()

    def inicioSesion(self):
        email=self.user_input.text()
        password = self.password_input.text()
        for usuario in usuarios.values():
            if usuario.correo == email and usuario.contrasena == password:
                QMessageBox.information(self, "Inicio de sesión exitoso", f"Bienvenido(a), {usuario.nombre}!")
                self.iniciar_mainview(usuario)
                return

        QMessageBox.warning(self, "Inicio de sesión fallido", "Correo electrónico o contraseña incorrectos.")

    def generarFormulario(self):
        self.label = QtWidgets.QLabel(self)
        self.label.setGeometry(QtCore.QRect(30, 30, 300, 420))
        self.imagen = QLabel(self)

    # Cargar la imagen en un QPixmap y establecerla en el QLabel
        pixmap = QPixmap("C:\\Users\\david\\Desktop\\pro\\login\\logo.png")
        self.imagen.setPixmap(pixmap)
        self.imagen.setFixedSize(250, 145)
        self.imagen.move(152, 10)
    # Permitir que la imagen se escale si el widget cambia de tamaño
        self.imagen.setScaledContents(True)
        self.label.setText("")
       
        
        self.user_input=QLineEdit(self)
        self.user_input.resize(250,24)
        self.user_input.move(152,185)
        self.user_input.setStyleSheet("background-color:rgba(0, 0, 0, 0);\n"
"border:none;\n"
"border-bottom:2px solid rgba(105, 118, 132, 255);\n"
"padding-bottom:7px;")
        self.user_input.setPlaceholderText("Usuario")
        self.password_input=QLineEdit(self)
        self.password_input.resize(250,24)
        self.password_input.move(152,230)
        self.password_input.setStyleSheet("background-color:rgba(0, 0, 0, 0);\n"
"border:none;\n"
"border-bottom:2px solid rgba(105, 118, 132, 255);\n"
"padding-bottom:7px;")
        self.password_input.setPlaceholderText("Contraseña")
        

        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        login_button = QtWidgets.QPushButton(self)
        login_button.setText('Ingresar')
        login_button.resize(250, 35)
        login_button.move(152,280)
        login_button.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        login_button.clicked.connect(self.inicioSesion)
        

    def iniciar_mainview(self,usuario):
        self.main_window=MainWindow(usuario)
        self.main_window.show()
        self.close()
    
    
if __name__=='__main__':
    app= QApplication(sys.argv)
    login= Login()
    sys.exit(app.exec())          
                 