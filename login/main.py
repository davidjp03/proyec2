import sys
from PyQt6.QtCore import Qt
from PyQt6.QtWidgets import (QApplication, QLabel,QWidget,QLineEdit,
 QPushButton, QMessageBox,QCheckBox,QVBoxLayout)
from PyQt6.QtGui import QFont, QPixmap
from PyQt5 import QtCore
from PyQt5.QtCore import QPropertyAnimation
from PyQt5 import QtCore, QtGui, QtWidgets

from Estudiante import Estudiante 
import sys
import smtplib
import getpass
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

import os
import tkinter as tk
from tkinter import ttk
import pandas as pd
from tabulate import tabulate

# Función para leer los horarios de los profesores desde archivos Excel
def leer_horarios_profesores(ruta_directorio):
    horarios_profesores = {}

    # Obtener la lista de archivos en el directorio
    archivos_excel = [archivo for archivo in os.listdir(ruta_directorio) if archivo.endswith('.xlsx')]

    # Leer cada archivo Excel
    for archivo in archivos_excel:
        iniciales_materia = os.path.splitext(archivo)[0]  # Obtener las iniciales de la materia
        ruta_archivo = os.path.join(ruta_directorio, archivo)
        horarios_profesores[iniciales_materia] = {}

        # Leer cada hoja dentro del archivo
        with pd.ExcelFile(ruta_archivo) as xls:
            for nombre_profesor in xls.sheet_names:
                # Leer horarios del profesor
                horarios = pd.read_excel(xls, nombre_profesor)
                # Eliminar filas donde todas las celdas son 'Hora Libre'
                horarios = horarios.dropna(how='all')
                # Almacenar en el diccionario
                horarios_profesores[iniciales_materia][nombre_profesor] = horarios

    return horarios_profesores

# Función para mostrar los horarios de los profesores en un formato tabular
def mostrar_horarios_tabla(horarios):
    for materia, profesores in horarios.items():
        print(f"Materia: {materia}")
        for profesor, horarios_profesor in profesores.items():
            print(f"Profesor: {profesor}")
            # Crear una nueva tabla con solo Día, Hora Inicio, Hora Fin y Asignatura
            tabla_horarios = horarios_profesor[['Inicio', 'Final', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado']].copy()
            tabla_horarios = tabla_horarios.melt(id_vars=['Inicio', 'Final'], var_name='Día', value_name='Asignatura')
            tabla_horarios = tabla_horarios[tabla_horarios['Asignatura'] != 'Hora Libre'].dropna()
            # Convertir a formato tabular usando tabulate
            tabla_formateada = tabulate(tabla_horarios, headers='keys', tablefmt='pretty', showindex=False)
            print(tabla_formateada)

# Función para seleccionar la carpeta
def seleccionar_carpeta(ruta_directorio_excel):
    ventana = tk.Toplevel()
    ventana.title("Seleccionar Subcarpeta")
    
    etiqueta = tk.Label(ventana, text="Selecciona una subcarpeta:")
    etiqueta.pack()
    
    combo_subcarpetas = ttk.Combobox(ventana, state="readonly")
    combo_subcarpetas.pack()
    
    # Obtener subcarpetas de la carpeta principal
    subcarpetas = [nombre for nombre in os.listdir(ruta_directorio_excel) if os.path.isdir(os.path.join(ruta_directorio_excel, nombre))]
    combo_subcarpetas["values"] = subcarpetas
    
    boton_seleccionar = tk.Button(ventana, text="Seleccionar", command=lambda: mostrar_horarios(ruta_directorio_excel, combo_subcarpetas.get()))
    boton_seleccionar.pack()

# Función para mostrar los horarios de la subcarpeta seleccionada
def mostrar_horarios(ruta_directorio_principal, subcarpeta):
    ruta_directorio_subcarpeta = os.path.join(ruta_directorio_principal, subcarpeta)
    horarios_profesores = leer_horarios_profesores(ruta_directorio_subcarpeta)
    print("Horarios de los profesores:")
    mostrar_horarios_tabla(horarios_profesores)



class MainWindow(QWidget):
    def __init__(self, datos_estudiante):
        super().__init__()
        self.datos_estudiante = datos_estudiante
        self.setWindowTitle("Menú de opciones")
        self.setGeometry(100,100,550,450)
        self.label= QLabel("Bienvenid@, "+self.datos_estudiante.nombre + ". Elige una opción", self)
        self.label.setFont(QFont('Arial',14))
        self.label.setStyleSheet("border: 2px solid transparent;\n"
                                   "border-bottom: 1px solid #21505D;\n"
                                   "border-radius: 4px;")
        self.label.move(20, 10)
        self.label.setFixedSize(400, 30)
        self.imagen = QLabel(self)
        pixmap = QPixmap("C:\\Users\\david\\Desktop\\pro\\login\\background.jpg")
        self.imagen.setPixmap(pixmap)
        self.imagen.setFixedSize(345, 208)
        self.imagen.move(185, 125)
        self.imagen.setScaledContents(True)
        layout = QVBoxLayout()
        self.button1 = QPushButton("Elegir horario")
        self.button1.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;\n")
        self.button1.setFixedSize(150, 30)
        self.button1.clicked.connect(self.iniciar_eleccion_horario)
        layout.addWidget(self.button1)

        self.button2 = QPushButton("Solicitar certificado")
        
        self.button2.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button2.setFixedSize(150, 30)
        self.button2.clicked.connect(self.menuCertificados)
        layout.addWidget(self.button2)

        self.button3 = QPushButton("Informacion estudiante")
        self.button3.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button3.setFixedSize(150, 30)
        self.button3.clicked.connect(self.infoEstudiantes)
        layout.addWidget(self.button3)
        
        self.button_salir = QPushButton("Salir")
        self.button_salir.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button_salir.setFixedSize(150, 30)
        self.button_salir.clicked.connect(self.salir)
        layout.addWidget(self.button_salir)

        self.setLayout(layout)

    


    def salir(self):
        self.close()


    def menuCertificados(self):
        self.menuCertificados=menuCertificados(self.datos_estudiante)
        self.menuCertificados.show()
        self.close()

    def iniciar_eleccion_horario(self):
        self.horarios=MenuHorarios(self.datos_estudiante)
        self.horarios.show()
        self.close()

    def infoEstudiantes(self):
        self.info=InfoEstudiante(self.datos_estudiante)
        self.info.show()
        self.close()

class MenuHorarios(QWidget):
    def __init__(self, datos_estudiante):
        super().__init__()
        self.datos_estudiante = datos_estudiante
        self.setWindowTitle("Eleccion de horario")
        self.setGeometry(100,100,550,450)
        self.label= QLabel("Certificados de estudio", self)
        self.label.setFont(QFont('Arial',14))
        self.label.setStyleSheet("border: 2px solid transparent;\n"
                                   "border-bottom: 1px solid #21505D;\n"
                                   "border-radius: 4px;")
        self.label.move(20, 10)
        self.label.setFixedSize(225, 30)
        self.imagen = QLabel(self)
        pixmap = QPixmap("C:\\Users\\david\\Desktop\\pro\\login\\im2.jpg")
        self.imagen.setPixmap(pixmap)
        self.imagen.setFixedSize(263, 350)
        self.imagen.move(220, 65)
        self.imagen.setScaledContents(True)
        layout = QVBoxLayout()

        self.button1 = QPushButton("Ing. Administrativa")
        self.button1.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button1.setFixedSize(150, 30)
        self.button1.clicked.connect(self.botonAdmin)

        layout.addWidget(self.button1)

        self.button2 = QPushButton("Ing. Biomedica")
        self.button2.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button2.setFixedSize(150, 30)
        self.button2.clicked.connect(self.botonBiomedica)
        
        layout.addWidget(self.button2)

        self.button3 = QPushButton("Ing. Sistemas")
        self.button3.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button3.setFixedSize(150, 30)
        self.button3.clicked.connect(self.botonSistemas)
       
        layout.addWidget(self.button3)
        
        self.button_atras = QPushButton("Atras")
        self.button_atras.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button_atras.setFixedSize(150, 30)
        self.button_atras.clicked.connect(self.volver)
        layout.addWidget(self.button_atras)

        self.setLayout(layout)

    def volver(self):
        self.main_window=MainWindow(self.datos_estudiante)
        self.main_window.show()
        self.close()    

    def botonBiomedica(self):
        ruta_directorio_excel = "C:\\Users\david\Desktop\pr2\Proyecto2\Biomedica"
        ventana_principal = tk.Tk()
        boton_seleccionar = tk.Button(ventana_principal, text="Seleccionar Carpeta", command=lambda: seleccionar_carpeta(ruta_directorio_excel))
        boton_seleccionar.pack()
        ventana_principal.mainloop()
        

    def botonAdmin(self):
        ruta_directorio_excel = "C:\\Users\david\Desktop\pr2\Proyecto2\Admin"
        ventana_principal = tk.Tk()
        boton_seleccionar = tk.Button(ventana_principal, text="Seleccionar Carpeta", command=lambda: seleccionar_carpeta(ruta_directorio_excel))
        boton_seleccionar.pack()
        ventana_principal.mainloop()     
        

    def botonSistemas(self):
        
        ruta_directorio_excel = "C:\\Users\david\Desktop\pr2\Proyecto2\Sistemas"
        ventana_principal = tk.Tk()
        boton_seleccionar = tk.Button(ventana_principal, text="Seleccionar Carpeta", command=lambda: seleccionar_carpeta(ruta_directorio_excel))
        boton_seleccionar.pack()
        ventana_principal.mainloop()

    def buscarCarrera(self,carrera):
        for elemento in self.datos_estudiante.carrera: 
            if elemento == carrera:
                return True
            break

class menuCertificados(QWidget):
    def __init__(self, datos_estudiante):
        super().__init__()
        self.datos_estudiante = datos_estudiante
        self.setWindowTitle("Solicitar certificados")
        self.setGeometry(100,100,550,450)
        self.label= QLabel("Certificados de estudio", self)
        self.label.setFont(QFont('Arial',14))
        self.label.setStyleSheet("border: 2px solid transparent;\n"
                                   "border-bottom: 1px solid #21505D;\n"
                                   "border-radius: 4px;")
        self.label.move(20, 10)
        self.label.setFixedSize(225, 30)
        self.imagen = QLabel(self)
        pixmap = QPixmap("C:\\Users\\david\\Desktop\\pro\\login\\im3.jpg")
        self.imagen.setPixmap(pixmap)
        self.imagen.setFixedSize(345, 192)
        self.imagen.move(190, 125)
        self.imagen.setScaledContents(True)

        layout = QVBoxLayout()


        self.button1 = QPushButton("Solicitar certificado de estudio")
        self.button1.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button1.setFixedSize(170, 30)
        self.button1.clicked.connect(self.button_estudios)
        layout.addWidget(self.button1)

        self.button2 = QPushButton("Solicitar certificado de notas")
        self.button2.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button2.clicked.connect(self.button_notas)
        self.button2.setFixedSize(170, 30)
        layout.addWidget(self.button2)

        self.button_atras = QPushButton("Atras")
        self.button_atras.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button_atras.setFixedSize(170, 30)
        self.button_atras.clicked.connect(self.volver)
        layout.addWidget(self.button_atras)
    
        self.setLayout(layout)

    def volver(self):
        self.main_window=MainWindow(self.datos_estudiante)
        self.main_window.show()
        self.close()

    def button_estudios(self):
        print("Botón certificado estudio clickeado")
        QMessageBox.information(self, "Envio exitoso", f"Certificado eviado exitosamente al correo!")
        self.generarPDF()

    def button_notas(self):
        print("Botón certificado notas clickeado")
        QMessageBox.information(self, "Envio exitoso", f"Certificado eviado exitosamente al correo!")
        self.generarPDF()
    
    def generarPDF(self):
         c = canvas.Canvas("certificado.pdf", pagesize=letter)
         c.setFont("Times-Roman", 12)
         c.setLineWidth(.3)
        
        # Titulo
         c.drawString(200, 750, "Certificado de Estudios")
         c.drawString(200, 730, "Universidad EIA")
        
        # Informacion del estudiante
         c.drawString(30, 700, "Facultad de [Nombre de la facultad]")  
         c.drawString(30, 680, "Programa Académico: " )
         c.drawString(30, 660, "CERTIFICA:")
         c.drawString(30, 640, "Que "+self.datos_estudiante.nombre +   ", identificado(a) con "+self.datos_estudiante.cc  + ", cursó y aprobó el programa académico")
         c.drawString(30, 620, "[carrera estudiante]" + " en la Universidad EIA, obteniendo el título de [carrera estudiante]"  + " el [fecha]"  + ".")
         c.drawString(30, 600, "Durante su formación académica, el(la) estudiante obtuvo un promedio acumulado de [prom estudiante]"  )
         c.drawString(30, 580, "y aprobó un total de [total creditos estudiante]" +  " créditos.")
        
        
        # Fecha de expedicion y firma
         c.drawString(30, 500, "Este certificado se expide a solicitud del(la) interesado(a) en Medellín, el día " + time.strftime("%Y-%m-%d"))
         c.drawString(30, 480, "Firma del secretario Académico:")
        
         c.showPage()
         c.save()

        # Envio del correo con el PDF adjunto
         cuerpo = "Certificado de estudio en la Universidad EIA"
         mensaje = MIMEMultipart()
         mensaje['From'] = "certificadoeia@gmail.com"
         mensaje['To'] = "jimenezd12345@gmail.com"
         mensaje['Subject'] = "Certificado"
         mensaje.attach(MIMEText(cuerpo, 'plain'))

         archivo_adjunto = open("certificado.pdf", 'rb')
         adjunto_MIME = MIMEBase('application', 'octet-stream')
         adjunto_MIME.set_payload((archivo_adjunto).read())
         encoders.encode_base64(adjunto_MIME)
         adjunto_MIME.add_header('Content-Disposition', "attachment; filename= %s" % "documento.pdf")
         mensaje.attach(adjunto_MIME)

         sesion_smtp = smtplib.SMTP('smtp.gmail.com', 587)
         sesion_smtp.starttls()
         sesion_smtp.login("certificadoeia@gmail.com", "gfqh xuoj axxu cwwl")
         texto = mensaje.as_string()
         sesion_smtp.sendmail("certificadoeia@gmail.com", self.datos_estudiante.correo, texto)
         sesion_smtp.quit()
    
class InfoEstudiante(QWidget):
    def __init__(self, datos_estudiante):
        super().__init__()
        self.datos_estudiante = datos_estudiante
        self.setWindowTitle("Info estudiante")
        self.setGeometry(100,100,550,450)

        layout = QVBoxLayout()
        nombre_label= QLabel(self)
        nombre_label.setText("Nombre: "+self.datos_estudiante.nombre)
        nombre_label.setFont(QFont('Arial',10))
        nombre_label.setStyleSheet("border: 2px solid transparent;\n"
                                   "border-bottom: 1px solid #21505D;\n"
                                   "border-radius: 4px;")
        layout.addWidget(nombre_label)
        car_label= QLabel(self)
        car_label.setText(f"Carrera(s): {self.datos_estudiante.carrera}")
        car_label.setFont(QFont('Arial',10))
        car_label.setStyleSheet("border: 2px solid transparent;\n"
                                   "border-bottom: 1px solid #21505D;\n"
                                   "border-radius: 4px;")
        layout.addWidget(car_label)
        sem_label= QLabel(self)
        sem_label.setText("Semestre: "+str(self.datos_estudiante.semestre))
        sem_label.setFont(QFont('Arial',10))
        layout.addWidget(sem_label)
        sem_label.setStyleSheet("border: 2px solid transparent;\n"
                                   "border-bottom: 1px solid #21505D;\n"
                                   "border-radius: 4px;")
        prom_label= QLabel(self)
        prom_label.setText("Promedio: "+str(self.datos_estudiante.promedio))
        prom_label.setFont(QFont('Arial',10))
        prom_label.setStyleSheet("border: 2px solid transparent;\n"
                                   "border-bottom: 1px solid #21505D;\n"
                                   "border-radius: 4px;")
        layout.addWidget(prom_label)
        self.button_atras = QPushButton("Atras")
        self.button_atras.setStyleSheet("border: 2px solid transparent;\n"
                                   "color: #fff;\n"
                                   "background-color: #21505D;\n"
                                   "border-radius: 4px;")
        self.button_atras.clicked.connect(self.volver)
        layout.addWidget(self.button_atras) 

        self.setLayout(layout)
    def volver(self):
        self.main_window=MainWindow(self.datos_estudiante)
        self.main_window.show()
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())