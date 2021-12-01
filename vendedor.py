import sys
import os
from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap , QIcon
from PyQt5.QtCore import Qt,QEasingCurve, QPropertyAnimation
from reportlab.lib.utils import isNonPrimitiveInstance
import rpyc
import socket
from time import sleep
from datetime import datetime, timedelta
import json
import ctypes
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch , mm
from reportlab.lib.colors import white, black
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder

from openpyxl.utils import get_column_letter
import subprocess

class Login(QDialog):
    ventana_principal = 0
    def __init__(self, parent):
        super(Login,self).__init__(parent)
        uic.loadUi('login2.ui',self)

        self.setModal(True)
        self.conexion = None
        self.host = None
        self.puerto = None
        self.exito = False
        self.actual = None
        self.datos_usuario = 'USER: NINGUNO BRO'
        self.txt_contra.setEchoMode(QLineEdit.Password)
        self.btn_manual.clicked.connect(self.conectar_manual)
        self.btn_iniciar.clicked.connect(self.iniciar)
        self.btn_manual.clicked.connect
        self.inicializar()
        self.txt_usuario.setFocus()
        

    def inicializar(self):
        actual = os.path.abspath(os.getcwd())
        actual = actual.replace('\\' , '/')
        self.actual = actual
        ruta = actual + '/icono_imagen/madenco logo.png'
        foto = QPixmap(ruta)
        self.lb_logo.setPixmap(foto)

        if os.path.isfile(actual + '/manifest.txt'):
            print('encontrado manifest')
            with open(actual + '/manifest.txt' , 'r', encoding='utf-8') as file:
                lines = file.readlines()
                try:
                    n_host = lines[0].split(':')
                    n_host = n_host[1]
                    host = n_host[:len(n_host)-1]

                    n_port = lines[1].split(':')
                    n_port = n_port[1]
                    port = n_port[:len(n_port)-1]

                    self.host = host
                    self.puerto = port
                except IndexError:
                    print('error de indice del manifest')
                    pass #si no encuentra alguna linea
        else:
            print('manifest no encontrado')
        
        if os.path.isfile(actual+ '/registry.txt'):
            with open(actual + '/registry.txt' , 'r', encoding='utf-8') as file:
                lines = file.readlines()
                try:
                    user = lines[0].split(':')
                    user = user[1]
                    user = user[:len(user)-1]

                    password = lines[1].split(':')
                    password = password[1]
                    password = password[:len(password)-1]
                    self.txt_usuario.setText(user)
                    self.txt_contra.setText(password)

                except IndexError:
                    print('error de indice del registry')
                    pass #si no encuentra alguna linea
        else:
            print('Datos de usuario no encontrados')


    def conectar(self):
        try:
            if self.host and self.puerto:
                self.conexion = rpyc.connect(self.host , self.puerto)
                self.lb_conexion.setText('CONECTADO')
            else:
                QMessageBox.about(self,'ERROR', 'Host y puerto no encontrados en el manifest' )

        except ConnectionRefusedError:
            self.lb_conexion.setText('EL SERVIDOR NO RESPONDE')
            
        except socket.error:
            self.lb_conexion.setText('SERVIDOR FUERA DE RED')

    def conectar_manual(self):
        dialog = InputDialog('HOST:','PUERTO:','CONECTAR MANUAL', self)
        dialog.resize(250,100)   
        if dialog.exec():
                hostx , puertox = dialog.getInputs()
                try:
                    if hostx != '':
                        puertox = int(puertox)
                        self.host = hostx
                        self.puerto = puertox
                        self.conexion = rpyc.connect(hostx , puertox)
                        self.lb_conexion.setText('CONECTADO')
                    else: 
                        QMessageBox.about(self,'ERROR' ,'Ingrese un host antes de continuar')
                except ValueError:
                    QMessageBox.about(self,'ERROR' ,'Ingrese solo numeros en el PUERTO')
                except ConnectionRefusedError:
                    self.lb_conexion.setText('EL SERVIDOR NO RESPONDE')
                    
                except socket.error:
                    self.lb_conexion.setText('SERVIDOR FUERA DE LA RED')
                    #QMessageBox.about(self,'ERROR' ,'No se puede establecer la conexion con el servidor')

    def guardar_datos(self):
        
        if self.checkBox.isChecked():

            with open(self.actual + '/registry.txt' , 'w', encoding='utf-8') as file:
                file.write('usuario:'+ self.txt_usuario.text() + '\n')
                file.write('contra:' + self.txt_contra.text()+ '\n')
                file.write('')

    def iniciar(self):
       

        if self.conexion == None: #no existe la conexion
            self.conectar()
            usuario = self.txt_usuario.text()
            contra = self.txt_contra.text()
            try:
                resultado = self.conexion.root.obtener_usuario_activo()
                encontrado = False
                for item in resultado:
                    if item[0] == usuario and item[1] == contra and item[6] == 'vendedor':
                        encontrado = True
                        self.datos_usuario = item
                        self.guardar_datos()
                        self.accept()
                if encontrado == False:
                    QMessageBox.about(self ,'ERROR', 'Usuario o contraseña invalidas')

            except EOFError:
                QMessageBox.about(self ,'Conexion', 'El servidor no responde')
            except AttributeError:
                pass

        else: #existe la conexion
            usuario = self.txt_usuario.text()
            contra = self.txt_contra.text()
            try:
                resultado = self.conexion.root.obtener_usuario_activo()
                encontrado = False
                for item in resultado:
                    if item[0] == usuario and item[1] == contra and item[6] == 'vendedor':
                        encontrado = True
                        self.datos_usuario = item
                        print('USUARIO ENCONTRADO')
                        self.guardar_datos()
                        self.accept()
                if encontrado == False:
                    QMessageBox.about(self ,'ERROR', 'Usuario o contraseña invalidas')

            except EOFError:
                QMessageBox.about(self ,'Conexion', 'El servidor no responde') 
            except AttributeError:
                pass
    
    def obt_datos(self):
        return self.datos_usuario , self.conexion ,self.host , self.puerto

    def closeEvent(self, event):
        """Generate 'question' dialog on clicking 'X' button in title bar.

        Reimplement the closeEvent() event handler to include a 'Question'
        dialog with options on how to proceed - Save, Close, Cancel buttons
        """
        reply = QMessageBox.question(
            self, "Salir",
            "¿Deseas salir de la aplicación?",
            QMessageBox.Close | QMessageBox.Cancel)

        if reply == QMessageBox.Close:
            event.accept()
        else:
            event.ignore()

class Vendedor(QMainWindow):
    ventana_login = 0

    def __init__(self):
        super( Vendedor,self).__init__()
        uic.loadUi('vendedor.ui',self)
        self.iniciar_session()
        self.inicializar()
        self.btn_buscar.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.buscar_venta))
        self.btn_menu.clicked.connect(self.mostrar_menu)

    def iniciar_session(self):
        self.ventana_login = Login(self)
        self.ventana_login.show()
        salir = self.ventana_login.exec()
        if salir == 0:
            print('cerrar aplicacion')
            sys.exit(0)
        elif salir == 1:
            datos , conn , host , puerto = self.ventana_login.obt_datos()
            self.datos_usuario = datos
            self.conexion = conn
            self.host = host
            self.puerto = puerto

            print(datos)
            print('mostrar el menu')
        print('continuando..')


    def inicializar(self):
        print('inicializando...')
        self.btn_generar_clave.show()
        self.btn_orden_manual.show()
        self.btn_informe.show()

        actual = os.path.abspath(os.getcwd())
        actual = actual.replace('\\' , '/')

        foto = QPixmap(actual + '/icono_imagen/madenco logo.png')
        self.logo.setPixmap(foto)
        menu = QPixmap(actual + '/icono_imagen/left_menu_v3.png')
        self.btn_menu.setIcon(QIcon(menu))

        self.lb_conexion.setText('CONECTADO')
        if self.datos_usuario[4] == 'NO': #Si no es super usuario
            self.btn_generar_clave.hide() #no puede generar claves
            detalle = json.loads(self.datos_usuario[7])
            funciones = detalle['vendedor']
            if not 'manual' in funciones:
                self.btn_orden_manual.hide() #no puede generar ordenes manuales
            if not 'informes' in funciones:
                self.btn_informe.hide() #no puede generar informes
        

        self.btn_atras.setIcon(QIcon('icono_imagen/atras.ico'))
    

    def mostrar_menu(self):
        
        if True:
            ancho = self.left_menu_container.width()
            normal = 100
            if ancho == 100:
                extender = 200
                self.logo.show()
                self.btn_buscar.setText('Notas de venta')
                self.btn_modificar.setText('Ordenes de trabajo')
                self.btn_orden_manual.setText('Ingreso manual')
                self.btn_generar_clave.setText('Generar clave')
                self.btn_informe.setText('Generar informe')
            else:
                self.logo.hide()
                self.btn_buscar.setText('')
                self.btn_modificar.setText('')
                self.btn_orden_manual.setText('')
                self.btn_generar_clave.setText('')
                self.btn_informe.setText('')
                self.btn_buscar.setIcon(QIcon('icono_imagen/close.png'))
                self.btn_modificar.setIcon(QIcon('icono_imagen/close.png'))
                self.btn_orden_manual.setIcon(QIcon('icono_imagen/close.png'))
                self.btn_generar_clave.setIcon(QIcon('icono_imagen/close.png'))
                self.btn_informe.setIcon(QIcon('icono_imagen/close.png'))
                extender = normal
            
            self.animation = QPropertyAnimation(self.left_menu_container, b"maximumWidth" )
            self.animation.setDuration(500)
            self.animation.setStartValue(ancho)
            self.animation.setEndValue(extender)
            self.animation.setEasingCurve(QEasingCurve.InOutQuart)
            self.animation.start()

class InputDialog(QDialog):
    def __init__(self,label1,label2,title ,parent=None):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.txt1 = QLineEdit(self)
        self.txt2 = QLineEdit(self)
        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self);
       
        layout = QFormLayout(self)
        layout.addRow(label1, self.txt1)
        layout.addRow(label2, self.txt2)
        layout.addWidget(buttonBox)

        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)
    
    def getInputs(self):
        return self.txt1.text(), self.txt2.text()


    
if __name__ == '__main__':
    x = 0
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('icono_imagen/madenco logo.ico'))
    myappid = 'madenco.personal.area' # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid) 
    vendedor = Vendedor()
    vendedor.show()
    sys.exit(app.exec_())
