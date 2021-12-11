import sys
import os
from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap , QIcon
from PyQt5.QtCore import Qt,QEasingCurve, QPropertyAnimation
from PyQt5.uic.uiparser import DEBUG
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
        self.datos_usuario = None
        self.conexion = None
        self.host = None
        self.puerto = None
        
        self.vendedores = None
        self.aux_tabla = None
        self.bol_fact = None
        self.guias = None
        self.iniciar_session()
        self.inicializar()
        self.mostrar_menu()
        self.stackedWidget.setCurrentWidget(self.inicio)
        
        #buscar venta
        self.btn_buscar_1.clicked.connect(self.buscar_documento)
        self.comboBox_1.currentIndexChanged['QString'].connect(self.filtrar_vendedor)
        self.btn_crear_1.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.crear_orden))
        self.btn_atras_1.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.inicio))
        #crear orden
        self.btn_registrar_2.clicked.connect(lambda: print('registrando...'))
        self.btn_atras_2.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.buscar_venta))
        #buscar orden
        self.btn_modificar_4.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.modificar_orden))
        self.btn_atras_4.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.inicio))
        #modificar orden
        self.btn_guardar_5.clicked.connect(lambda: QMessageBox.about(self,'error','orden modiicada xd'))
        self.btn_atras_5.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.buscar_orden))

        #informes
        #SIDE MENU BOTONES
        self.btn_buscar.clicked.connect(self.inicializar_buscar_venta)
        self.btn_modificar.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.buscar_orden))
        self.btn_orden_manual.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.manual))
        self.btn_informe.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.informes))
        self.btn_atras.clicked.connect(self.cerrar_sesion)
        self.btn_conectar.clicked.connect(self.conectar)

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
        self.btn_atras.setIcon(QIcon('icono_imagen/logout.png'))
        self.btn_buscar.setIcon(QIcon('icono_imagen/venta.png'))
        self.btn_modificar.setIcon(QIcon('icono_imagen/orden.png'))
        self.btn_orden_manual.setIcon(QIcon('icono_imagen/manual.png'))
        self.btn_generar_clave.setIcon(QIcon('icono_imagen/key2.png'))
        self.btn_informe.setIcon(QIcon('icono_imagen/informe.png'))

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
        if self.datos_usuario:  #Si existen los datos del usuario, x ende se inicio sesion correctamente...
            self.lb_vendedor.setText(self.datos_usuario[8])
            if self.datos_usuario[4] == 'NO': #Si no es super usuario
                self.btn_generar_clave.hide() #no puede generar claves
                detalle = json.loads(self.datos_usuario[7])
                funciones = detalle['vendedor']
                if not 'manual' in funciones:
                    self.btn_orden_manual.hide() #no puede generar ordenes manuales
                if not 'informes' in funciones:
                    self.btn_informe.hide() #no puede generar informes
        

        self.btn_atras.setIcon(QIcon('icono_imagen/atras.ico'))
    
    def conectar(self):
        if self.conexion == None:
            try:
                if self.host and self.puerto:
                    self.conexion = rpyc.connect(self.host , self.puerto)
                    self.lb_conexion.setText('CONECTADO')
                else:
                    self.lb_conexion.setText('Host y Puerto no encontrados')

            except ConnectionRefusedError:
                self.lb_conexion.setText('EL SERVIDOR NO RESPONDE')
                
            except socket.error:
                self.lb_conexion.setText('SERVIDOR FUERA DE RED')
    # ------------   FUNCIONES BUSCAR VENTA  -----------------
    def inicializar_buscar_venta(self):
        self.tableWidget_1.setRowCount(0)
        self.radio2_1.setChecked(True)
        self.dateEdit_1.setCalendarPopup(True)
        self.dateEdit_1.setDate(datetime.now().date())
        self.stackedWidget.setCurrentWidget(self.buscar_venta)

    def buscar_documento(self):
        largo = self.comboBox_1.count()
        if largo > 1:
            for i in range(largo):
                #print("borranto: " + str(largo - 1) )
                self.comboBox_1.removeItem(1)

        if self.conexion:
            no_encontrados = True
            self.tableWidget_1.setRowCount(0)
            if self.radio1_1.isChecked():  #BUSCANDO POR NUMERO INTERNO
                inter = self.txt_interno_1.text()
                try:
                    inter = int(inter)
                    consulta = self.conexion.root.buscar_venta_interno(inter)
                    guia = self.conexion.root.obtener_guia_interno(inter)
                    if consulta != None :
                        no_encontrados = False
                        print(consulta[1])
                        fila = self.tableWidget_1.rowCount()
                        self.tableWidget_1.insertRow(fila)

                        self.tableWidget_1.setItem(fila , 0 , QTableWidgetItem(str(consulta[0]))) #INTERNO
                        if consulta[3] == 0 : #es boleta
                            self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'BOLETA' )) #TIPO DOCUMENTO
                            self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str(consulta[4]) ))      #NRO DOCUMENTO
                        elif consulta[4] == 0: #es factura
                            self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'FACTURA' )) #TIPO DOCUMENTO
                            self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str( consulta[3])))      #NRO DOCUMENTO

                        self.tableWidget_1.setItem(fila , 3 , QTableWidgetItem(   str(consulta[1]  ))) #FECHA VENTA
                        self.tableWidget_1.setItem(fila , 4 , QTableWidgetItem( str(consulta[6] )   ))      #CLIENTE
                        self.tableWidget_1.setItem(fila , 5 , QTableWidgetItem(    consulta[2] ))             #VENDEDOR
                        
                        self.tableWidget_1.setItem(fila , 6 , QTableWidgetItem(   str(consulta[5]))   )      #TOTAL
                    if guia != None:
                        no_encontrados = False
                        
                        consulta = guia
                        detalle = json.loads(consulta[2])

                        fila = self.tableWidget_1.rowCount()
                        self.tableWidget_1.insertRow(fila)

                        self.tableWidget_1.setItem(fila , 0 , QTableWidgetItem(str(consulta[1]))) #INTERNO
                        
                        self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'GUIA' )) #TIPO DOCUMENTO
                        self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str( consulta[0])))      #NRO DOCUMENTO

                        self.tableWidget_1.setItem(fila , 3 , QTableWidgetItem(   str(consulta[4]  ))) #FECHA VENTA
                        self.tableWidget_1.setItem(fila , 4 , QTableWidgetItem( str(consulta[3]  )   ) )     #CLIENTE

                        self.tableWidget_1.setItem(fila , 5 , QTableWidgetItem( detalle['vendedor'] ))             #VENDEDOR
                        self.tableWidget_1.setItem(fila , 6 , QTableWidgetItem(   str(detalle['monto_final']))   )      #TOTAL

                    if no_encontrados:
                        QMessageBox.about(self,'Busqueda' ,'Documentos NO encontrados para la fecha indicada')
                    
                except ValueError:
                    QMessageBox.about(self,'ERROR' ,'Ingrese solo numeros')
                except EOFError:
                    self.conexion_perdida()

            elif self.radio2_1.isChecked(): #BUSCANDO POR FECHA ACTUAL
                lista_vendedores = []
                aux_lista = []
                date = self.dateEdit_1.date()
                aux = date.toPyDate()
                inicio = str(aux) + ' ' + '00:00:00'
                fin = str(aux) + ' ' + '23:59:59'
                try:
                    bol_fact = self.conexion.root.buscar_venta_fecha(inicio,fin)
                    guias = self.conexion.root.obtener_guia_fecha(inicio,fin)
                    if bol_fact != ():
                        no_encontrados = False
                        self.bol_fact = bol_fact
                        for consulta in bol_fact:
                            fila = self.tableWidget_1.rowCount()
                            self.tableWidget_1.insertRow(fila)
                            self.tableWidget_1.setItem(fila , 0 , QTableWidgetItem(str(consulta[0]))) #INTERNO
                            if consulta[3] == 0 : #es boleta
                                self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'BOLETA' )) #TIPO DOCUMENTO
                                self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str(consulta[4]) ))      #NRO DOCUMENTO
                            elif consulta[4] == 0: #es factura
                                self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'FACTURA' )) #TIPO DOCUMENTO
                                self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str( consulta[3])))      #NRO DOCUMENTO

                            self.tableWidget_1.setItem(fila , 3 , QTableWidgetItem(   str(consulta[1]  ))) #FECHA VENTA
                            self.tableWidget_1.setItem(fila , 4 , QTableWidgetItem( str(consulta[6] )   ))      #CLIENTE
                            self.tableWidget_1.setItem(fila , 5 , QTableWidgetItem(    consulta[2] ))             #VENDEDOR
                            aux = consulta[2]
                            aux = aux[0:10]
                            #print(aux)
                            
                            if aux not in aux_lista:
                                aux_lista.append(aux)
                                lista_vendedores.append( consulta[2] )
    
                            self.tableWidget_1.setItem(fila , 6 , QTableWidgetItem(   str(consulta[5]))   )      #TOTAL
                    
                    if guias != ():
                        self.guias = guias
                        no_encontrados = False
                        for consulta in guias:
                            detalle = json.loads(consulta[2])

                            fila = self.tableWidget_1.rowCount()
                            self.tableWidget_1.insertRow(fila)

                            self.tableWidget_1.setItem(fila , 0 , QTableWidgetItem(str(consulta[1]))) #INTERNO
                           
                            self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'GUIA' )) #TIPO DOCUMENTO
                            self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str( consulta[0])))      #NRO DOCUMENTO

                            self.tableWidget_1.setItem(fila , 3 , QTableWidgetItem(   str(consulta[4]  ))) #FECHA VENTA
                            
                            self.tableWidget_1.setItem(fila , 4 , QTableWidgetItem( str(consulta[3]  )   ) )     #CLIENTE

                            self.tableWidget_1.setItem(fila , 5 , QTableWidgetItem( detalle['vendedor'] ))             #VENDEDOR
                            aux = detalle['vendedor']
                            aux = aux[0:10]
                            #print(aux)
                            
                            if aux not in aux_lista:
                                aux_lista.append(aux)
                                lista_vendedores.append( detalle['vendedor'] )
    
                            self.tableWidget_1.setItem(fila , 6 , QTableWidgetItem(   str(detalle['monto_final']))   )      #TOTAL
                

                    self.vendedores = lista_vendedores
                    for item in self.vendedores:
                        self.comboBox_1.addItem(item)

                    if no_encontrados:
                        QMessageBox.about(self,'Busqueda' ,'Documentos NO encontrados para la fecha indicada')

                except EOFError:
                    self.conexion_perdida()
            elif self.radio3_1.isChecked(): #BUSCANDO POR NOMBRE
                X = 0
                lista_vendedores = []
                aux_lista = []
                nombre = self.txt_cliente_1.text()
                try:
                    bol_fact = self.conexion.root.obtener_venta_nombre(nombre)
                    guias = self.conexion.root.obtener_guia_nombre(nombre)
                    if bol_fact != ():
                        no_encontrados = False
                        self.bol_fact = bol_fact
                        for consulta in bol_fact:
                            fila = self.tableWidget_1.rowCount()
                            self.tableWidget_1.insertRow(fila)
                            self.tableWidget_1.setItem(fila , 0 , QTableWidgetItem(str(consulta[0]))) #INTERNO
                            if consulta[3] == 0 : #es boleta
                                self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'BOLETA' )) #TIPO DOCUMENTO
                                self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str(consulta[4]) ))      #NRO DOCUMENTO
                            elif consulta[4] == 0: #es factura
                                self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'FACTURA' )) #TIPO DOCUMENTO
                                self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str( consulta[3])))      #NRO DOCUMENTO

                            self.tableWidget_1.setItem(fila , 3 , QTableWidgetItem(   str(consulta[1]  ))) #FECHA VENTA
                            self.tableWidget_1.setItem(fila , 4 , QTableWidgetItem( str(consulta[6] )   ))      #CLIENTE

                            self.tableWidget_1.setItem(fila , 5 , QTableWidgetItem(    consulta[2] ))             #VENDEDOR
                            aux = consulta[2]
                            aux = aux[0:10]
                            #print(aux)
                            
                            if aux not in aux_lista:
                                aux_lista.append(aux)
                                lista_vendedores.append( consulta[2] )
    
                            self.tableWidget_1.setItem(fila , 6 , QTableWidgetItem(   str(consulta[5]))   )      #TOTAL
                    
                    if guias != ():
                        self.guias = guias
                        no_encontrados = False
                        for consulta in guias:
                            detalle = json.loads(consulta[2])

                            fila = self.tableWidget_1.rowCount()
                            self.tableWidget_1.insertRow(fila)

                            self.tableWidget_1.setItem(fila , 0 , QTableWidgetItem(str(consulta[1]))) #INTERNO
                           
                            self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'GUIA' )) #TIPO DOCUMENTO
                            self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str( consulta[0])))      #NRO DOCUMENTO

                            self.tableWidget_1.setItem(fila , 3 , QTableWidgetItem(   str(consulta[4]  ))) #FECHA VENTA
                            
                            self.tableWidget_1.setItem(fila , 4 , QTableWidgetItem( str(consulta[3] )   ))      #CLIENTE

                            self.tableWidget_1.setItem(fila , 5 , QTableWidgetItem( detalle['vendedor'] ))             #VENDEDOR
                            aux = detalle['vendedor']
                            aux = aux[0:10]
                            #print(aux)
                            
                            if aux not in aux_lista:
                                aux_lista.append(aux)
                                lista_vendedores.append( detalle['vendedor'] )
    
                            self.tableWidget_1.setItem(fila , 6 , QTableWidgetItem(   str(detalle['monto_final']))   )      #TOTAL
                

                    self.vendedores = lista_vendedores
                    for item in self.vendedores:
                        self.comboBox_1.addItem(item)

                    if no_encontrados:
                        QMessageBox.about(self,'Busqueda' ,'Documentos NO encontrados para la fecha indicada')

                except EOFError:
                    self.conexion_perdida()


            self.aux_tabla = self.tableWidget_1
        else:
            self.conexion_perdida()

    def rellenar_tabla(self):
        self.tableWidget_1.setRowCount(0)
        if self.bol_fact != None:
            for consulta in self.bol_fact:
                fila = self.tableWidget_1.rowCount()
                self.tableWidget_1.insertRow(fila)
                self.tableWidget_1.setItem(fila , 0 , QTableWidgetItem(str(consulta[0]))) #INTERNO
                if consulta[3] == 0 : #es boleta
                    self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'BOLETA' )) #TIPO DOCUMENTO
                    self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str(consulta[4]) ))      #NRO DOCUMENTO
                elif consulta[4] == 0: #es factura
                    self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'FACTURA' )) #TIPO DOCUMENTO
                    self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str( consulta[3])))      #NRO DOCUMENTO

                self.tableWidget_1.setItem(fila , 3 , QTableWidgetItem(   str(consulta[1]  ))) #FECHA VENTA
                self.tableWidget_1.setItem(fila , 4 , QTableWidgetItem( str(consulta[6] )   ))      #CLIENTE
                self.tableWidget_1.setItem(fila , 5 , QTableWidgetItem(    consulta[2] ))             #VENDEDOR
                self.tableWidget_1.setItem(fila , 6 , QTableWidgetItem(   str(consulta[5]))   )      #TOTAL

        if self.guias != None:
            for consulta in self.guias:
                detalle = json.loads(consulta[2])
                fila = self.tableWidget_1.rowCount()
                self.tableWidget_1.insertRow(fila)
                self.tableWidget_1.setItem(fila , 0 , QTableWidgetItem(str(consulta[1]))) #INTERNO
                
                self.tableWidget_1.setItem(fila , 1 , QTableWidgetItem( 'GUIA' )) #TIPO DOCUMENTO
                self.tableWidget_1.setItem(fila , 2 , QTableWidgetItem( str( consulta[0])))      #NRO DOCUMENTO

                self.tableWidget_1.setItem(fila , 3 , QTableWidgetItem(   str(consulta[4]  ))) #FECHA VENTA
                self.tableWidget_1.setItem(fila , 4 , QTableWidgetItem( str(consulta[3] )   ))      #CLIENTE
                self.tableWidget_1.setItem(fila , 5 , QTableWidgetItem( detalle['vendedor'] ))             #VENDEDOR
            
                self.tableWidget_1.setItem(fila , 6 , QTableWidgetItem(   str(detalle['monto_final']))   )      #TOTAL

    def filtrar_vendedor(self):
        vendedor = self.comboBox_1.currentText()
        
        if vendedor == 'TODOS':
            self.rellenar_tabla()
        else:
            self.rellenar_tabla()
            self.tableWidget_1 = self.aux_tabla
            remover = []
            vendedor = vendedor[0:10]
            print('solo: ' + vendedor)
            column = 5 #COLUMNA DEL CVENDEDOR EN LA TABLA
            # rowCount() This property holds the number of rows in the table
            for row in range(self.tableWidget_1.rowCount()): 
                # item(row, 0) Returns the item for the given row and column if one has been set; otherwise returns nullptr.
                _item = self.tableWidget_1.item(row, column) 
                if _item:            
                    item = self.tableWidget_1.item(row, column).text()
                    print(f'row: {row}, column: {column}, item={item}')
                    aux_item = item[0:10]
                    if aux_item != vendedor:
                        remover.append(row)
            print(remover)
            k = 0
            for i in remover:
                self.tableWidget_1.removeRow(i - k)
                k += 1

    #  -------------------------------------------------------
    # ------------ FUNCIONES PARA CREAR ORDEN ---------------------
    def inicializar_crear_orden():
        X = 0


    def cerrar_sesion(self):
        self.iniciar_session()
        self.inicializar()

    def mostrar_menu(self):   
        if True:
            ancho = self.left_menu_container.width()
            normal = 70
            if ancho == 70:
                extender = 150
                self.logo.show()
                self.btn_buscar.setText('Notas de venta')
                self.btn_modificar.setText('Ordenes de trabajo')
                self.btn_orden_manual.setText('Ingreso manual')
                self.btn_generar_clave.setText('Generar clave')
                self.btn_informe.setText('Generar informe')
                self.btn_atras.setText('Cerrar sesión')
            else:
                self.logo.hide()
                self.btn_buscar.setText('')
                self.btn_modificar.setText('')
                self.btn_orden_manual.setText('')
                self.btn_generar_clave.setText('')
                self.btn_informe.setText('')
                self.btn_atras.setText('')
                
                extender = normal
            
            self.animation = QPropertyAnimation(self.left_menu_container, b"maximumWidth" )
            self.animation.setDuration(500)
            self.animation.setEasingCurve(QEasingCurve.Linear)
            self.animation.setStartValue(ancho)
            self.animation.setEndValue(extender)
            
            self.animation.start()

    def conexion_perdida(self):
        self.conexion = None
        self.lb_conexion.setText('DESCONECTADO')
        QMessageBox.about(self,'ERROR','Se perdio la conexion con el servidor')

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
