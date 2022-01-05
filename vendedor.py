import sys
import os
from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap , QIcon
from PyQt5.QtCore import Qt,QEasingCurve, QPropertyAnimation,QEvent
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

        self.carpeta = None
        self.dir_informes = None
        
        self.vendedores = None
        self.aux_tabla = None
        self.bol_fact = None
        self.guias = None
        # -----------
        self.nro_doc = 0  
        self.nro_orden = 0 #NRO FOLIO DE LA ORDEN CREADA
        self.items = None
        self.vendedor = None
        self.tipo_doc = None
        self.fecha_venta = None
        self.inter = None #nro interno del doc
        self.tipo = None # tipo de orden (dim,elab,pall o carp)
        self.manual = None #Si la orden se hizo manual o no
        self.fecha_orden = None #FECHA en la que se creo la orden, formato DATE 
        self.nro_reingreso = 0 #folio del reingreso
        self.clave = None #usada para funciones de super usuario
        self.nro_reingreso = 0
        #---------------
        self.iniciar_session()
        self.inicializar()
        self.stackedWidget.setCurrentWidget(self.inicio)
        
        #buscar venta
        self.btn_buscar_1.clicked.connect(self.buscar_documento)
        self.comboBox_1.currentIndexChanged['QString'].connect(self.filtrar_vendedor)
        self.btn_crear_1.clicked.connect(self.inicializar_crear_orden)
        self.btn_atras_1.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.inicio))
        #crear orden
        self.btn_registrar_2.clicked.connect(self.registrar_orden)
        self.btn_atras_2.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.buscar_venta))
        self.btn_agregar.clicked.connect(self.agregar)
        self.btn_eliminar.clicked.connect(self.eliminar)
        self.btn_rellenar.clicked.connect(self.rellenar)
        self.tableWidget_2.setColumnWidth(0,80)
        self.tableWidget_2.setColumnWidth(1,430)
        self.tableWidget_2.setColumnWidth(2,85)
        
        #buscar orden
        self.btn_dimensionado.clicked.connect(self.buscar_dimensionado)
        self.btn_elaboracion.clicked.connect(lambda: self.busqueda_general('ELABORACION'))
        self.btn_carpinteria.clicked.connect(lambda: self.busqueda_general('CARPINTERIA'))
        self.btn_pallets.clicked.connect(lambda: self.busqueda_general('PALLETS'))
        self.btn_modificar_4.clicked.connect(self.inicializar_modificar_orden)
        self.btn_atras_4.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.inicio))

        #modificar orden
        self.btn_guardar_5.clicked.connect( self.actualizar_orden )
        self.btn_atras_5.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.buscar_orden))
        self.btn_ver_5.clicked.connect(lambda: self.ver_pdf(self.tipo) )
        self.btn_agregar_.clicked.connect(self.agregar_2)
        self.btn_eliminar_.clicked.connect(self.eliminar_2)
        self.btn_anular_5.clicked.connect(self.anular)
        # reingreso
        self.btn_reingreso_4.clicked.connect(self.inicializar_reingreso)
        self.btn_generar_reingreso.clicked.connect(self.registrar_reingreso)
        self.btn_agregar_2.clicked.connect(self.agregar_3)
        self.btn_eliminar_2.clicked.connect(self.eliminar_3)
        self.btn_atras_7.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.buscar_orden))
        # INGRESO MANUAL
        #  ---- ORDEN MANUAL ----
        self.txt_descripcion_1.textChanged.connect(self.buscar_descripcion)
        self.txt_codigo_1.textChanged.connect(self.buscar_codigo)
        self.btn_add.clicked.connect(self.add_descripcion)
        self.r_uso_interno_1.stateChanged.connect(self.cambiar_observacion)
        self.btn_registrar_1.clicked.connect(self.registrar_orden_manual)
        self.btn_agregar_1.clicked.connect(self.agregar_4)
        self.btn_eliminar_1.clicked.connect(self.eliminar_4)
        #  ---- REINGRESO MANUAL ----
        self.btn_registrar_6.clicked.connect(self.reingreso_manual)
        self.txt_descripcion_7.textChanged.connect(self.buscar_descripcion_2)
        self.btn_add_6.clicked.connect(self.add_descripcion_2)

        self.btn_agregar_6.clicked.connect(self.agregar_6)
        self.btn_eliminar_6.clicked.connect(self.eliminar_6)

        self.btn_atras_3.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.inicio))
        self.btn_atras_6.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.inicio))
        #informes
        self.btn_generar_informe.clicked.connect(self.generar_informe)
        self.comboBox.currentIndexChanged['QString'].connect(self.vista_reingreso)
        self.btn_eliminar_exel.clicked.connect(self.eliminar_excel)
        self.btn_actualizar.clicked.connect(self.actualizar)
        self.btn_abrir.clicked.connect(self.abrir)

        #generar clave
        self.btn_generar_clave.clicked.connect(self.generar_clave)
        #SIDE MENU BOTONES
        self.btn_buscar.clicked.connect(self.inicializar_buscar_venta)
        self.btn_modificar.clicked.connect(self.inicializar_buscar_orden)
        self.btn_orden_manual.clicked.connect(self.inicializar_ingreso_manual)
        self.btn_informe.clicked.connect(self.inicializar_informe)
        self.btn_atras.clicked.connect(self.cerrar_sesion)
        self.btn_conectar.clicked.connect(self.conectar)
        self.btn_estadisticas.clicked.connect(lambda: QMessageBox.about(self,'PROXIMAMENTE', 'Este apartado esta en desarrollo'))

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
        self.btn_estadisticas.setIcon(QIcon('icono_imagen/estadisticas.png'))
        

        self.btn_generar_clave.show()
        self.btn_orden_manual.show()
        self.btn_informe.show()

        actual = os.path.abspath(os.getcwd())
        actual = actual.replace('\\' , '/')
        self.carpeta = actual.replace('\\' , '/')
        self.dir_informes = actual + '/informes/'
        print(self.carpeta)
        foto = QPixmap(actual + '/icono_imagen/ENCO_Log.png')
        self.logo.setPixmap(foto)
        menu = QPixmap(actual + '/icono_imagen/menu_v4.png')
        self.btn_menu.setIcon(QIcon(menu))
        self.lb_conexion.setText('CONECTADO')
        if self.datos_usuario:  #Si existen los datos del usuario, x ende se inicio sesion correctamente...
            self.lb_vendedor.setText(self.datos_usuario[8]) #NOMBRE DEL VENDEDOR
            if self.datos_usuario[4] == 'NO': #Si no es super usuario
                self.btn_generar_clave.hide() #no puede generar claves
                detalle = json.loads(self.datos_usuario[7])
                funciones = detalle['vendedor']
                if not 'manual' in funciones:
                    self.btn_orden_manual.hide() #no puede generar ordenes manuales
                if not 'informes' in funciones:
                    self.btn_informe.hide() #no puede generar informes
        self.stackedWidget.setCurrentWidget(self.inicio)
    
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
        self.tableWidget_1.setColumnWidth(0,80) #interno
        self.tableWidget_1.setColumnWidth(1,80) #documento
        self.tableWidget_1.setColumnWidth(2,80) #nro doc
        self.tableWidget_1.setColumnWidth(3,125) #fecha venta
        self.tableWidget_1.setColumnWidth(4,200) #vendededor
        self.tableWidget_1.setColumnWidth(5,100) #total
        self.stackedWidget.setCurrentWidget(self.buscar_venta)
    
    def buscar_documento(self):
        self.vendedores = []
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

# ------------ FUNCIONES PARA CREAR ORDEN ---------------------
    def inicializar_crear_orden(self) :
        seleccion = self.tableWidget_1.currentRow()
        if seleccion != -1:
            _item = self.tableWidget_1.item( seleccion, 0) 
            if _item:            
                interno = self.tableWidget_1.item(seleccion, 0 ).text()
                aux_tipo_doc = self.tableWidget_1.item(seleccion, 1 ).text()
                
                nro_interno = int(interno)
                self.inter = nro_interno
                print('ventana crear orden ... para: '+ aux_tipo_doc+ ' ' + str(nro_interno))

                fecha = datetime.now().date()
                self.fecha.setCalendarPopup(True)
                self.fecha.setDate(fecha)    #FECHA ESTIMADA DE ENTREGA
                self.lb_interno.setText(str(nro_interno)) #nro interno
                self.nombre_2.setText('')
                self.telefono_2.setText('')
                self.contacto_2.setText('')
                self.oce_2.setText('')
                self.tableWidget_2.setColumnWidth(0,80)
                self.tableWidget_2.setColumnWidth(1,430)
                self.tableWidget_2.setColumnWidth(2,85)
                #------- se rellena la ventana ------------------#
                try:
                    if aux_tipo_doc != 'GUIA' :
                        items = self.conexion.root.obtener_item_interno(nro_interno)
                        venta = self.conexion.root.obtener_venta_interno(nro_interno) #v5 nombre obtenido

                        aux = datetime.fromisoformat(str(venta[3] ) )
                        self.fecha_venta = aux
                        self.items = items
                        self.lb_fecha.setText(str(aux.strftime("%d-%m-%Y %H:%M:%S")))
                        if venta[1] == 0:
                            self.tipo_doc = 'BOLETA'
                            self.lb_doc.setText( self.tipo_doc )
                            self.lb_documento.setText(str(venta[2]))
                            self.nro_doc = venta[2]
                            
                        elif venta[2] == 0:
                            self.tipo_doc = 'FACTURA'
                            self.lb_doc.setText( self.tipo_doc )
                            self.lb_documento.setText(str(venta[1]))
                            self.nro_doc = venta[1]
               
                        self.vendedor = venta[4]  #VENDEDOR
                        if venta[5]:
                            self.nombre_2.setText(venta[5]) #nombre del cliente factura

                    else:
                        guia = self.conexion.root.obtener_guia_interno(nro_interno)
                        #print(guia)
                        detalle = json.loads(guia[2])
                        aux = datetime.fromisoformat(str(guia[4] ) ) #fecha de venta
                        self.fecha_venta = aux
                        self.lb_fecha.setText(str(aux.strftime("%d-%m-%Y %H:%M:%S")))
                        self.tipo_doc = aux_tipo_doc
                        self.lb_doc.setText( aux_tipo_doc )
                        self.lb_documento.setText(str(guia[0]))
                        self.nro_doc = guia[0]
                        self.nombre_2.setText(guia[3]) #nombre del cliente factura
                        self.vendedor = detalle['vendedor']  #VENDEDOR
                        i = 0
                        items = []
                        descripciones = detalle['descripciones']
                        cantidades = detalle['cantidades']
                        netos_totales = detalle['totales']
                        while i < len(cantidades) :
                            item = (cantidades[i] , descripciones[i] ,netos_totales[i])
                            items.append(item)
                            i += 1
                        self.items = items
                    print(items)
                    self.rellenar()
                    self.stackedWidget.setCurrentWidget(self.crear_orden)


                except EOFError:
                    QMessageBox.about(self, 'ERROR', 'Se perdio la conexion con el servidor')
                
        else:
            QMessageBox.about(self,'ERROR', 'Seleccione un Nro interno antes de continuar')

    def rellenar(self):
        if self.items != ():
            self.tableWidget_2.setRowCount(0)
            for item in self.items:
                fila = self.tableWidget_2.rowCount()
                self.tableWidget_2.insertRow(fila)
                self.tableWidget_2.setItem(fila , 0 , QTableWidgetItem(str(item[0])) ) #CANTIDAD
                self.tableWidget_2.setItem(fila , 1 , QTableWidgetItem(str(item[1])) )  #DESCRIPCION
                if self.tipo_doc == 'BOLETA':
                    neto = ( item[2] / 1.19 )
                    neto = round(neto,2)  #maximo 2 decimales  continuar modificando el json de crear orden trabajo
                    self.tableWidget_2.setItem(fila , 2 , QTableWidgetItem( str(neto) ) )  #VALOR NETO BOLETAS VIENEN CON IVA 
                else: # GUIAS Y FACTURAS NETOS VIENEN SI IVA
                    self.tableWidget_2.setItem(fila , 2 , QTableWidgetItem(str(item[2] )) )  #VALOR NETO FACTURAS VIENEN SIN IVA

    def agregar(self):
        if self.tableWidget_2.rowCount() <=16 :
            fila = self.tableWidget_2.rowCount()
            self.tableWidget_2.insertRow(fila)
        else:
            QMessageBox.about(self, 'ERROR', 'Ha alcanzado el limite maximo de filas. Intente crear otra Orden para continuar agregando items.')

    def eliminar(self):
        fila = self.tableWidget_2.currentRow()  #FILA SELECCIONADA , retorna -1 si no se selecciona una fila
        if fila != -1:
            #print('Eliminando la fila ' + str(fila))
            self.tableWidget_2.removeRow(fila)
        else: 
            QMessageBox.about(self,'Consejo', 'Seleccione una fila para eliminar')

    def registrar_orden(self):
        nombre = self.nombre_2.text()
        telefono = self.telefono_2.text()
        cont = self.contacto_2.text()
        oce = self.oce_2.text()
        fecha_orden = datetime.now().date() #FECHA DE CREACION DE ORDEN

        fecha = self.fecha.date()  #FECHA ESTIMADA
        fecha = fecha.toPyDate()
        lineas_totales = 0
        if nombre != '':
            if telefono != '':
                try:
                    telefono = int(telefono)
                    #continuar testeando con el analisis.py
                    cant = self.tableWidget_2.rowCount()
                    print(f'cantidad de datos: {cant}')
                    vacias = False
                    correcto = True
                    cantidades = []
                    descripciones = []
                    valores_neto = []
                    i = 0
                    while i< cant:
                        cantidad = self.tableWidget_2.item(i,0) #Collumna cantidades
                        descripcion = self.tableWidget_2.item(i,1) #Columna descripcion
                        neto = self.tableWidget_2.item(i,2) #Columna de valor neto
                        if cantidad != None and descripcion != None and neto != None :  #Checkea si se creo una fila, esta no este vacia.
                            if cantidad.text() != '' and descripcion.text() != '' and neto.text() != '' :  #Chekea si se modifico una fila, esta no este vacia
                                try: 
                                    nueva_cant = cantidad.text().replace(',','.',3)
                                    nuevo_neto = neto.text().replace(',','.',3)
                                    cantidades.append( float(nueva_cant) )
                                    descripciones.append(descripcion.text())
                                    lineas = self.separar(descripcion.text())
                                    lineas_totales = lineas_totales + len(lineas)
                                    valores_neto.append(float(nuevo_neto))

                                except ValueError:
                                    correcto = False
                            else:
                                vacias=True
                        else:
                            vacias = True
                        i+=1
                    if vacias:
                        QMessageBox.about(self, 'Alerta' ,'Una fila y/o columna esta vacia, rellenela para continuar' )
                    elif lineas_totales > 14:
                        QMessageBox.about(self, 'Alerta' ,'Filas totales: '+str(lineas_totales) + ' - El maximo soportado por el formato de la orden es de 14 filas.' )
                    elif correcto == False:
                        QMessageBox.about(self,'Alerta', 'Se encontro un error en una de las cantidades o Valores neto ingresados. Solo ingrese numeros en dichos campos')
                    else:
                        formato = {
                            "cantidades" : cantidades,
                            "descripciones" : descripciones,
                            "valores_neto": valores_neto,
                            "creado_por" : self.datos_usuario[8]
                        }
                        detalle = json.dumps(formato)
                        try:
                            enchape = 'NO'
                            despacho = 'NO'

                            if self.r_despacho.isChecked():
                                despacho = 'SI'
                            
                            if self.r_dim.isChecked():

                                if self.r_enchape.isChecked():
                                    fecha = fecha + timedelta(days=2)
                                    enchape = 'SI'

                                self.conexion.root.registrar_orden_dimensionado( self.inter , str(self.fecha_venta), nombre , telefono, str(fecha) , detalle, self.tipo_doc, self.nro_doc,enchape,despacho,str(fecha_orden),cont,oce,self.vendedor)
                                resultado = self.conexion.root.buscar_orden_dim_interno(self.inter)
                                self.nro_orden = self.buscar_nro_orden(resultado)
                                # crear vinculo a documento de venta
                                self.crear_vinculo('dimensionado')
                                datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, enchape, cont,oce,self.vendedor)
                                self.crear_pdf(datos,'dimensionado',despacho)

                                boton = QMessageBox.question(self, 'Orden de dimensionado registrada correctamente', 'Desea abrir la Orden?')
                                if boton == QMessageBox.Yes:
                                    self.ver_pdf('dimensionado')
                            elif self.r_elab.isChecked():
                                self.conexion.root.registrar_orden_elaboracion( nombre,telefono,str(fecha_orden), str(fecha),self.nro_doc,self.tipo_doc,cont,oce, despacho, self.inter ,detalle, str(self.fecha_venta), self.vendedor)
                                resultado = self.conexion.root.buscar_orden_elab_interno(self.inter)
                                self.nro_orden = self.buscar_nro_orden(resultado)
                                self.crear_vinculo('elaboracion')
                                datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, 'NO', cont, oce, self.vendedor)
                                self.crear_pdf(datos , 'elaboracion',despacho)
                                boton = QMessageBox.question(self, 'Orden de elaboracion registrada correctamente', 'Desea abrir la Orden?')
                                if boton == QMessageBox.Yes:
                                    self.ver_pdf('elaboracion')
                            elif self.r_carp.isChecked():
                                self.conexion.root.registrar_orden_carpinteria( nombre,telefono,str(fecha_orden), str(fecha),self.nro_doc,self.tipo_doc,cont,oce, despacho, self.inter ,detalle, str(self.fecha_venta), self.vendedor)
                                resultado = self.conexion.root.buscar_orden_carp_interno(self.inter)
                                self.nro_orden = self.buscar_nro_orden(resultado)
                                self.crear_vinculo('carpinteria')
                                datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, 'NO', cont, oce, self.vendedor)
                                self.crear_pdf(datos , 'carpinteria',despacho)
                                boton = QMessageBox.question(self, 'Orden de elaboracion registrada correctamente', 'Desea abrir la Orden?')
                                if boton == QMessageBox.Yes:
                                    self.ver_pdf('carpinteria')
                            elif self.r_pall.isChecked():
                                self.conexion.root.registrar_orden_pallets( nombre,telefono,str(fecha_orden), str(fecha),self.nro_doc,self.tipo_doc,cont,oce, despacho, self.inter ,detalle, str(self.fecha_venta), self.vendedor)
                                resultado = self.conexion.root.buscar_orden_pall_interno(self.inter)
                                self.nro_orden = self.buscar_nro_orden(resultado)
                                self.crear_vinculo('pallets')
                                datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, 'NO', cont, oce, self.vendedor)
                                self.crear_pdf( datos , 'pallets',despacho)
                                boton = QMessageBox.question(self, 'Orden de elaboracion registrada correctamente', 'Desea abrir la Orden?')
                                if boton == QMessageBox.Yes:
                                    self.ver_pdf('pallets')
                            else:
                                QMessageBox.about(self, 'ERROR', 'Seleccione un tipo de orden a generar, antes de proceder a registrar')    

                        except EOFError:
                            self.conexion_perdida()
                        except AttributeError:
                            QMessageBox.about(self,'IMPORTANTE', 'Este mensaje se debe a que hubo un error al ingresar los datos a la base de datos. Contacte con el soporte')


                except ValueError:
                    QMessageBox.about(self, 'ERROR', 'Solo ingrese numeros en el campo "Telefono" ')          
            else:
                QMessageBox.about(self, 'Sugerencia', 'Ingrese un telefono antes de continuar')           
        else:
            QMessageBox.about(self, 'Sugerencia', 'Ingrese un nombre antes de continuar')

    def crear_vinculo(self,tipo):
        vinc = {
            "tipo" : tipo,
            "folio" : self.nro_orden }
        vinculo = json.dumps(vinc)
        # BUSCA SI EXISTEN VINCULOS.
        if self.conexion.root.añadir_vinculo_orden_a_venta(self.tipo_doc, vinculo,self.nro_doc):
            print('vinculo de orden de trabajo a venta creado exitosamente')
        else:
            print('no se encontro el documento o sucedio un error')
        
# ------ Funciones de buscar orden de trabajo -------------
    def inicializar_buscar_orden(self):
        self.dateEdit.setDate(datetime.now().date())
        self.dateEdit.setCalendarPopup(True)
        self.r_fecha.setChecked(True)
        self.tb_buscar_orden.setRowCount(0)
        self.tb_buscar_orden.setColumnWidth(0,70)
        self.tb_buscar_orden.setColumnWidth(1,70)
        self.tb_buscar_orden.setColumnWidth(2,100)
        self.tb_buscar_orden.setColumnWidth(3,170)
        self.tb_buscar_orden.setColumnWidth(4,120)
        self.tb_buscar_orden.setColumnWidth(5,100)
        self.tb_buscar_orden.setColumnWidth(6,100)

        self.stackedWidget.setCurrentWidget(self.buscar_orden)

    def buscar_dimensionado(self):
        self.ch_nulas.setChecked(False)
        self.lb_tipo_orden.setText('DIMENSIONADO')
        estado = 'Ninguno'
        if self.conexion:
            self.tb_buscar_orden.setRowCount(0)
            if self.r_orden.isChecked(): #Busqueda por numero interno
                try:
                    orden = int(self.txt_orden.text())
                    consulta = self.conexion.root.buscar_orden_dim_numero(orden)
                    if consulta != None :
                        items = []
                        if consulta[19]:
                            estado = 'ANULADA'
                        else:
                            estado = 'VALIDA'
                        #campos: (nro_orden, nro_interno,fecha_creacion,nombre_cliente,fecha_venta,fecha_estimada,estado del doc)
                        item = (str(consulta[0]), str(consulta[1]),  str(consulta[11]), consulta[3] , str(consulta[2]) ,str(consulta[5]), estado )
                        items.append(item)
                        self.mostrar_en_tabla(items)
                    else:
                        QMessageBox.about(self,'Busqueda' ,'Orden de dimensionado NO encontrada')
                except ValueError:
                    QMessageBox.about(self,'ERROR' ,'Ingrese solo numeros')
                except EOFError:
                    QMessageBox.about(self,'ERROR','Se perdio la conexion con el servidor')

            elif self.r_fecha.isChecked():
                date = self.dateEdit.date()
                aux = date.toPyDate() 
                try:
                    datos = self.conexion.root.buscar_orden_dim_fecha( str(aux) )
                    if datos != ():
                        items = []
                        for dato in datos:
                            if dato[6]:
                                estado = 'ANULADA'
                            else:
                                estado = 'VALIDA'
                            #campos: (nro_orden, nro_interno,fecha_creacion,nombre_cliente,fecha_venta,fecha_estimada,estado del doc)
                            item = (str(dato[0]), str(dato[1]),  str(dato[4]), dato[3] , str(dato[2]) ,str(dato[5]), estado )
                            items.append(item)
                        self.mostrar_en_tabla(items)
                    else:
                        QMessageBox.about(self ,'Resultado', 'No se encontraron Ordenes de dimensionado')
                except EOFError:
                    QMessageBox.about(self,'ERROR','Se perdio la conexion con el servidor')
            elif self.r_cliente.isChecked():
                nombre_cliente = self.txt_cliente.text()
                print(nombre_cliente)
                try:
                    datos = self.conexion.root.buscar_orden_nombre('dimensionado', nombre_cliente)
                    if datos != ():
                        self.mostrar_en_tabla(datos)
                    else:
                        QMessageBox.about(self ,'Resultado', 'No se encontraron Ordenes de dimensionado coincidentes')
                except EOFError:
                    self.conexion_perdida()
        else:
            QMessageBox.about(self ,'Conexion', 'Se perdio la conexion')

    def busqueda_general(self,tipo):
        self.ch_nulas.setChecked(False)
        self.lb_tipo_orden.setText(tipo)
        if self.conexion:
            self.tb_buscar_orden.setRowCount(0)
            if self.r_orden.isChecked(): #Busqueda por numero de orden
                try:
                    orden = int(self.txt_orden.text())
                    if tipo == 'PALLETS':
                        consulta = self.conexion.root.buscar_orden_pall_numero(orden)
                    elif tipo == 'CARPINTERIA':
                        consulta = self.conexion.root.buscar_orden_carp_numero(orden)
                    elif tipo == 'ELABORACION':
                        consulta = self.conexion.root.buscar_orden_elab_numero(orden)
                    items = []
                    estado = 'Ninguno'
                    if consulta != None :
                        if consulta[16]:
                            estado = 'ANULADA'
                        else:
                            estado = 'VALIDA'
                        #campos: (nro_orden, nro_interno,fecha_creacion,nombre_cliente,fecha_venta,fecha_estimada,estado del doc)
                        item = (str(consulta[0]), str(consulta[10]),  str(consulta[3]), consulta[1] , str(consulta[12]) ,str(consulta[4]), estado )
                        items.append(item)
                        self.mostrar_en_tabla(items)
                    else:
                        QMessageBox.about(self,'Busqueda' ,'Orden de: '+ tipo+' NO encontrada')
                except ValueError:
                    QMessageBox.about(self,'ERROR' ,'Ingrese solo numeros')
                except EOFError:
                    QMessageBox.about(self,'ERROR','Se perdio la conexion con el servidor')

            elif self.r_fecha.isChecked(): #busqueda por fecha
                date = self.dateEdit.date()
                aux = date.toPyDate()
                #inicio = str(aux) + ' ' + '00:00:00'
                #fin = str(aux) + ' ' + '23:59:59'    
                try:
                    if tipo == 'PALLETS':
                        datos = self.conexion.root.buscar_orden_pall_fecha(str(aux))
                    elif tipo == 'CARPINTERIA':
                        datos = self.conexion.root.buscar_orden_carp_fecha(str(aux))
                    elif tipo == 'ELABORACION':
                        datos = self.conexion.root.buscar_orden_elab_fecha(str(aux))
                    items = []
                    estado = 'Ninguno'
                    if datos != ():
                        for dato in datos:
                            if dato[6]:
                                estado = 'ANULADA'
                            else:
                                estado = 'VALIDA'
                            #campos: (nro_orden, nro_interno,fecha_creacion,nombre_cliente,fecha_venta,fecha_estimada,estado del doc)
                            item = (str(dato[0]), str(dato[1]),  str(dato[2]), dato[3] , str(dato[4]) ,str(dato[5]), estado )
                            items.append(item)
                        self.mostrar_en_tabla(items)

                    else:
                        QMessageBox.about(self ,'Resultado', 'No se encontraron Ordenes de pallets')
                except EOFError:
                    QMessageBox.about(self,'ERROR','Se perdio la conexion con el servidor')

            elif self.r_cliente.isChecked():
                nombre_cliente = self.txt_cliente.text()
                try:
                    datos = self.conexion.root.buscar_orden_nombre( tipo.lower() , nombre_cliente)
                    items = []
                    estado = 'Ninguno'
                    if datos != ():
                        for dato in datos:
                            if dato[6]:
                                estado = 'ANULADA'
                            else:
                                estado = 'VALIDA'
                            #campos: (nro_orden, nro_interno,fecha_creacion,nombre_cliente,fecha_venta,fecha_estimada,estado del doc)
                            item = (str(dato[0]), str(dato[1]),  str(dato[2]), dato[3] , str(dato[4]) ,str(dato[5]), estado )
                            items.append(item)
                        self.mostrar_en_tabla(items)
                    else:
                        QMessageBox.about(self ,'Resultado', 'No se encontraron Ordenes de: '+ tipo +' coincidentes')
                except EOFError:
                    self.conexion_perdida()
        else:
            QMessageBox.about(self ,'Conexion', 'Se perdio la conexion')

    def mostrar_en_tabla(self,datos):
        for dato in datos:
            fila = self.tb_buscar_orden.rowCount()
            self.tb_buscar_orden.insertRow(fila)
            self.tb_buscar_orden.setItem(fila , 0 , QTableWidgetItem( str(dato[0])) )  #nro orden
            self.tb_buscar_orden.setItem(fila , 1 , QTableWidgetItem( str(dato[1]) ))  #nro interno
            self.tb_buscar_orden.setItem(fila , 2 , QTableWidgetItem( str(dato[2]) ) ) #fecha creacion
            self.tb_buscar_orden.setItem(fila , 3 , QTableWidgetItem(dato[3]))       #nombre
            self.tb_buscar_orden.setItem(fila , 4 , QTableWidgetItem( str(dato[4]) ) )       #fecha venta
            self.tb_buscar_orden.setItem(fila , 5 , QTableWidgetItem( str(dato[5]) ) )       #fecha estimada
            self.tb_buscar_orden.setItem(fila , 6 , QTableWidgetItem(dato[6]) )       #estado
            
#------ Funciones para modificar orden ----------- 

    def inicializar_modificar_orden(self):
        seleccion = self.tb_buscar_orden.currentRow()
        if seleccion != -1:
            _item = self.tb_buscar_orden.item( seleccion, 0) 
            if _item:            
                orden = self.tb_buscar_orden.item(seleccion, 0 ).text()
                self.nro_orden = int(orden)
                tipo = self.lb_tipo_orden.text()
                print('ventana modificar orden ... para: '+ tipo + ' -->' + str(self.nro_orden))
                self.tb_modificar_orden.setColumnWidth(0,80)
                self.tb_modificar_orden.setColumnWidth(1,430)
                self.tb_modificar_orden.setColumnWidth(2,85)
                self.rellenar_datos_orden(tipo)
                self.stackedWidget.setCurrentWidget(self.modificar_orden)
        
        else:
            QMessageBox.about(self,'ERROR', 'Seleccione un Nro interno antes de continuar')

    def rellenar_datos_orden(self,tipo):
        self.r_despacho_5.setChecked(False)
        self.date_venta_5.setCalendarPopup(True)
        self.fecha_5.setCalendarPopup(True)
        cantidades = None
        descripciones = None
        enchapado = 'NO'
        despacho = 'NO'
        self.date_venta_5.setDate( datetime.now())
        self.tipo = tipo
        self.tb_modificar_orden.setRowCount(0)
        self.comboBox_5.clear()
     
        if tipo == 'DIMENSIONADO':
            try:
                resultado = self.conexion.root.buscar_orden_dim_numero(self.nro_orden)
                if resultado != None :

                    if resultado[18]: #si es manual
                        self.manual = True
                        if resultado[8]:
                            self.txt_nro_doc_5.setText( str(resultado[8]) )       #NUMERO DOCUMENTO
                            self.nro_doc = int( resultado[8] )

                        else:
                            self.txt_nro_doc_5.setText( '0' )       #NUMERO DOCUMENTO

                        if resultado[7]:
                            if resultado[7] == 'BOLETA':
                                self.comboBox_5.addItem( resultado[7] )                 #TIPO DOCUMENTO
                                #self.comboBox_5.addItem('FACTURA')
                                self.tipo_doc = resultado[7]

                            elif resultado[7] == 'FACTURA':
                                self.comboBox_5.addItem( resultado[7] )                 #TIPO DOCUMENTO
                                #self.comboBox_5.addItem('BOLETA')
                                self.tipo_doc = resultado[7]
                            elif resultado[7] == 'GUIA':
                                self.comboBox_5.addItem(resultado[7])
                                self.tipo_doc = resultado[7]
                        else:
                            print('no existe su tipo doc')
                            self.comboBox_5.addItem('NO ASIGNADO') 
                            self.comboBox_5.addItem('FACTURA')  
                            self.comboBox_5.addItem('BOLETA')
                            self.comboBox_5.addItem('GUIA')
                        if resultado[2]:
                            aux3 = datetime.fromisoformat(str(resultado[2]) )
                            self.date_venta_5.setDate( aux3 )   #FECHA DE VENTA
                        if resultado[15]:
                            self.vendedor = resultado[14]   #VENDEDOR
                            self.txt_vendedor_5.setText( resultado[15] )             
                    else: #SI NO ES MANUAL
                        self.manual = False
                        self.txt_interno_5.setEnabled(False) #solo lectura
                        self.txt_nro_doc_5.setEnabled(False)
                        self.date_venta_5.setEnabled(False)
                        self.comboBox_5.setEnabled(False)

                        self.comboBox_5.addItem( resultado[7] )    #tipo documento
                        self.tipo_doc = resultado[7]
                        aux1 = datetime.fromisoformat(str(resultado[2]) )
                        self.date_venta_5.setDate( aux1 )     #fecha de venta
                        self.txt_nro_doc_5.setText( str(resultado[8]) ) #numero de documento
                        self.nro_doc = resultado[8]
                        self.vendedor = resultado[15]    #vendedor
                        self.txt_vendedor_5.setText( resultado[15] ) #vendedor

                    self.txt_vendedor_5.setEnabled(False)
                    self.txt_interno_5.setText(str( resultado[1] )) #INTERNO
                    self.inter = str( resultado[1] )
                    
                    self.nombre_5.setText( resultado[3] )   #nombre
                    self.telefono_5.setText( str(resultado[4]) ) #telefono
                    aux = datetime.fromisoformat(str(resultado[5]) )
                    self.fecha_5.setDate( aux )  #FECHA ESTIMADA           
                    detalle = json.loads(resultado[6])
                    
                    #self.lb_planchas.setText( str(detalle["total_planchas"]) )
                    cantidades = detalle["cantidades"]
                    descripciones = detalle["descripciones"]
                    valores_neto = detalle["valores_neto"]
                    j = 0
                    while j < len( cantidades ):
                        fila = self.tb_modificar_orden.rowCount()
                        self.tb_modificar_orden.insertRow(fila)
                        self.tb_modificar_orden.setItem(fila , 0 , QTableWidgetItem( str( cantidades[j] )) ) 
                        self.tb_modificar_orden.setItem(fila , 1 , QTableWidgetItem( descripciones[j] ) )
                        self.tb_modificar_orden.setItem(fila , 2 , QTableWidgetItem( str( valores_neto[j] )) ) 
                        j+=1      
                    if resultado[9] == 'SI' :
                        self.r_enchape_5.setChecked(True)      #enchape
                        enchapado = 'SI'
                    if resultado[10] == 'SI' :
                        self.r_despacho_5.setChecked(True)  #despacho
                        despacho = 'SI'

                    self.lb_fecha_orden_5.setText( str(resultado[11]))
                    self.fecha_orden = datetime.fromisoformat( str(resultado[11]) )  #Fecha orden
                    self.contacto_5.setText( resultado[12] ) #contacto
                    self.oce_5.setText( resultado[13] )      #orden comprar
                    

            except EOFError:
                self.conexion_perdida()

        else: 
            try:
                if tipo == 'ELABORACION':
                    print('buscando elaboracion datos...')
                    resultado = self.conexion.root.buscar_orden_elab_numero(self.nro_orden)
                elif tipo == 'CARPINTERIA':
                    print('buscando carpinteria datos...')
                    resultado = self.conexion.root.buscar_orden_carp_numero(self.nro_orden)
                elif tipo == 'PALLETS':
                    print('buscando pallets datos...')
                    resultado = self.conexion.root.buscar_orden_pall_numero(self.nro_orden)
                
                
                if resultado != None :
                    if resultado[ 15 ]: #si es manual
                        self.manual = True
                        if resultado[5]:
                            self.txt_nro_doc_5.setText( str(resultado[5]) )       #NUMERO DOCUMENTO
                            self.nro_doc = int( resultado[5] )
                        else:
                            self.txt_nro_doc_5.setText( '0' )       #NUMERO DOCUMENTO

                        if resultado[6]:
                            if resultado[6] == 'BOLETA':
                                self.comboBox_5.addItem( resultado[6] )                 #TIPO DOCUMENTO
                                #self.comboBox_5.addItem('FACTURA')
                                self.tipo_doc = resultado[6]
                            elif resultado[6] == 'FACTURA':
                                self.comboBox_5.addItem( resultado[6] )                 #TIPO DOCUMENTO
                                #self.comboBox_5.addItem('BOLETA')
                                self.tipo_doc = resultado[6]
                            elif resultado[6] == 'GUIA':
                                self.comboBox_5.addItem(resultado[6])
                                self.tipo_doc = resultado[6]
                        else:
                            self.comboBox_5.addItem('NO ASIGNADO')  
                            self.comboBox_5.addItem('FACTURA')  
                            self.comboBox_5.addItem('BOLETA')
                            self.comboBox_5.addItem('GUIA')

                        if resultado[12]:
                            aux3 = datetime.fromisoformat(str(resultado[12]) )
                            self.date_venta_5.setDate( aux3 )   #FECHA DE VENTA
                        if resultado[14]:
                            self.vendedor = resultado[14]   #VENDEDOR
                            self.txt_vendedor_5.setText( resultado[14] )

                    else: #si no es manual
                        self.manual = False
                        self.txt_interno_5.setEnabled(False) #solo lectura
                        self.txt_vendedor_5.setEnabled(False)
                        self.txt_nro_doc_5.setEnabled(False)
                        self.date_venta_5.setEnabled(False)
                        self.comboBox_5.setEnabled(False)

                        self.txt_nro_doc_5.setText( str(resultado[5]) )       #NUMERO DOCUMENTO
                        self.nro_doc = resultado[5]
                        self.comboBox_5.addItem( resultado[6] )                 #TIPO DOCUMENTO
                        self.tipo_doc = resultado[6]
                        aux3 = datetime.fromisoformat(str(resultado[12]) )
                        self.date_venta_5.setDate( aux3 )   #FECHA DE VENTA
                        self.vendedor = resultado[14]   #VENDEDOR
                        self.txt_vendedor_5.setText( resultado[14] )

                    self.txt_interno_5.setText(str( resultado[10] ))        #NRO INTERNO
                    self.inter = str( resultado[10] )
                    self.r_enchape_5.setVisible(False)
                    self.nombre_5.setText( resultado[1] )         #NOMBRE CLIENTE
                    self.telefono_5.setText( str(resultado[2]) )  #TELEFONO

                    self.fecha_orden = datetime.fromisoformat( str(resultado[3]) )  
                    self.lb_fecha_orden_5.setText( str(resultado[3]))   #FECHA DE ORDEN
                    
                    f_estimada = datetime.fromisoformat(str(resultado[4]) )
                    self.fecha_5.setDate( f_estimada )                   #FECHA ESTIMADA DE ENTREGA

                    
                    self.contacto_5.setText( resultado[7] )             #CONTACTO
                    self.oce_5.setText( resultado[8] )                   #ORDEN DE COMPRA
                              
                    if resultado[9] == 'SI' :
                        self.r_despacho_5.setChecked(True)               #DESPACHO
                        despacho = 'SI'


                    detalle = json.loads(resultado[11])                  #DETALLE
                    cantidades = detalle["cantidades"]
                    descripciones = detalle["descripciones"]
                    valores_neto = detalle["valores_neto"]
                    j = 0
                    while j < len( cantidades ):
                        fila = self.tb_modificar_orden.rowCount()
                        self.tb_modificar_orden.insertRow(fila)
                        self.tb_modificar_orden.setItem(fila , 0 , QTableWidgetItem( str( cantidades[j] )) ) 
                        self.tb_modificar_orden.setItem(fila , 1 , QTableWidgetItem( descripciones[j] ) )
                        self.tb_modificar_orden.setItem(fila , 2 , QTableWidgetItem( str( valores_neto[j] )) ) 
                        j+=1    
            except EOFError:
                self.conexion_perdida()

        # ------ comprobar si existe el PDF en los archivos locales -------------
        
        if tipo == "DIMENSIONADO":
            abrir = self.carpeta + '/ordenes/' + 'dimensionado_' +str(self.nro_orden) + '.pdf'
            tipo = 'dimensionado'
        elif tipo == "ELABORACION":
            abrir = self.carpeta + '/ordenes/'  + 'elaboracion_' +str(self.nro_orden) + '.pdf'
            tipo = 'elaboracion'
        elif tipo == "CARPINTERIA":
            abrir = self.carpeta + '/ordenes/'  + 'carpinteria_' +str(self.nro_orden) + '.pdf'
            tipo = 'carpinteria'
        elif tipo == "PALLETS":
            abrir = self.carpeta + '/ordenes/'  + 'pallets_' +str(self.nro_orden) + '.pdf'
            tipo = 'pallets'
        print('verificando existencia del PDF ....')
        if not os.path.isfile(abrir):
            print(self.tipo_doc)
            datos = ( str(self.nro_orden) ,str(self.fecha_orden.strftime("%d-%m-%Y")),self.nombre_5.text(),self.telefono_5.text(), str((self.fecha_5.date().toPyDate()).strftime("%d-%m-%Y")),cantidades,descripciones,enchapado ,self.contacto_5.text(),self.oce_5.text() ,self.vendedor )
            self.crear_pdf(datos, tipo ,despacho)
            print('pdf no encontrado, pero se acaba de crear')
        else:
            print('El pdf si existe localmente')

    def agregar_2(self):
        if self.tb_modificar_orden.rowCount() <=16 :
            fila = self.tb_modificar_orden.rowCount()
            self.tb_modificar_orden.insertRow(fila)
        else:
            QMessageBox.about(self, 'ERROR', 'Ha alcanzado el limite maximo de filas. Intente crear otra Orden para continuar agregando items.')
    def eliminar_2(self):
        fila = self.tb_modificar_orden.currentRow()  #FILA SELECCIONADA , retorna -1 si no se selecciona una fila
        if fila != -1:
            #print('Eliminando la fila ' + str(fila))
            self.tb_modificar_orden.removeRow(fila)
        else: 
            QMessageBox.about(self,'Consejo', 'Seleccione una fila para eliminar')
    def actualizar_orden(self):
        nombre = self.nombre_5.text()
        telefono = self.telefono_5.text()
        contacto = self.contacto_5.text()
        oce = self.oce_5.text()
        fecha = self.fecha_5.date()  #fecha estimada
        fecha = fecha.toPyDate()
        enchapado = 'NO'
        despacho = 'NO'

        self.tipo_doc = self.comboBox_5.currentText()
        self.vendedor = self.txt_vendedor_5.text()
        fecha_venta = str( self.date_venta_5.date().toPyDate() )
        lineas_totales = 0 
        if nombre != '':
            if telefono != '':
                try:
                    telefono = int(telefono)
                    self.inter = int( self.txt_interno_5.text() )
                    self.nro_doc = int( self.txt_nro_doc_5.text() )
                    
                    cant = self.tb_modificar_orden.rowCount() #cantidad de items de la tabla
                    vacias = False
                    correcto = True
                    cantidades = []
                    descripciones = []
                    valores_neto = []
                    i = 0
                    while i< cant:
                        cantidad = self.tb_modificar_orden.item(i,0) #Collumna cantidades
                        descripcion = self.tb_modificar_orden.item(i,1) #Columna descripcion
                        neto = self.tb_modificar_orden.item(i,2) #Columna de valor neto
                        if cantidad != None and descripcion != None and neto != None :  #Checkea si se creo una fila, esta no este vacia.
                            if cantidad.text() != '' and descripcion.text() != '' and neto.text() != '' :  #Chekea si se modifico una fila, esta no este vacia
                                try: 
                                    nueva_cant = cantidad.text().replace(',','.',3)
                                    nuevo_neto = neto.text().replace(',','.',3)
                                    cantidades.append( float(nueva_cant) )
                                    descripciones.append(descripcion.text())
                                    lineas = self.separar(descripcion.text())
                                    lineas_totales = lineas_totales + len(lineas)
                                    valores_neto.append(float(nuevo_neto))

                                except ValueError:
                                    correcto = False
                            else:
                                vacias=True
                        else:
                            vacias = True
                        i+=1
                    if vacias:
                        QMessageBox.about(self, 'Alerta' ,'Una fila y/o columna esta vacia, rellenela para continuar' )
                    elif lineas_totales > 14:
                        QMessageBox.about(self, 'Alerta' ,'Filas totales: '+str(lineas_totales) + ' - El maximo soportado por el formato de la orden es de 14 filas.' )
                    elif correcto == False:
                        QMessageBox.about(self,'Alerta', 'Se encontro un error en una de las cantidades o Valores neto ingresados. Solo ingrese numeros en dichos campos')
                    else:
                        formato = {
                            "cantidades" : cantidades,
                            "descripciones" : descripciones,
                            "valores_neto": valores_neto,
                            "creado_por" : self.datos_usuario[8]
                        }
                        detalle = json.dumps(formato)
                        #print('---------------------------------------------------')
                        try:
                           
                            if self.r_despacho_5.isChecked():
                                despacho = 'SI'
                            if self.tipo == 'DIMENSIONADO':
                                if self.r_enchape_5.isChecked():
                                    enchapado = 'SI'
                                
                                datos = ( str(self.nro_orden) , str(self.fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones,enchapado, contacto,oce, self.vendedor)
                                self.crear_pdf(datos,'dimensionado',despacho)

                                if self.conexion.root.actualizar_orden_dim( self.manual,self.inter,fecha_venta,self.tipo_doc,self.nro_doc,self.vendedor, self.nro_orden,nombre,telefono,str(fecha),detalle,despacho,enchapado,contacto,oce):
                                    QMessageBox.about(self,'EXITO','La orden de dimensionado fue ACTUALIZADA correctamente')
                                else:
                                    QMessageBox.about(self,'ERROR','La orden de dimensionado NO se actualizo, porque no se modificaron los datos que ya existian.')

                            elif self.tipo == 'ELABORACION':
                                datos = ( str(self.nro_orden) , str(self.fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones,enchapado, contacto,oce, self.vendedor)
                                self.crear_pdf(datos,'elaboracion',despacho)

                                if self.conexion.root.actualizar_orden_elab( self.manual,self.inter,fecha_venta,self.tipo_doc,self.nro_doc,self.vendedor,  nombre,telefono, str(fecha), detalle, contacto, oce,despacho, self.nro_orden ):
                                    QMessageBox.about(self,'EXITO','La orden de elaboracion fue ACTUALIZADA correctamente')
                                else:
                                    QMessageBox.about(self,'ERROR','La orden de elaboracion NO se actualizo, porque no se modificaron los datos que ya existian.')
                            elif self.tipo == 'CARPINTERIA':
                                datos = ( str(self.nro_orden) , str(self.fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones,enchapado, contacto,oce, self.vendedor)
                                self.crear_pdf(datos,'carpinteria',despacho)

                                if self.conexion.root.actualizar_orden_carp(self.manual,self.inter,fecha_venta,self.tipo_doc,self.nro_doc,self.vendedor, nombre,telefono, str(fecha), detalle, contacto, oce,despacho, self.nro_orden ):
                                    QMessageBox.about(self,'EXITO','La orden de carpinteria fue ACTUALIZADA correctamente')
                                else:
                                    QMessageBox.about(self,'ERROR','La orden de carpinteria NO se actualizo, porque no se modificaron los datos que ya existian.')
                            elif self.tipo == 'PALLETS':
                                datos = ( str(self.nro_orden) , str(self.fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones,enchapado, contacto,oce, self.vendedor)
                                self.crear_pdf(datos,'pallets',despacho)

                                if self.conexion.root.actualizar_orden_pall(self.manual,self.inter,fecha_venta,self.tipo_doc,self.nro_doc,self.vendedor, nombre,telefono, str(fecha), detalle, contacto, oce,despacho, self.nro_orden ):
                                    QMessageBox.about(self,'EXITO','La orden de pallets fue ACTUALIZADA correctamente')
                                else:
                                    QMessageBox.about(self,'ERROR','La orden de pallets NO se actualizo, porque no se modificaron los datos que ya existian.')
    
                        except EOFError:
                            self.conexion_perdida()
                        except PermissionError:
                            QMessageBox.about(self,'ERROR' ,'La orden no se actualizo debido a que otro programa esta haciendo uso del archivo. Cierre dicho programa para continuar.')
                        #except AttributeError:
                         #   QMessageBox.about(self,'IMPORTANTE', 'Este mensaje se debe a que hubo un error al ingresar los datos a la base de datos. Contacte con el soporte')


                except ValueError:
                    if self.manual:
                        QMessageBox.about(self, 'ERROR', 'Solo ingrese numeros en el campo "Telefono" , "Numero Interno" y "Numero Documento" ') 
                    else:
                        QMessageBox.about(self, 'ERROR', 'Solo ingrese numeros en el campo "Telefono" ')          
            else:
                QMessageBox.about(self, 'Sugerencia', 'Ingrese un telefono antes de continuar')           
        else:
            QMessageBox.about(self, 'Sugerencia', 'Ingrese un nombre antes de continuar')
 # Funcion para generar REINGRESO -------------
    def inicializar_reingreso(self):
        seleccion = self.tb_buscar_orden.currentRow()
        if seleccion != -1:
            _item = self.tb_buscar_orden.item( seleccion, 0) 
            if _item:            
                orden = self.tb_buscar_orden.item(seleccion, 0 ).text()
                self.nro_orden = int(orden)
                tipo = self.lb_tipo_orden.text()
                print('ventana  reingreso ... para: '+ tipo + ' -->' + str(self.nro_orden))
                self.lb_folio_orden.setText(str(self.nro_orden))
                self.tb_reingreso_2.setRowCount(0)
                self.tb_reingreso_2.setColumnWidth(1,80)
                self.tb_reingreso_2.setColumnWidth(0,430)
                self.tb_reingreso_2.setColumnWidth(2,85)
                self.rellenar_datos_reingreso()
                self.stackedWidget.setCurrentWidget(self.reingreso)
        else:
            QMessageBox.about(self,'ERROR', 'Seleccione un Nro interno antes de continuar')

    def registrar_reingreso(self):
    
        tipo_doc = self.lb_doc_2.text()
        try:
            nro_doc = int( self.lb_documento_2.text() )
        except ValueError:
            nro_doc = None
            
        fecha = datetime.now().date()
        motivo = ''
        if self.r_cambio_2.isChecked():
            motivo = 'CAMBIO'
        elif self.r_devolucion_2.isChecked():
            motivo = 'DEVOLUCION'
        elif self.r_otro_2.isChecked():
            motivo = self.txt_otro_2.text()
        
        proceso = self.lb_proceso_2.text()
        descr = self.txt_descripcion_2.toPlainText()
        solucion = self.txt_solucion_2.toPlainText()
        lineas = 0

        if motivo != '' and descr != '' and solucion != '' :
            cant = self.tb_reingreso_2.rowCount()
            vacias = False #Determinna si existen campos vacios
            correcto = True #Determina si los datos estan escritos correctamente. campos cantidad y valor son numeros.
            cantidades = []
            descripciones = []
            valores_neto = []
            i = 0
            while i< cant:
                descripcion = self.tb_reingreso_2.item(i,0) #Collumna descripcion
                cantidad = self.tb_reingreso_2.item(i,1) #Columna cantidad
                neto = self.tb_reingreso_2.item(i,2) #Columna de valor neto
                if cantidad != None and descripcion != None and neto != None :  #Checkea si se creo una fila, esta no este vacia.
                    if cantidad.text() != '' and descripcion.text() != '' and neto.text() != '' :  #Chekea si se modifico una fila, esta no este vacia
                        try: 
                            nueva_cant = cantidad.text().replace(',','.',3)
                            nuevo_neto = neto.text().replace(',','.',3)
                            cantidades.append( float(nueva_cant) )
                            descripciones.append(descripcion.text())
                            print(descripcion.text())
                            linea = self.separar2( descripcion.text() , 60 )
                            lineas += len(linea)
                            print('total de lineas: '+ str(lineas))
                            valores_neto.append(float(nuevo_neto))

                        except ValueError:
                            correcto = False
                    else:
                        vacias=True
                else:
                    vacias = True
                i+=1
            if vacias:
                QMessageBox.about(self, 'Alerta' ,'Una fila y/o columna esta vacia, rellenela para continuar' )
            elif correcto == False:
                QMessageBox.about(self,'Alerta', 'Se encontro un error en una de las cantidades o Valores neto ingresados. Solo ingrese numeros en dichos campos')
            elif lineas > 4:
                QMessageBox.about(self, 'Alerta' ,'El maximo de filas por el formato de impresion es de 4.' )
            else:
                #print(cantidades)
                #print(valores_neto)
                formato = {
                        "cantidades" : cantidades,
                        "descripciones" : descripciones,
                        "valores_neto": valores_neto
                    }
                detalle = json.dumps(formato)
                '''print(fecha)
                print(str(type(fecha)))
                print(str(type(self.nro_orden)))
                print(fecha)
                print(tipo_doc)
                print(nro_doc)
                print(self.nro_orden)
                print(motivo)
                print(descr)
                print(proceso)       
                print(solucion)
                print(str(type(detalle)))'''
                print('procediendo a registrar reingreso ..')
                if self.conexion.root.registrar_reingreso( str(fecha), tipo_doc, nro_doc, self.nro_orden, motivo, descr, proceso, detalle,solucion):
                    resultado = self.conexion.root.obtener_max_reingreso()
                    self.nro_reingreso = resultado[0]
                    print('max nro reingreso: ' + str(resultado[0]) + ' de tipo: ' + str(type(resultado[0])))
                    datos = (resultado[0], str(fecha) , tipo_doc , nro_doc , motivo , descr , proceso , solucion, cantidades, descripciones, valores_neto)
                    self.crear_pdf_reingreso(datos)

                    boton = QMessageBox.question(self, 'Reingreso registrado correctamente', 'Desea ver el reingreso?')
                    if boton == QMessageBox.Yes:
                        self.ver_pdf_reingreso()
                else:
                    QMessageBox.about(self,'ERROR','404 NOT FOUND. Contacte con Don Huber ...problemas al registrar')

        else:
            QMessageBox.about(self,'Datos incompletos','Los campos "descripcion" , "solucion" son obligatiorios. Como tambien si selecciona "OTROS" debe rellenar su campo')
    def rellenar_datos_reingreso(self):
        
            self.lb_proceso_2.setText(self.lb_tipo_orden.text())

            if self.lb_tipo_orden.text() == 'DIMENSIONADO':
                try:
                    resultado = self.conexion.root.buscar_orden_dim_numero(self.nro_orden)
                    if resultado != None :

                        if resultado[18]: #si es manual
                            self.manual = True
                            if resultado[7]:
                                self.lb_doc_2.setText( resultado[7] ) #TIPO DOCUMENTO
                            else:
                                self.lb_doc_2.setText( 'No asignado' ) #TIPO DOCUMENTO
                            if resultado[8]:
                                self.lb_documento_2.setText( str(resultado[8]) ) #nro DOCUMENTO
                            else:
                                self.lb_documento_2.setText( 'No asignado' ) #nro DOCUMENTO
                        else:

                            self.lb_doc_2.setText( resultado[7] ) 
                            self.lb_documento_2.setText( str(resultado[8]) )

                        detalle = json.loads(resultado[6])

                        cantidades = detalle["cantidades"]
                        descripciones = detalle["descripciones"]
                        netos = detalle["valores_neto"]
                        j = 0
                        while j < len( cantidades ):
                            fila = self.tb_reingreso_2.rowCount()
                            self.tb_reingreso_2.insertRow(fila)
                            self.tb_reingreso_2.setItem(fila , 0 , QTableWidgetItem( descripciones[j] ) ) 
                            self.tb_reingreso_2.setItem(fila , 1 , QTableWidgetItem( str( cantidades[j] )  ) )
                            self.tb_reingreso_2.setItem(fila , 2 , QTableWidgetItem( str( netos[j] )  ) )
                            j+=1      
                                    
                except EOFError:
                    self.conexion_perdida()

            else:
                
                try:
                    if self.lb_tipo_orden.text() == 'ELABORACION':
                        resultado = self.conexion.root.buscar_orden_elab_numero(self.nro_orden)
                    elif self.lb_tipo_orden.text() == 'CARPINTERIA':
                        resultado = self.conexion.root.buscar_orden_carp_numero(self.nro_orden)
                    elif self.lb_tipo_orden.text() == 'PALLETS':
                        resultado = self.conexion.root.buscar_orden_pall_numero(self.nro_orden)

                    if resultado != None :
                        
                        if resultado[15]: #si es manual
                            self.manual = True

                            if resultado[6]:
                                self.lb_doc_2.setText( resultado[6] ) #TIPO DOCUMENTO
                            else:
                                self.lb_doc_2.setText( 'No asignado' ) #TIPO DOCUMENTO

                            if resultado[5]:
                                self.lb_documento_2.setText( str(resultado[5]) ) #nro DOCUMENTO
                            else:
                                self.lb_documento_2.setText( 'No asignado' ) #nro DOCUMENTO
                        else:                        
                            self.lb_doc_2.setText( resultado[6] )                 #TIPO DOCUMENTO
                            self.lb_documento_2.setText( str(resultado[5]) )       #NUMERO DOCUMENTO
                        
                                
                        detalle = json.loads(resultado[11])                  #DETALLE
                        cantidades = detalle["cantidades"]
                        descripciones = detalle["descripciones"]
                        netos = detalle["valores_neto"]
                        j = 0
                        while j < len( cantidades ):
                            fila = self.tb_reingreso_2.rowCount()
                            self.tb_reingreso_2.insertRow(fila)
                            self.tb_reingreso_2.setItem(fila , 0 , QTableWidgetItem( descripciones[j] ) ) 
                            self.tb_reingreso_2.setItem(fila , 1 , QTableWidgetItem( str( cantidades[j] )  ) )
                            self.tb_reingreso_2.setItem(fila , 2 , QTableWidgetItem( str( netos[j] )  ) )
                            j+=1     

                except EOFError:
                    self.conexion_perdida()

    def agregar_3(self):
        if self.tb_reingreso_2.rowCount() <=16 :
            fila = self.tb_reingreso_2.rowCount()
            self.tb_reingreso_2.insertRow(fila)
        else:
            QMessageBox.about(self, 'ERROR', 'Ha alcanzado el limite maximo de filas. Intente crear otro REINGRESO para continuar agregando items.')

    def eliminar_3(self):
        fila = self.tb_reingreso_2.currentRow()  #FILA SELECCIONADA , retorna -1 si no se selecciona una fila
        if fila != -1:
            #print('Eliminando la fila ' + str(fila))
            self.tb_reingreso_2.removeRow(fila)
        else: 
            QMessageBox.about(self,'Consejo', 'Seleccione una fila para eliminar')
 
# --------- Funciones de INGRESO MANUAL -----------------
    def inicializar_ingreso_manual(self):
        if self.datos_usuario[4] == 'SI' :
            self.rellenar_datos_manual()
            self.stackedWidget.setCurrentWidget(self.ingreso_manual)
        else:
            dialog = InputDialog2('CLAVE:', True ,'INGRESE CLAVE',self)
            if dialog.exec():
                clave = dialog.getInputs()
                try:
                    resultado = self.conexion.root.obtener_clave()
                    end = 0
                    for item in resultado:
                        if clave == item[0]:
                            self.clave = clave
                            self.rellenar_datos_manual()
                            self.stackedWidget.setCurrentWidget(self.ingreso_manual)
                            
                            end = 1
                    if end == 0:
                        QMessageBox.about(self,'ERROR' ,'CLAVE INVALIDA')

                except EOFError:
                    self.conexion_perdida()

    def rellenar_datos_manual(self):
        #-----datos de orden manual  ------
        self.r_uso_interno_1.setCheckState(False)
        self.r_facturar_1.setCheckState(False)
        self.r_enchape_1.setCheckState(False)
        self.r_despacho_1.setCheckState(False)
        self.nombre_1.setText('')
        self.telefono_1.setText('')
        self.contacto_1.setText('')
        self.oce_1.setText('')
        self.fecha_1.setDate( datetime.now())
        self.fecha_1.setCalendarPopup(True)
        self.tb_orden_manual.setRowCount(0)
        self.tb_orden_manual.setColumnWidth(0,80)
        self.tb_orden_manual.setColumnWidth(1,430)
        self.tb_orden_manual.setColumnWidth(2,85)
        #-----datos de reingreso manual  ------
        self.comboBox_6.clear()
        self.comboBox_6.addItem('No asignado')
        self.comboBox_6.addItem('BOLETA')
        self.comboBox_6.addItem('FACTURA')
        self.comboBox_6.addItem('GUIA')
        self.txt_orden_6.setText('0')
        self.txt_nro_doc_6.setText('0')
        self.txt_otro_6.setText('')
        self.txt_descripcion_6.clear()
        self.txt_solucion_6.clear()
        self.txt_descripcion_7.setText('')
        self.tb_reingreso_manual.setRowCount(0)
        self.tb_reingreso_manual.setColumnWidth(1,80)
        self.tb_reingreso_manual.setColumnWidth(0,430)
        self.tb_reingreso_manual.setColumnWidth(2,85)

    def registrar_orden_manual(self):
        nombre = self.nombre_1.text()     #NOMBRE CLIENTE
        telefono = self.telefono_1.text() #TELEFONO
        fecha = self.fecha_1.date()  #FECHA ESTIMADA
        fecha = fecha.toPyDate()
        #self.tipo_doc = self.comboBox.currentText() #TIPO DE DOCUMENTO
        self.tipo_doc = None
        interno = 0 #NRO INTERNO
        self.nro_doc = None #NRO DOCUMENTO
        f_venta = None #FECHA DE  VENTA
        #f_venta = self.fecha_venta.dateTime() #FECHA DE  VENTA
        #f_venta = f_venta.toPyDateTime()
        #vendedor = self.txt_vendedor.text() #VENDEDOR
        vendedor = self.datos_usuario[8]
        observacion = self.txt_obs_1.toPlainText()
        lineas_totales = 0
        if nombre != '' and telefono != '' and observacion != '' :
            
            
            try:
                telefono = int(telefono)
                cant = self.tb_orden_manual.rowCount()
                vacias = False
                correcto = True
                cantidades = []
                descripciones = []
                valores_neto = []
                i = 0
                while i< cant:
                    cantidad = self.tb_orden_manual.item(i,0) #Collumna cantidades
                    descripcion = self.tb_orden_manual.item(i,1) #Columna descripcion
                    neto = self.tb_orden_manual.item(i,2) #Columna de valor neto
                    if cantidad != None and descripcion != None and neto != None :  #Checkea si se creo una fila, esta no este vacia.
                        if cantidad.text() != '' and descripcion.text() != '' and neto.text() != '' :  #Chekea si se modifico una fila, esta no este vacia
                            try: 
                                nueva_cant = cantidad.text().replace(',','.',3)
                                nuevo_neto = neto.text().replace(',','.',3)
                                cantidades.append( float(nueva_cant) )
                                descripciones.append(descripcion.text())

                                lineas = self.separar(descripcion.text())
                                lineas_totales = lineas_totales + len(lineas)

                                valores_neto.append(float(nuevo_neto))

                            except ValueError:
                                correcto = False
                        else:
                            vacias=True
                    else:
                        vacias = True
                    i+=1
                print('LINEAS TOTALES: ' + str(lineas_totales))
                if vacias:
                    QMessageBox.about(self, 'Alerta' ,'Una fila y/o columna esta vacia, rellenela para continuar' )
                elif lineas_totales > 14:
                    QMessageBox.about(self, 'Alerta' ,'Filas totales: '+str(lineas_totales) + ' - El maximo soportado por el formato de la orden es de 14 filas.' )
                elif correcto == False:
                    QMessageBox.about(self,'Alerta', 'Se encontro un error en una de las cantidades o Valores neto ingresados. Solo ingrese numeros en dichos campos')
                else:
                    formato = {
                        "cantidades" : cantidades,
                        "descripciones" : descripciones,
                        "valores_neto": valores_neto,
                        "creado_por" : self.datos_usuario[8]
                    }
                    detalle = json.dumps(formato)
                    try:
                        enchape = 'NO'
                        despacho = 'NO'                        
                        if self.r_despacho_1.isChecked():
                            despacho = 'SI'
                        
                        oce = self.oce_1.text()
                        fecha_orden = datetime.now().date()
                        cont = self.contacto_1.text()

                        if self.r_dim_1.isChecked():
                            if self.r_enchape_1.isChecked():
                                fecha = fecha + timedelta(days=2)
                                enchape = 'SI'
                            self.conexion.root.registrar_orden_dimensionado( interno , f_venta , nombre , telefono, str(fecha) , detalle, self.tipo_doc, self.nro_doc ,enchape,despacho,str(fecha_orden),cont,oce, vendedor )
                            
                            resultado = self.conexion.root.buscar_orden_dim_interno(interno)
                            if self.clave:
                                self.conexion.root.eliminar_clave(self.clave)
                                self.clave = None
                            self.nro_orden = self.buscar_nro_orden(resultado)
                            self.conexion.root.actualizar_orden_dim_obser(observacion , self.nro_orden)

                            datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, enchape, cont,oce, vendedor)
                            self.crear_pdf(datos,'dimensionado', despacho )
                            boton = QMessageBox.question(self, 'Orden de dimensionado registrada correctamente', 'Desea abrir la Orden?')
                            if boton == QMessageBox.Yes:
                                self.ver_pdf('dimensionado')
                            self.stackedWidget.setCurrentWidget(self.inicio)
                            
                        elif self.r_elab_1.isChecked():
                            self.conexion.root.registrar_orden_elaboracion( nombre,telefono,str(fecha_orden), str(fecha),self.nro_doc,self.tipo_doc,cont,oce, despacho, interno ,detalle,f_venta,vendedor)
                            if self.clave:
                                self.conexion.root.eliminar_clave(self.clave)
                                self.clave = None
                            resultado = self.conexion.root.buscar_orden_elab_interno(interno)
                            self.nro_orden = self.buscar_nro_orden(resultado)
                            self.conexion.root.actualizar_orden_elab_obser(observacion , self.nro_orden)
                            datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, 'NO', cont, oce,vendedor)
                            self.crear_pdf(datos , 'elaboracion', despacho)
                            boton = QMessageBox.question(self, 'Orden de elaboracion registrada correctamente', 'Desea abrir la Orden?')
                            if boton == QMessageBox.Yes:
                                self.ver_pdf('elaboracion')
                            self.stackedWidget.setCurrentWidget(self.inicio)

                        elif self.r_carp_1.isChecked():
                            self.conexion.root.registrar_orden_carpinteria( nombre,telefono,str(fecha_orden), str(fecha),self.nro_doc,self.tipo_doc,cont,oce, despacho, interno ,detalle, f_venta ,vendedor)
                            if self.clave:
                                self.conexion.root.eliminar_clave(self.clave)
                                self.clave = None
                            resultado = self.conexion.root.buscar_orden_carp_interno(interno)
                            self.nro_orden = self.buscar_nro_orden(resultado)
                            self.conexion.root.actualizar_orden_carp_obser(observacion , self.nro_orden)
                            datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, 'NO', cont, oce, vendedor)
                            self.crear_pdf(datos , 'carpinteria', despacho)
                            boton = QMessageBox.question(self, 'Orden de elaboracion registrada correctamente', 'Desea abrir la Orden?')
                            if boton == QMessageBox.Yes:
                                self.ver_pdf('carpinteria')
                            self.stackedWidget.setCurrentWidget(self.inicio)

                        elif self.r_pall_1.isChecked():
                            self.conexion.root.registrar_orden_pallets( nombre,telefono,str(fecha_orden), str(fecha),self.nro_doc,self.tipo_doc,cont,oce, despacho, interno ,detalle, f_venta,vendedor)
                            if self.clave:
                                self.conexion.root.eliminar_clave(self.clave)
                                self.clave = None
                            resultado = self.conexion.root.buscar_orden_pall_interno(interno)
                            self.nro_orden = self.buscar_nro_orden(resultado)
                            self.conexion.root.actualizar_orden_pall_obser(observacion , self.nro_orden)
                            datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, 'NO', cont, oce,vendedor)
                            self.crear_pdf(datos , 'pallets',despacho)
                            boton = QMessageBox.question(self, 'Orden de elaboracion registrada correctamente', 'Desea abrir la Orden?')
                            if boton == QMessageBox.Yes:
                                self.ver_pdf('pallets')
                            self.stackedWidget.setCurrentWidget(self.inicio)

                        else:
                            QMessageBox.about(self, 'ERROR', 'Seleccione un tipo de orden a generar, antes de proceder a registrar')    

                    except EOFError:
                        self.conexion_perdida()   
                    #except AttributeError:

                     #   QMessageBox.about(self,'IMPORTANTE', 'Este mensaje se debe a que hubo un error al ingresar los datos a la base de datos. Contacte con el soporte')


            except ValueError:
                QMessageBox.about(self, 'ERROR', 'Solo ingrese Numeros en los campos "Telefono", "Numero de documento" y "Numero interno" ')          
        else:
            QMessageBox.about(self, 'Sugerencia', 'Los campos "Nombre" , "Telefono" y "Observacion" son obligatorios')         

    def agregar_4(self):
        if self.tb_orden_manual.rowCount() <=16 :
            fila = self.tb_orden_manual.rowCount()
            self.tb_orden_manual.insertRow(fila)
        else:
            QMessageBox.about(self, 'ERROR', 'Ha alcanzado el limite maximo de filas. Intente crear otro REINGRESO para continuar agregando items.')

    def eliminar_4(self):
        fila = self.tb_orden_manual.currentRow()  #FILA SELECCIONADA , retorna -1 si no se selecciona una fila
        if fila != -1:
            #print('Eliminando la fila ' + str(fila))
            self.tb_orden_manual.removeRow(fila)
        else: 
            QMessageBox.about(self,'Consejo', 'Seleccione una fila para eliminar')
    def buscar_descripcion(self):
        self.productos.clear()
        descr = self.txt_descripcion_1.text()
        try:
            resultado = self.conexion.root.buscar_prod_descr(descr)
            for item in resultado:
                self.productos.addItem(item[1])
                #print(resultado)
        except EOFError:
            self.conexion_perdida()
        
    def buscar_codigo(self):
        self.productos.clear()
        codigo = self.txt_codigo_1.text()
        try:
            resultado = self.conexion.root.buscar_prod_cod(codigo)
            for item in resultado:
                self.productos.addItem(item[1])
        except EOFError:
            self.conexion_perdida()
    def add_descripcion(self):
        descripcion = self.productos.currentText()

        if self.tb_orden_manual.rowCount() <=16 :
            fila = self.tb_orden_manual.rowCount()
            self.tb_orden_manual.insertRow(fila)
            self.tb_orden_manual.setItem(fila, 0 , QTableWidgetItem( '0' ))
            self.tb_orden_manual.setItem(fila, 1 , QTableWidgetItem( descripcion ))
            self.tb_orden_manual.setItem(fila, 2 , QTableWidgetItem( '0' ))
        else:
            QMessageBox.about(self, 'ERROR', 'Ha alcanzado el limite maximo de filas. Intente crear otra Orden para continuar agregando items.')
    def cambiar_observacion(self):
        if self.r_uso_interno_1.isChecked():
            obs = 'uso carpinteria'
            self.txt_obs_1.clear() 
            self.txt_obs_1.appendPlainText(obs)
        else:
            self.txt_obs_1.clear() 
 
    def reingreso_manual(self):
        nro_orden = self.txt_orden_6.text()           #NUMERO DE ORDEN
        tipo_doc = self.comboBox_6.currentText() #TIPO DE DOCUMENTO
        nro_doc = self.txt_nro_doc_6.text()        #NUMERO DE DOCUMENTO
        fecha = datetime.now().date()              #FECHA DE REINGRESO
        motivo = ''                                 #MOTIVO
        if self.r_cambio_6.isChecked():
            motivo = 'CAMBIO'
        elif self.r_devolucion_6.isChecked():
            motivo = 'DEVOLUCION'
        elif self.r_otro_6.isChecked():
            motivo = self.txt_otro_6.text()
        
        proceso = None
        if self.r_d_6.isChecked():
            proceso = 'DIMENSIONADO'
        elif self.r_e_6.isChecked():
            proceso = 'ELABORACION'
        elif self.r_c_6.isChecked():
            proceso = 'CARPINTERIA'
        elif self.r_p_6.isChecked():
            proceso = 'PALLETS'
        descr = self.txt_descripcion_6.toPlainText()   #descripcion
        solucion = self.txt_solucion_6.toPlainText()   #solucion
        lineas = 0
        if motivo != '' and descr != '' and solucion != '' :
            cant = self.tb_reingreso_manual.rowCount()
            vacias = False #Determinna si existen campos vacios
            correcto = True #Determina si los datos estan escritos correctamente. campos cantidad y valor son numeros.
            cantidades = []
            descripciones = []
            valores_neto = []
            i = 0
            while i< cant:
                descripcion = self.tb_reingreso_manual.item(i,0) #Collumna descripcion
                cantidad = self.tb_reingreso_manual.item(i,1) #Columna cantidad
                neto = self.tb_reingreso_manual.item(i,2) #Columna de valor neto
                if cantidad != None and descripcion != None and neto != None :  #Checkea si se creo una fila, esta no este vacia.
                    if cantidad.text() != '' and descripcion.text() != '' and neto.text() != '' :  #Chekea si se modifico una fila, esta no este vacia
                        try: 
                            nueva_cant = cantidad.text().replace(',','.',3)
                            nuevo_neto = neto.text().replace(',','.',3)
                            cantidades.append( float(nueva_cant) )
                            descripciones.append(descripcion.text())
                            linea = self.separar2(descripcion.text(), 60 )
                            lineas += len(linea)
                            print('total de lineas: '+ str(lineas))
                            valores_neto.append(float(nuevo_neto))

                        except ValueError:
                            correcto = False
                    else:
                        vacias=True
                else:
                    vacias = True
                i+=1
            if vacias:
                QMessageBox.about(self, 'Alerta' ,'Una fila y/o columna esta vacia, rellenela para continuar' )
            elif correcto == False:
                QMessageBox.about(self,'Alerta', 'Se encontro un error en una de las cantidades o Valores neto ingresados. Solo ingrese numeros en dichos campos')
            elif lineas > 4:
                QMessageBox.about(self, 'Alerta' ,'El maximo de filas por el formato de impresion es de 4.' )
            elif lineas < 1:
                QMessageBox.about(self, 'Alerta' ,'Como minimo ingrese 1 item.' )
            elif proceso == None:
                QMessageBox.about(self, 'Alerta' ,'Seleccione el proceso de la orden de trabajo antes de continuar' )
            else:
                #print(cantidades)
                #print(valores_neto)
                formato = {
                        "cantidades" : cantidades,
                        "descripciones" : descripciones,
                        "valores_neto": valores_neto,
                        "creado_por" : self.datos_usuario[8]
                    }
                detalle = json.dumps(formato)
                try:
                    nro_orden = int(nro_orden)
                    nro_doc = int(nro_doc)

                    if self.conexion.root.registrar_reingreso( str(fecha), tipo_doc, nro_doc, nro_orden, motivo, descr, proceso, detalle,solucion):
                        resultado = self.conexion.root.obtener_max_reingreso()
                        self.nro_reingreso = resultado[0]
                        print('max nro reingreso: ' + str(resultado[0]) + ' de tipo: ' + str(type(resultado[0])))
                        datos = (resultado[0], str(fecha) , tipo_doc , nro_doc , motivo , descr , proceso , solucion, cantidades, descripciones, valores_neto)
                        self.crear_pdf_reingreso(datos)
                        if self.clave:
                                self.conexion.root.eliminar_clave(self.clave)
                                self.clave = None

                        boton = QMessageBox.question(self, 'Reingreso registrado correctamente', 'Desea ver el reingreso?')
                        if boton == QMessageBox.Yes:
                            self.ver_pdf_reingreso()
                        self.stackedWidget.setCurrentWidget(self.inicio)
                    else:
                        QMessageBox.about(self,'ERROR','404 NOT FOUND. Contacte con Don Huber ...problemas al registrar')
                except ValueError:
                    QMessageBox.about(self,'ERROR','Ingresar solo numeros en los campos: "NUMERO DE ORDEN" y "NUMERO DE DOCUMENTO" ')
                
                except EOFError:
                    self.conexion_perdida()
        else:
            QMessageBox.about(self,'Datos incompletos','Los campos "descripcion" , "solucion" son obligatiorios. Como tambien si selecciona "OTROS" debe rellenar su campo')
    def agregar_6(self):
        if self.tb_reingreso_manual.rowCount() <=16 :
            fila = self.tb_reingreso_manual.rowCount()
            self.tb_reingreso_manual.insertRow(fila)
        else:
            QMessageBox.about(self, 'ERROR', 'Ha alcanzado el limite maximo de filas. Intente crear otro REINGRESO para continuar agregando items.')

    def eliminar_6(self):
        fila = self.tb_reingreso_manual.currentRow()  #FILA SELECCIONADA , retorna -1 si no se selecciona una fila
        if fila != -1:
            #print('Eliminando la fila ' + str(fila))
            self.tb_reingreso_manual.removeRow(fila)
        else: 
            QMessageBox.about(self,'Consejo', 'Seleccione una fila para eliminar')
    def buscar_descripcion_2(self):
        self.productos_6.clear()
        descr = self.txt_descripcion_7.text()
        try:
            resultado = self.conexion.root.buscar_prod_descr(descr)
            for item in resultado:
                self.productos_6.addItem(item[1])
        except EOFError:
            self.conexion_perdida()
    def add_descripcion_2(self):
        descripcion = self.productos_6.currentText()

        if self.tb_reingreso_manual.rowCount() < 4 :
            fila = self.tb_reingreso_manual.rowCount()
            self.tb_reingreso_manual.insertRow(fila)
            self.tb_reingreso_manual.setItem(fila, 0 , QTableWidgetItem( descripcion ))
            self.tb_reingreso_manual.setItem(fila, 1 , QTableWidgetItem( '0' ))
            self.tb_reingreso_manual.setItem(fila, 2 , QTableWidgetItem( '0' ))


        else:
            QMessageBox.about(self, 'ERROR', 'Ha alcanzado el limite maximo de filas soportado para la impresión de un reingreso. Intente hacer otro reingreso para los items faltantes ')

# --------- FUNCIONES DE INFORME ------------
    def inicializar_informe(self):
        self.tableWidget.setColumnWidth(0,371)
        self.vista_reingreso()
        self.actualizar()
        self.d_inicio.setDate( datetime.now().date() )
        self.d_inicio.setCalendarPopup(True)
        self.d_termino.setDate( datetime.now().date() )
        self.d_termino.setCalendarPopup(True)
        self.stackedWidget.setCurrentWidget(self.informes)

    def generar_informe(self):
        tipo_orden = self.comboBox.currentText()
        inicio = self.d_inicio.date()
        inicio = inicio.toPyDate()
        termino = self.d_termino.date()
        termino = termino.toPyDate()

        nombre =self.dir_informes + tipo_orden + '_'+ str( inicio.strftime("%d-%m-%Y")) + '_HASTA_' + str( termino.strftime("%d-%m-%Y")) + '.xlsx' 
        datos = None
        if self.r_orden_2.isChecked():
            try:
                if tipo_orden == 'DIMENSIONADO':
                    datos = self.conexion.root.informe_dimensionado(str(inicio) , str(termino))
                    self.informe_dimensionado(datos,nombre)
                    
                elif tipo_orden == 'ELABORACION':
                    datos = self.conexion.root.informe_elaboracion(str(inicio) , str(termino))
                    self.informe_generico(datos,nombre)
                elif tipo_orden == 'CARPINTERIA':
                    datos = self.conexion.root.informe_carpinteria(str(inicio) , str(termino))
                    self.informe_generico(datos,nombre)
                elif tipo_orden == 'PALLETS':
                    datos = self.conexion.root.informe_pallets(str(inicio) , str(termino))
                    self.informe_generico(datos,nombre)
                elif tipo_orden == 'REINGRESO':
                    if self.r_d.isChecked() or self.r_e.isChecked() or self.r_c.isChecked() or self.r_p.isChecked() :
                        acept = []
                        nombre =self.dir_informes + tipo_orden
                        entremedio = ''
                        if self.r_d.isChecked():
                            acept.append('DIMENSIONADO')
                            entremedio = entremedio + '_D'
                        if self.r_e.isChecked():
                            acept.append('ELABORACION')
                            entremedio = entremedio + '_E'
                        if self.r_c.isChecked():
                            acept.append('CARPINTERIA')
                            entremedio = entremedio + '_C'
                        if self.r_p.isChecked():
                            acept.append('PALLETS')
                            entremedio = entremedio + '_P'
                        fin = '_'+ str( inicio.strftime("%d-%m-%Y")) + '_HASTA_' + str( termino.strftime("%d-%m-%Y")) + '.xlsx' 
                        nombre = nombre + entremedio + fin
                        datos = self.conexion.root.informe_reingreso(str(inicio), str(termino))
                        self.informe_reingreso(datos, acept, nombre)
                    else:
                        QMessageBox.about(self,'CONSEJO', 'SELECCIONE ALMENOS UN TIPO DE PROCESO PARA CCREAR EL INFORME DE REINGRESO')

                if datos:
                    self.actualizar()
                    QMessageBox.about(self,'EXITO', 'Informe generado correctamente')
                
            except PermissionError:
                QMessageBox.about(self,'ERROR',  'Otro programa tiene abierto el documento EXCEL. Intente cerrandolo y luego proceda a generar el EXCEL')
            except EOFError:
                self.conexion_perdida()

        elif self.r_venta.isChecked():
            QMessageBox.about(self,'EN DESARROLLO', 'Informes por fecha de venta no disponible en esta version')
            
    def informe_dimensionado(self,datos,nombre):
        if datos:
            wb = Workbook()
            ws = wb.active
            encabezado = ['TIPO DOCUMENTO','NRO DOCUMENTO','NOMBRE CLIENTE', 'NRO ORDEN', 'TELEFONO','DESCRIPCION','CANTIDAD','PRECIO NETO', 'DIMENSIONADOR','VENDEDOR','FECHA VENTA','FECHA ORDEN','FECHA INGRESO','FECHA ESTIMADA','FECHA REAL','ENCHAPADO','DESPACHO','CONTACTO','ORDEN COMPRA','OBSERVACION','ESTADO','MOTIVO','CREADO POR']

            ws.append(encabezado)
            f_ven = None

            for item in datos:
                if item[8]:
                    fecha_venta = datetime.fromisoformat( str( item[8] )) 
                    f_ven =  str( fecha_venta.strftime( "%d-%m-%Y %H:%M:%S" ) )  #FECHA DE VENTA
                f_ing = 'No asignada'
                f_real = 'No asignada'
                if item[4]:
                    ingreso = datetime.fromisoformat(str( item[4]) )
                    f_ing =  str(ingreso.strftime("%d-%m-%Y") ) #FECHA DE INGRESO
                estimada = datetime.fromisoformat(str( item[3]) )
                f_est =  str(estimada.strftime("%d-%m-%Y") )  #FECHA ESTIMADA DE ENTREGA
                if item[5]:
                    real = datetime.fromisoformat(str( item[5]) )
                    f_real =  str(real.strftime("%d-%m-%Y") ) #FECHA REAL DE ENTREGA
                fecha_orden = datetime.fromisoformat(str( item[14]) )
                f_ord =  str(fecha_orden.strftime("%d-%m-%Y") )  #FECHA DE ORDEN    
                detalle = json.loads( item[6])
                cant = detalle['cantidades']
                desc = detalle['descripciones']
                net = detalle["valores_neto"]
                try:
                    creador = detalle["creado_por"]
                except KeyError:
                    creador = 'No registrado'

                if item[19]:
                    extra = json.loads( item[19] )
                    estado = extra["estado"]
                    motivo = extra["motivo"]
                    
                else:
                    estado = 'VALIDA'
                    motivo = 'Ninguno'
                j = 0
                while j < len(cant):
                    fila = [item[11],item[10],item[1], item[0], item[2], desc[j],cant[j],net[j] ,item[9],item[17],f_ven,f_ord,f_ing,f_est,f_real,item[12],item[13],item[15],item[16],item[18],estado,motivo,creador]
                    ws.append(fila)
                    j +=1
            filas_total = ws.max_row
            tab = Table(displayName="tabla_dimensionado" , ref="A1:W"+ str(filas_total) )
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                showLastColumn=False, showRowStripes=True, showColumnStripes=True)
            tab.tableStyleInfo = style
            dim_holder = DimensionHolder(worksheet=ws)
            for col in range(ws.min_column, ws.max_column + 1):
                dim_holder[get_column_letter(col)] = ColumnDimension(ws, min=col, max=col, width=15)
            dim_holder['C'] = ColumnDimension(ws, min=3 ,max = 3, width=30)
            dim_holder['F'] = ColumnDimension(ws, min=6 ,max = 6, width=35)
            dim_holder['J'] = ColumnDimension(ws, min=10 ,max = 10, width=30)
            ws.column_dimensions = dim_holder
            ws.add_table(tab)
            wb.save(nombre)
        
        else:
            QMessageBox.about(self, 'ERROR', 'No se encontraron ordenes de trabajo para el rango de fechas definido')   
 
    def informe_generico(self,datos,nombre):
        if datos:
            wb = Workbook()
            ws = wb.active
            encabezado = ['TIPO DOCUMENTO','NRO DOCUMENTO','NOMBRE CLIENTE', 'NRO ORDEN', 'TELEFONO','DESCRIPCION','CANTIDAD','PRECIO NETO','VENDEDOR','FECHA VENTA','FECHA ORDEN','FECHA ESTIMADA','FECHA INGRESO','TRABAJADOR','FECHA REAL','DESPACHO','CONTACTO','ORDEN COMPRA','OBSERVACION','ESTADO','MOTIVO','CREADA POR']
            ws.append(encabezado)
            fv = None
            for item in datos:
                td =  item[7] #TIPO DOCUMENTO
                nd =  item[6]  #NUMERO DOCUMENTO  
                cl =  item[1]  #CLIENTE
                no =  item[0]  #NUMERO DE ORDEN
                tel =  item[2]  #TELEFONO
                detalle = json.loads( item[12]  )
                cant = detalle['cantidades']  #cantidades
                desc = detalle['descripciones'] #descripciones
                net = detalle["valores_neto"]  #valores neto
                try:
                    creador = detalle["creado_por"]
                except KeyError:
                    creador = 'No registrado'

                if item[13]:
                    fecha_venta = datetime.fromisoformat( str( item[13] ))
                    fv =  str( fecha_venta.strftime( "%d-%m-%Y %H:%M:%S" ) )  #FECHA DE VENTA
                estimada = datetime.fromisoformat(str( item[4]) )
                fe =  str(estimada.strftime("%d-%m-%Y") )  #FECHA ESTIMADA DE ENTREGA
                fr = 'No asignada'
                if item[5]:
                    real = datetime.fromisoformat(str( item[5]) )
                    fr = str(real.strftime("%d-%m-%Y") )   #FECHA REAL
                de =  item[10]  #DESPACHO
                co =  item[8]  #CONTACTO
                oc =  item[9]  #ORDEN DE COMPRA
                ve =  item[14]  #VENDEDOR
                fecha_orden = datetime.fromisoformat(str( item[3]) )
                fo =  str(fecha_orden.strftime("%d-%m-%Y") )  #FECHA DE ORDEN
                ob =  item[15] #OBSERVACION
                if item[16]: #ESTADO
                    extra = json.loads( item[16] )
                    estado = extra["estado"]
                    motivo = extra["motivo"]
                    
                else:
                    estado = 'VALIDA'
                    motivo = 'Ninguno'
                # version 5.5.1 informe actualizado compatible con personal area 5.3.1
                if item[17]:
                    trabajador = item[17]
                else:
                    trabajador = 'NO ASIGNADO'
                if item[18]:
                    ingreso = datetime.fromisoformat(str( item[18]) )
                    fi = str(ingreso.strftime("%d-%m-%Y") )   #FECHA REAL
                else:
                    fi = 'NO INGRESADA'
                #--------------------------------------
                j = 0
                while j < len(cant):
                    fila = [ td,nd,cl,no,tel,desc[j],cant[j],net[j],ve ,fv,fo,fe,fi,trabajador,fr,de,co,oc,ob,estado,motivo,creador]
                    ws.append(fila)
                    j +=1

            filas_total = ws.max_row
            tab = Table(displayName="tabla1" , ref="A1:V"+ str(filas_total) )
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                showLastColumn=False, showRowStripes=True, showColumnStripes=True)
            tab.tableStyleInfo = style
            dim_holder = DimensionHolder(worksheet=ws)
            for col in range(ws.min_column, ws.max_column + 1):
                dim_holder[get_column_letter(col)] = ColumnDimension(ws, min=col, max=col, width=15)
            dim_holder['C'] = ColumnDimension(ws, min=3 ,max = 3, width=30)
            dim_holder['F'] = ColumnDimension(ws, min=6 ,max = 6, width=35)
            dim_holder['I'] = ColumnDimension(ws, min=9 ,max = 9, width=30)
            ws.column_dimensions = dim_holder
            ws.add_table(tab)
            wb.save(nombre)
            
        else:
            QMessageBox.about(self, 'ERROR', 'No se encontraron ordenes de trabajo para el rango de fechas definido')

    def vista_reingreso(self):
        if self.comboBox.currentText() == 'REINGRESO':
            print('modo reingreso')
            self.groupBox.show()
            self.r_orden_2.hide()
            self.r_venta.hide()
            self.lb_buscar.show()
            self.lb_buscar.setText('FECHA DE CREACION')
        else:
            self.groupBox.hide()
            self.lb_buscar.hide()
            self.r_orden_2.show()
            self.r_venta.show()

    def informe_reingreso(self, datos,acept, nombre):
        if datos:
            wb = Workbook()
            ws = wb.active
            ws.title = 'REINGRESO_1'
            encabezado = ['NUMERO REINGRESO','FECHA REINGRESO','NUMERO DE ORDEN', 'TIPO DOCUMENTO', 'NUMERO DOCUMENTO', 'PROCESO','MERCADERIA','CANTIDAD','VALOR NETO','MOTIVO','DESCRIPCIÓN' ,'SOLUCIÓN','CREADO POR']
            ws.append(encabezado)
            for item in datos:
                nro_re =  str( item[0] ) #NRO REINGRESO INT
                reingreso = datetime.fromisoformat(str( item[1]) )
                fr =  str(reingreso.strftime("%d-%m-%Y") )  #FECHA REINGRESO
                td =  item[2] #TIPO DOCUMENTO   STR
                nd =  str( item[3] )  #NUMERO DOCUMENTO  INT
                no =  str( item[4] )  #NUMERO DE ORDEN    INT
                mot =  item[5]  #MOTIVO   STR
                descripcion =  item[6]  #DESCRIPCION  STR
                proc =  item[7]  #PROCESO   STR
                if proc in acept: #se filtra
                    detalle = json.loads( item[8]  ) #DETALLE
                    cant = detalle['cantidades']  #cantidades
                    merc = detalle['descripciones'] #mercaderia
                    net = detalle["valores_neto"]  #valores neto
                    try:
                        creador = detalle["creado_por"]
                    except KeyError:
                        creador = 'No registrado'

                    sol =  item[9] #SOLUCION  STR
                    j = 0
                    while j < len(cant):
                        fila = [ nro_re, fr, no, td, nd, proc, merc[j], cant[j], net[j], mot, descripcion ,sol, creador]
                        ws.append(fila)
                        j +=1

            filas_total = ws.max_row
            tab = Table(displayName="tabla1" , ref="A1:M"+ str(filas_total) )
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                showLastColumn=False, showRowStripes=True, showColumnStripes=True)
            tab.tableStyleInfo = style
            dim_holder = DimensionHolder(worksheet=ws)
            for col in range(ws.min_column, ws.max_column + 1):
                dim_holder[get_column_letter(col)] = ColumnDimension(ws, min=col, max=col, width=10)
            dim_holder['F'] = ColumnDimension(ws, min=6 ,max = 6, width=16)  #PROCESO
            dim_holder['G'] = ColumnDimension(ws, min=7 ,max = 7, width=45) #MERCADERIA 45px
            dim_holder['I'] = ColumnDimension(ws, min=9 ,max = 9, width=14) #VALOR NETO 10px
            dim_holder['J'] = ColumnDimension(ws, min=10 ,max = 10, width=14) #MOTIVO
            dim_holder['K'] = ColumnDimension(ws, min=11 ,max = 11, width=30)  #DESCRIPCION
            dim_holder['L'] = ColumnDimension(ws, min=12 ,max = 12, width=30)  #SOLUCION
            ws.column_dimensions = dim_holder
            ws.add_table(tab)
            wb.save(nombre)
        else: 
            QMessageBox.about(self,'ERROR', 'No se encontraron datos de reingreso para las fechas ingresadas.')

    def actualizar(self):
        self.tableWidget.setRowCount(0)
        informes = os.listdir(self.dir_informes)
        for item in informes:
            fila = self.tableWidget.rowCount()
            self.tableWidget.insertRow(fila)
            self.tableWidget.setItem(fila , 0 , QTableWidgetItem( item ) )  #nombre informe
    def abrir(self):
        seleccion = self.tableWidget.selectedItems()
        if seleccion != [] :
            nombre = seleccion[0].text()
            abrir = self.dir_informes + nombre
            if os.path.isfile(abrir):
                subprocess.Popen([abrir], shell=True)
            else:
                QMessageBox.about(self,'ERROR', 'El archivo no se encontro en la carpeta "informes" ')
        else:
            QMessageBox.about(self,'Sugerencia', 'Primero seleccione el nombre del informe para poder abrirlo ')

    def eliminar_excel(self):
        seleccion = self.tableWidget.selectedItems()
        if seleccion != [] :
            nombre = seleccion[0].text()
            ruta = self.dir_informes + nombre
            if os.path.isfile(ruta):
                try:
                    os.remove(ruta)
                    self.actualizar()
                except PermissionError:
                    QMessageBox.about(self,'ERROR', 'No se pudo eliminar el informe, ya que archivo esta siendo utilizado por otro programa')
            else:
                QMessageBox.about(self,'ERROR', 'El archivo no se encontro')
        else:
            QMessageBox.about(self,'Sugerencia', 'Primero seleccione el nombre del informe para poder abrirlo ')


 # -------- funcion para generar clave ------------
    def generar_clave(self):
        dialog = InputDialog2('CLAVE:', False, 'REGISTRAR CLAVE',self)
        if dialog.exec():
            clave = dialog.getInputs()
            if clave != '':
                try:
                    if self.conexion.root.registrar_clave(clave):
                        QMessageBox.about(self,'EXITO' ,'CLAVE: ' + clave +' REGISTRADA')
                    else:
                        QMessageBox.about(self,'ERROR' ,'ERROR AL REGISTRAR LA CLAVE')
                except EOFError:
                    self.conexion_perdida()
            else:
                QMessageBox.about(self,'ERROR' ,'Ingrese una clave antes de continuar')

#----- funciones generales -----------
    def crear_pdf(self, lista, tipo, despacho): #PDF DE ORDEN DE TRABAJO
        ruta = ( self.carpeta +'/ordenes/' + tipo +'_' + lista[0] + '.pdf' )  #NRO DE ORDEN  
        formato = self.carpeta +"/formatos/" + tipo +".jpg"
        agua = self.carpeta + "/formatos/despacho.png"
        uso_interno = self.carpeta + "/formatos/uso interno.png"
        hojas = 2
        if tipo == 'carpinteria':
            hojas = 1
        try:
            documento = canvas.Canvas(ruta)

            for pagina in range(hojas):
                documento.setPageSize(( 216 * mm , 279 * mm))
                documento.drawImage( formato, 0* mm , 2 * mm , 216 *mm ,279 *mm )


                if despacho == 'SI':
                    documento.setFillAlpha(0.6)
                    documento.drawImage( agua , 83* mm , 30* mm , 100*mm ,100*mm , mask= 'auto')
                    documento.drawImage( agua , 83* mm , (30+136)* mm , 100*mm ,100*mm , mask= 'auto')

                if self.r_uso_interno_1.isChecked():
                    documento.setFillAlpha(0.6)
                    documento.drawImage( uso_interno , 100* mm , 30* mm , 69*mm ,94.5 *mm , mask= 'auto')
                    documento.drawImage( uso_interno , 100* mm , (30+136)* mm , 69*mm ,94.5*mm , mask= 'auto')

                documento.setFillAlpha(1)
                documento.drawString( 0 * mm, 139.5 * mm ,'------------------------------------------------------------------------------------------')
                documento.drawString( 105 * mm, 139.5 * mm ,'----------------------------------------------------------------------------------------------')

                documento.rotate(90)

                documento.setFont('Helvetica',10)

                k = 2.5 #constante
                salto = 0
                for i in range(2):

                    if self.r_facturar_1.isChecked():
                        documento.setFont('Helvetica',11)
                        documento.setFillAlpha(0.6)
                        documento.drawString( (53 + k + salto) *mm , -20.5 * mm , '"POR FACTURAR"' )  #por facturar
                        documento.setFillAlpha(1)

                    documento.setFont('Helvetica',9)
                    documento.drawString( (28 + k + salto) *mm , -59.5 * mm , lista[2] )  #NOMBRE

                    documento.drawString( (100 + k + salto) *mm , -66 * mm , str(lista[3]) )   #TELEFONO
                    documento.drawString( (106 + k +salto) *mm , -59.5 * mm , lista[1]  )    #FECHA DE ORDEN

                    
                    
                    if self.tipo_doc == 'FACTURA':
                        documento.drawString( (45 + k + salto) *mm , -85 * mm , str(self.nro_doc) )     #NRO FACTURA FOLIO
                    elif self.tipo_doc == 'BOLETA':
                        documento.drawString( (15 + k + salto) *mm , -85 * mm ,  str(self.nro_doc) )      #NRO BOLETA FOLIO
                    elif self.tipo_doc == 'GUIA':
                        documento.drawString( (78 + k + salto) *mm , -85 * mm , str(self.nro_doc) )     #NRO GUIA  V5.4 VENDEDOR

                    if tipo == 'dimensionado':
                        documento.drawString( (88 + k + salto) *mm , -94 * mm ,  lista[7] )   #ENCHAPE

                    documento.drawString( (110 + k + salto) *mm , -85 * mm ,   lista[9]  )       #ORDEN DE COMPRA

                    documento.drawString( (32 + k + salto) *mm , -66 * mm ,  lista[8] )   #CONTACTO

                    documento.drawString( (33+ salto) *mm , -205 * mm , lista[10] ) #NOMBRE VENDEDOR
                    
                    documento.setFont('Helvetica-Bold',12)
                    documento.drawString( (106 + k + salto) *mm , -44.5 * mm , lista[4]  ) #FECHA ESTIMADA
                    documento.drawString( (106 + k + salto) *mm , -20.5 * mm , lista[0]  ) #NRO DE ORDEN

                    documento.setFillColor(white)
                    documento.setStrokeColor(white)
                    documento.rect((5+ k + salto )* mm, -200 * mm ,80*mm,3*mm, fill=1 )
                    documento.setFillColor(black)
                    documento.setStrokeColor(black)
                    documento.setFont('Helvetica-Bold',10)

                    if pagina == 0:
                        if k == 2.5:
                            documento.drawString( (10+ salto) *mm , -200 * mm , 'ORIGINAL' )
                        else:
                            documento.drawString( (10+ salto) *mm , -200 * mm , 'COPIA ' + tipo.upper() )
                    else:
                        if k == 2.5:
                            documento.drawString( (10+ salto) *mm , -200 * mm , 'COPIA CLIENTE' )
                        else:
                            documento.drawString( (10+ salto) *mm , -200 * mm , 'COPIA BODEGA' )
                    documento.drawString( (10+ salto) *mm , -205 * mm , 'VENDEDOR:' )        
                    salto += 139.5
                    k = 0 

                #items
                cons = 108
                documento.setFont('Helvetica',9)
                align_cant_1 = 12
                align_descr_1 = 22
                align_cant_2 = 151
                align_descr_2 = 161

                cantidades = self.normalizar_cantidades(lista[5]) #arreglado los int se ven como int y float como float en la impresion
                descripciones = lista[6]
                i = 0
                while i < len(cantidades):
                    documento.drawString(align_cant_1 *mm , -cons* mm , str(cantidades[i]) )  #cantidad
                    
                    documento.drawString(align_cant_2 *mm , -cons * mm , str(cantidades[i]) )

                    cadenas = self.separar(descripciones[i])

                    for cadena in cadenas:
                        documento.drawString(align_descr_1 *mm , -cons* mm , cadena )  #descripcion

                        documento.drawString(align_descr_2 *mm , -cons* mm , cadena ) #54 LONG MAXIMA
                        cons += 5
                    i +=1

                documento.showPage()
            #documento.drawString( (10+ salto) *mm , -205 * mm , 'VENDEDOR:' )
            documento.save()
            sleep(1)
        except PermissionError:
            QMessageBox.about(self,'ERROR', 'Otro programa esta modificando este archivo por lo cual no se puede modificar actualmente.')
    
    def crear_pdf_reingreso(self, datos): #pdf de reingreso de mercaderias
            print('creando pd reingreso...')
            documento = canvas.Canvas(self.carpeta +'/reingresos/reingreso_' + str(datos[0]) + '.pdf')
            imagen =  self.carpeta + "/formatos/reingreso_solo.jpg" 

            documento.setPageSize(( 216 * mm , 279 * mm))
            documento.drawImage( imagen, 0* mm , 0 * mm , 216 *mm , 139.5 *mm )
            documento.drawImage( imagen, 0* mm , 139.5 * mm , 216 *mm , 139.5 *mm )
            documento.drawString( 0*mm , 139.5 *mm , '------------------------------------------------------------------------------')
            documento.drawString( 108*mm , 139.5 *mm , '----------------------------------------------------------------------------')
            salto = 0
        
            for i in range(2):
                documento.setFont('Helvetica',9)
                
                if datos[2] == 'FACTURA':
                    documento.drawString(129*mm, (106.5 + salto )*mm , str(datos[3])  )   #NUMERO DOCUMENTO , FACTURA
                elif datos[2] == 'BOLETA':
                    documento.drawString(52* mm, (106.5+ salto )*mm , str(datos[3])  )   #NUMERO DOCUMENTO , BOLETA
                elif datos[2] == 'GUIA':
                    documento.drawString(95* mm, (106.5+ salto )*mm , str(datos[3]) )   #NUMERO DOCUMENTO , guia
            
                lista = self.separar2(datos[5],94) #DESCRIPCION
                
                print(len( datos[5] ))
                k = 0 
                j = 0 
                for item in lista:
                    documento.drawString(20* mm, (92.5 + salto - k)*mm , lista[j] )  #descripcion del problema
                    #documento.drawString(20* mm, (92.5 + salto - k )*mm , descr )  #descripcion del problema
                    k += 6
                    j += 1
                
                lista = self.separar2( datos[7] , 85) #SOLUCION
                print(len(datos[7]))
                k = 0 
                j = 0
                for item in lista:
                    documento.drawString(40*mm, (39 + salto - k )*mm , lista[j] )  #solucion del problema
                    k += 6
                    j += 1
                
                cants = datos[8]
                descrs = datos[9]
                netos = datos[10]
                p = 0
                q = 0
                for item in cants:
                    documento.drawString(150*mm, (65 + salto - q )*mm , str(cants[p]) )  #cantidad
                    documento.drawString(170*mm, (65 + salto - q )*mm , str(netos[p]) )  #neto
                    cadenas = self.separar2(descrs[p] , 60 ) 
                    for cadena in cadenas:
                        documento.drawString(20*mm, (65 + salto -q )*mm , cadena)  #descripcion
                        q += 5
                    p +=1

                documento.setFillAlpha(0.6)
                documento.drawString(18*mm, (12 + salto - k )*mm , 'Vendedor:' )  #nombre vendedor
                documento.drawString(38*mm, (12 + salto - k )*mm , self.datos_usuario[8] )  #nombre vendedor
                documento.setFillAlpha(1)

                documento.setFont('Helvetica-Bold', 9 )

                documento.drawString(177*mm, (124.5 + salto )*mm ,  str(datos[0]) )  #NRO DE REINGRESO
                documento.drawString(177*mm, (115.5 + salto ) *mm , datos[1] )    #FECHA DEL REINGRESO

                if datos[6] == 'DIMENSIONADO':
                    documento.drawString(73*mm, (77 + salto )*mm , 'X' )  #PROCESO DIMENSIONADO
                elif datos[6] == 'ELABORACION':    
                    documento.drawString(150*mm, (77 + salto )*mm , 'X' )  #PROCESO ELABORACION
                elif datos[6] == 'CARPINTERIA':
                    documento.drawString(111*mm, (77 + salto )*mm , 'X' )  #PROCESO CARPINTERIA
                elif datos[6] == 'PALLETS':
                    documento.drawString(178.5*mm, (77 + salto )*mm , 'X' )  #PROCESO PALLETS


                if datos[4] == 'CAMBIO':
                    documento.drawString(53*mm, (99.5 + salto )*mm , 'X' )  #motivo cambio
                elif datos[4] == 'DEVOLUCION':
                    documento.drawString(92*mm, (99.5 + salto )*mm , 'X' )   #motivo devolucion
                else:
                    documento.setFont('Helvetica',9)
                    documento.drawString(128*mm, (99.5 + salto )*mm , datos[4] )  #motivo otro

                salto += 139.5 
            documento.save()
    def anular(self):
        print(self.datos_usuario)

        if self.datos_usuario[4] == 'SI':
            dialog = InputDialog2('MOTIVO:' , False ,'INGRESE UN MOTIVO DE ANULACIÓN' ,self)
            dialog.resize(330,80)
            if dialog.exec():
                motivo  = dialog.getInputs()
                print(motivo)
                formato = {
                            "estado" : 'ANULADA',
                            "motivo" : motivo,
                            "usuario": self.datos_usuario[8]
                        }
                extra = json.dumps(formato)
                
                tipo = self.lb_tipo_orden.text()
                try:
                    self.conexion.root.anular_orden( tipo.lower(), extra, self.nro_orden )
                    QMessageBox.about(self,'EXITO' ,'Orden nro: ' + str(self.nro_orden) + ' Anulada')
                    self.inicializar_buscar_orden()
                except EOFError:
                    self.conexion_perdida()
        else:
            dialog = InputDialog('CLAVE:' ,'MOTIVO:', 'INGRESE UNA CLAVE Y MOTIVO DE ANULACION',self)
            dialog.resize(400,100)
            if dialog.exec():
                aux_clave , motivo = dialog.getInputs()
                claves = self.conexion.root.obtener_clave()
                clave = (aux_clave ,)
                if clave in claves:
                    print('clave valida')
                    formato = {
                            "estado" : 'ANULADA',
                            "motivo" : motivo,
                            "usuario": self.datos_usuario[0]
                        }
                    extra = json.dumps(formato)
                    
                    tipo = self.lb_tipo_orden.text()
                    
                    try:
                        self.conexion.root.anular_orden( tipo.lower(), extra, self.nro_orden )
                        self.conexion.root.eliminar_clave(aux_clave)
                        QMessageBox.about(self,'EXITO' ,'Orden nro: ' + str(self.nro_orden) + ' Anulada')
                        self.inicializar_buscar_orden()
                    except EOFError:
                        self.conexion_perdida()
                else:
                    #print('clave invalida')
                    QMessageBox.about(self,'ERROR' ,'La clave ingresada no es valida')
    def buscar_nro_orden(self,tupla):
        mayor = 0
        for item in tupla:
            if item[0] > mayor :
                mayor = item[0]
        return mayor
    def ver_pdf_reingreso(self):
        ruta = self.carpeta + '/reingresos/reingreso_' + str(self.nro_reingreso) + '.pdf'
        subprocess.Popen([ruta], shell=True)

    def ver_pdf(self, tipo):
        abrir = self.carpeta+ '/ordenes/' + tipo.lower() +'_' +str(self.nro_orden) + '.pdf'
        subprocess.Popen([abrir], shell=True)

    def separar(self,cadena):
        lista = []
        iter = len(cadena)/54
        iter = int(iter) + 1 #cantidad de items a escribir
        #print('----------------------------------------------------')
       # print(cadena)
        #print('espacios necesarios: ' + str(iter))
        i = 0
        while len(cadena)> 54:
        
            #print('long > 54:')
            aux = cadena[0:54]
            index = aux[::-1].find(' ')
        
            aux = aux[:(54-index)]
           # print('Iteracion: '+str(i)+ ': '+ aux)
            lista.append(aux)
            cadena = cadena[54 - index :]
            i += 1

        if len(cadena) > 0 :
            vacias = cadena.count(' ')
            if vacias == len(cadena):
                print('item O CADENA vacio')
                #print('----------------------------------------------------')
            else:
               # print('LONG < 54: ' + cadena)
                lista.append(cadena)
                #print('----------------------------------------------------')

        return lista

    def separar2(self,cadena, long): #(TEXTO, cantidad de caracteres por linea) : 54 para pdf de ordenes y 85 para reingresos
        lista = []
        iter = len(cadena)/long
        iter = int(iter) + 1 #cantidad de items a escribir
        print('----------------------------------------------------')
        print(cadena)
        print('espacios necesarios: ' + str(iter))
        i = 0
        while len(cadena)> long:
        
            print('long > '+ str(long) +':')
            aux = cadena[0:long]
            index = aux[::-1].find(' ')
        
            aux = aux[:(long-index)]
            print('Iteracion: '+str(i)+ ': '+ aux)
            lista.append(aux)
            cadena = cadena[long - index :]
            i += 1

        if len(cadena) > 0 :
            vacias = cadena.count(' ')
            if vacias == len(cadena):
                print('item vacio')
                print('----------------------------------------------------')
            else:
                print('fin long < '+ str(long) +':' + cadena)
                lista.append(cadena)
                print('----------------------------------------------------')

        return lista

    def normalizar_cantidades(self, lista):
        aux = []
        for i in lista:
            number = str(i)
            n = number.split('.')
            if n[1] == '0':
                aux.append(int(i))
            else:
                aux.append(i)
        return aux
# ------ Funciones del menu -----------
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
                self.btn_estadisticas.setText('Estadisticas')
            else:
                self.logo.hide()
                self.btn_buscar.setText('')
                self.btn_modificar.setText('')
                self.btn_orden_manual.setText('')
                self.btn_generar_clave.setText('')
                self.btn_informe.setText('')
                self.btn_atras.setText('')
                self.btn_estadisticas.setText('')
                
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

    def changeEvent(self, event):
        if event.type() == QEvent.WindowStateChange:
            if event.oldState() and Qt.WindowMinimized:
                print("WindowMinimized")
            elif event.oldState() == Qt.WindowNoState or self.windowState() == Qt.WindowMaximized:
                print("WindowMaximized, cambiando tamaño de las tablas ...")
                # ----- SE ADAPTAN TODAS LAS TABLAS  -------
                #Buscar venta
                self.tableWidget_1.setColumnWidth(0,80) #interno
                self.tableWidget_1.setColumnWidth(1,80) #documento
                self.tableWidget_1.setColumnWidth(2,80) #nro doc
                self.tableWidget_1.setColumnWidth(3,125) #fecha venta
                self.tableWidget_1.setColumnWidth(4,200) #vendededor
                self.tableWidget_1.setColumnWidth(5,100) #total
                #Crear venta
                #Buscar orden
                #Modificar orden
                #reingreso
                #ORDEN MANUAL
                #REINGRESO MANUAL


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

class InputDialog2(QDialog):
    def __init__(self, label1, hide , titulo , parent=None ):
        super().__init__(parent)
        self.hide = hide
        self.setWindowTitle(titulo)
        self.txt1 = QLineEdit(self)
        self.inicializar()
        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self);

        layout = QFormLayout(self)
        layout.addRow(label1, self.txt1)
        layout.addWidget(buttonBox)

        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)
    
    def inicializar(self):
        if self.hide == True:
            self.txt1.setEchoMode(QLineEdit.Password)

    def getInputs(self):
        return self.txt1.text()


    
if __name__ == '__main__':
    x = 0
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('icono_imagen/icono_barv3.png'))
    myappid = 'madenco.personal.area' # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid) 
    vendedor = Vendedor()
    vendedor.show()
    sys.exit(app.exec_())
