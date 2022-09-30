import sys
import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from  matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from PyQt5 import uic
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QPixmap , QIcon
from PyQt5.QtCore import Qt,QEasingCurve, QPropertyAnimation,QEvent,QStringListModel
import rpyc
import socket
from time import sleep, strftime
from datetime import datetime, timedelta,date 
import json
import ctypes
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch , mm
from reportlab.lib.colors import white, black
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
from subprocess import Popen
from concurrent.futures import ThreadPoolExecutor
from random import randint


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
# ------ VARIABLES GLOBALES -----------
        self.datos_usuario = None
        self.conexion = None
        self.host = None
        self.puerto = None

        self.carpeta = None
        self.dir_informes = None
        
        self.vendedores = None
        self.aux_tabla = None #tabla de respaldo, para usar filtros. usada par fitrar x vendedor, nulas y manuales incompletas
        self.bol_fact = None # boletas y facturas encontradas
        self.guias = None    # guias encontradas
        self.manuales = None # respaldo de solo ordenes manuales
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
        self.aux_vendedor = None
        self.anterior = None #Determina si al volver atras, vuelve a buscar ordenes o estadisticas -> buscar ordenes manuales
        self.anterior2 = None #Determina si volver atras vuelve a gestion de reingresos o busqueda de ordenes de trabajo.
        self.aux_detalle = None
        self.version = 'version-5.8'
        self.executor = ThreadPoolExecutor(max_workers=3) #Hilos para hacer Sondeos a la base de datos
        self.sondeo_params = {
            "estado_sondeo_manual" : False, 
            "intervalo_sondeo_manual" : 300, # 300 segundos x DEFECTO
            "omitir_uso_carpinteria" : False,
            "omitir_vinculo_exitoso" : False
        }
#--------------- FUNCIONES GLOBALES ---------
        self.iniciar_session()
        self.inicializar()
        self.stackedWidget.setCurrentWidget(self.inicio)
        
        #buscar venta
        self.btn_buscar_1.clicked.connect(self.buscar_documento)
        self.comboBox_1.currentIndexChanged['QString'].connect(self.filtrar_vendedor)
        self.btn_crear_1.clicked.connect(self.inicializar_crear_orden)
        self.btn_atras_1.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.inicio))
        self.btn_despacho_1.clicked.connect(self.asignar_despacho)

        #crear orden
        self.btn_registrar_2.clicked.connect(self.registrar_orden)
        self.btn_atras_2.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.buscar_venta))
        self.btn_agregar.clicked.connect(self.agregar)
        self.btn_eliminar.clicked.connect(self.eliminar)
        self.btn_rellenar.clicked.connect(self.rellenar)
        self.completer = QCompleter()
        self.completer.activated.connect(self.rellenar_datos_cliente)
        self.nombres = []
        self.telefonos = []
        self.contactos = []

        self.r_dim.toggled.connect(lambda: self.cargar_clientes('dimensionado','normal') if(self.r_dim.isChecked()) else print('no ckeck dim') )
        self.r_elab.toggled.connect(lambda: self.cargar_clientes('elaboracion','normal') if(self.r_elab.isChecked()) else print('no ckeck elab'))
        self.r_carp.toggled.connect(lambda: self.cargar_clientes('carpinteria','normal') if(self.r_carp.isChecked()) else print('no ckeck carp'))
        self.r_pall.toggled.connect(lambda: self.cargar_clientes('pallets','normal') if(self.r_pall.isChecked()) else print('no ckeck pall'))

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
        self.btn_atras_5.clicked.connect(self.decidir_atras)
        self.btn_ver_5.clicked.connect(lambda: self.ver_pdf(self.tipo) )
        self.btn_agregar_.clicked.connect(self.agregar_2)
        self.btn_eliminar_.clicked.connect(self.eliminar_2)
        self.btn_anular_5.clicked.connect(self.anular)
        # reingreso
        self.btn_reingreso_4.clicked.connect(self.inicializar_reingreso)
        self.btn_generar_reingreso.clicked.connect(self.registrar_reingreso)
        self.btn_agregar_2.clicked.connect(self.agregar_3)
        self.btn_eliminar_2.clicked.connect(self.eliminar_3)
        self.btn_atras_7.clicked.connect( lambda: self.stackedWidget.setCurrentWidget(self.buscar_orden) )
        # GESTION REINGRESO
        self.btn_buscar_reingreso.clicked.connect(self.buscar_reingreso_1)
        self.btn_mod_reingreso_1.clicked.connect(self.vista_modificar_reingreso)
        self.btn_anular_reingreso.clicked.connect(lambda:  self.anular_validar_reingreso('ANULAR') )
        self.btn_validar_reingreso.clicked.connect( lambda:  self.anular_validar_reingreso('VALIDAR')  )

        self.btn_guardar_6.clicked.connect(lambda: self.reingreso_manual('actualizar')  )
        self.btn_ver_pdf_reingreso.clicked.connect(self.verificar_pdf_reingreso)
        # INGRESO MANUAL
        #  ---- ORDEN MANUAL ----
        self.txt_descripcion_1.textChanged.connect(self.buscar_descripcion)
        self.txt_codigo_1.textChanged.connect(self.buscar_codigo)
        self.btn_add.clicked.connect(self.add_descripcion)
        self.r_uso_interno_1.stateChanged.connect(self.cambiar_observacion)
        self.btn_registrar_1.clicked.connect(self.registrar_orden_manual)
        self.btn_agregar_1.clicked.connect(self.agregar_4)
        self.btn_eliminar_1.clicked.connect(self.eliminar_4)
        self.completer_manual = QCompleter()
        self.completer_manual.activated.connect(self.rellenar_datos_cliente_manual)

        self.r_dim_1.toggled.connect(lambda: self.cargar_clientes('dimensionado','manual') if(self.r_dim_1.isChecked()) else print('no ckeck dim') )
        self.r_elab_1.toggled.connect(lambda: self.cargar_clientes('elaboracion','manual') if(self.r_elab_1.isChecked()) else print('no ckeck elab'))
        self.r_carp_1.toggled.connect(lambda: self.cargar_clientes('carpinteria','manual') if(self.r_carp_1.isChecked()) else print('no ckeck carp'))
        self.r_pall_1.toggled.connect(lambda: self.cargar_clientes('pallets','manual') if(self.r_pall_1.isChecked()) else print('no ckeck pall'))
        #  ---- REINGRESO MANUAL ----
        self.btn_registrar_6.clicked.connect(lambda: self.reingreso_manual('registrar') )
        self.txt_descripcion_7.textChanged.connect(self.buscar_descripcion_2)
        self.btn_add_6.clicked.connect(self.add_descripcion_2)

        self.btn_agregar_6.clicked.connect(self.agregar_6)
        self.btn_eliminar_6.clicked.connect(self.eliminar_6)

        self.btn_atras_3.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.inicio))
        self.btn_atras_6.clicked.connect(self.decidir_atras2)
        #informes
        self.btn_generar_informe.clicked.connect(self.generar_informe)
        self.comboBox.currentIndexChanged['QString'].connect(self.vista_reingreso)
        self.btn_eliminar_exel.clicked.connect(self.eliminar_excel)
        self.btn_actualizar.clicked.connect(self.actualizar)
        self.btn_abrir.clicked.connect(self.abrir)
        #ESTADISTICAS
        self.btn_grafico.clicked.connect(self.crear_grafico)
        self.btn_buscar_manuales.clicked.connect(self.buscar_manuales)
        self.btn_modificar_5.clicked.connect(self.continuar_orden_manual)
        self.check_incompletas.stateChanged.connect(self.filtrar_solo_incompletas)
        self.btn_notificacion1.clicked.connect(self.easter_egg )
        #SONDEOS

        self.btn_iniciar_sondeo_manual.clicked.connect(self.iniciar_sondeo)
        self.btn_detener_sondeo_manual.clicked.connect(self.finalizar_sondeo)
        self.btn_actualizar_config_sondeo.clicked.connect(self.actualizar_sondeo_config)
        self.btn_obtener_sondeo.clicked.connect(lambda: self.obtener_datos_sondeo(True))
        #generar clave
        self.btn_generar_clave.clicked.connect(self.generar_clave)

        #CONFIGURACION
        self.btn_configuracion.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.configuracion))
        self.btn_inicio.clicked.connect(lambda: self.stackedWidget.setCurrentWidget(self.inicio))
        #SIDE MENU BOTONES
        self.btn_buscar.clicked.connect(self.inicializar_buscar_venta)
        self.btn_modificar.clicked.connect(self.inicializar_buscar_orden)
        self.btn_orden_manual.clicked.connect(self.inicializar_ingreso_manual)
        self.btn_informe.clicked.connect(self.inicializar_informe)
        self.btn_atras.clicked.connect(self.cerrar_sesion)
        self.btn_conectar.clicked.connect(self.conectar)
        self.btn_estadisticas.clicked.connect(self.inicializar_estadisticas)
        self.btn_reingreso.clicked.connect(self.inicializar_buscar_reingreso)
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

        self.btn_buscar.setIcon(QIcon('icono_imagen/buscar venta.png'))
        self.btn_modificar.setIcon(QIcon('icono_imagen/buscar orden.png'))
        self.btn_orden_manual.setIcon(QIcon('icono_imagen/orden manual.png'))
        self.btn_generar_clave.setIcon(QIcon('icono_imagen/generar clave.png'))
        self.btn_informe.setIcon(QIcon('icono_imagen/informes.png'))
        self.btn_estadisticas.setIcon(QIcon('icono_imagen/estadisticas2.png'))
        self.btn_configuracion.setIcon(QIcon('icono_imagen/configuracion.png'))
        self.btn_reingreso.setIcon(QIcon('icono_imagen/reingreso.png'))

        self.btn_generar_clave.show()
        self.btn_orden_manual.show()
        self.btn_informe.show()

        actual = os.path.abspath(os.getcwd())
        actual = actual.replace('\\' , '/')
        self.carpeta = actual.replace('\\' , '/')
        self.dir_informes = actual + '/informes/'
        print(self.carpeta)
        
        self.btn_inicio.setIcon(QIcon('icono_imagen/enco_logov3.png'))

        menu = QPixmap(actual + '/icono_imagen/menu_v4.png')
        self.btn_menu.setIcon(QIcon(menu))
        self.lb_conexion.setText('CONECTADO')
        if self.datos_usuario:  #Si existen los datos del usuario, x ende se inicio sesion correctamente...
            self.lb_vendedor.setText(self.datos_usuario[8]) #NOMBRE DEL VENDEDOR
            self.vendedor = self.datos_usuario[8]   #solucion al error cuando se modifica la orden y se crear el pdf interno con vendedor = none. Esto cierra el programa.

            if self.datos_usuario[4] == 'NO': #Si no es super usuario
                self.btn_generar_clave.hide() #no puede generar claves
                detalle = json.loads(self.datos_usuario[7])
                funciones = detalle['vendedor']
                if not 'manual' in funciones:
                    self.btn_orden_manual.hide() #no puede generar ordenes manuales
                if not 'informes' in funciones:
                    self.btn_informe.hide() #no puede generar informes
        if os.path.isfile(actual + '/manifest.txt'):
            with open(actual + '/manifest.txt' , 'r', encoding='utf-8') as file:
                lines = file.readlines()
                try:
                    t_sondeo = lines[2].split(':')
                    t_sondeo = t_sondeo[1]
                    t_sondeo = t_sondeo[:len(t_sondeo)-1]
                    print(t_sondeo)
                    self.sondeo_params['intervalo_sondeo_manual'] = int(t_sondeo)

                    o_uso_carp = lines[3].split(':')
                    o_uso_carp = o_uso_carp[1]
                    o_uso_carp = o_uso_carp[:len(o_uso_carp)-1]
                    if o_uso_carp == 'SI':
                        self.sondeo_params["omitir_uso_carpinteria"] = True
                        self.label_75.setText('SI')

                    o_vinc_ex = lines[4].split(':')
                    o_vinc_ex = o_vinc_ex[1]
                    o_vinc_ex = o_vinc_ex[:len(o_vinc_ex)-1]

                    if o_vinc_ex == 'SI':
                        self.sondeo_params["omitir_vinculo_exitoso"] = True
                        self.label_76.setText('SI')

                    
                except IndexError:
                    print('error de indice del manifest - sondeo')
                    pass #si no encuentra alguna linea
        else:
            print('manifest no encontrado')

        self.iniciar_sondeo()

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
        self.txt_interno_1.setText('')
        self.txt_cliente_1.setText('')
        self.tableWidget_1.setRowCount(0)
        self.radio2_1.setChecked(True)
        self.dateEdit_1.setCalendarPopup(True)
        self.dateEdit_1.setDate(datetime.now().date())
        if self.windowState() != Qt.WindowMaximized:
            print('Ajustando tamaño de las tablas para ventana NORMAL')
            self.tableWidget_1.setColumnWidth(0,70) #interno
            self.tableWidget_1.setColumnWidth(1,70) #documento
            self.tableWidget_1.setColumnWidth(2,65) #nro doc
            self.tableWidget_1.setColumnWidth(3,100) #fecha venta
            self.tableWidget_1.setColumnWidth(4,100) #cliente
            self.tableWidget_1.setColumnWidth(5,100) #vendededor
            self.tableWidget_1.setColumnWidth(6,50) #total
            self.tableWidget_1.setColumnWidth(7,60) #despacho
        else:
            print('VENTANA MAXIMIZADA, no es necesario ajustar el tamaño de las tablas')
        
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

                        self.tableWidget_1.setItem(fila , 7 , QTableWidgetItem(   str(consulta[7]))   )      #DESPACHO
                        self.tableWidget_1.setItem(fila , 8 , QTableWidgetItem(   str(consulta[8]))   )      #rut

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
                        self.tableWidget_1.setItem(fila , 7 , QTableWidgetItem(   str(consulta[7]))   )      #DESPACHO
                        self.tableWidget_1.setItem(fila , 8 , QTableWidgetItem(    detalle['rut']  ) )      #rut

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
                            self.tableWidget_1.setItem(fila , 7 , QTableWidgetItem(   str(consulta[7]))   )      #DESPACHO
                            self.tableWidget_1.setItem(fila , 8 , QTableWidgetItem(   str(consulta[8]))   )      #rut
                    
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
                            self.tableWidget_1.setItem(fila , 7 , QTableWidgetItem(   str(consulta[7]))   )      #DESPACHO
                            self.tableWidget_1.setItem(fila , 8 , QTableWidgetItem(    detalle['rut']  ) )      #rut
                

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
                            self.tableWidget_1.setItem(fila , 7 , QTableWidgetItem(   str(consulta[7]))   )      #DESPACHO
                            self.tableWidget_1.setItem(fila , 8 , QTableWidgetItem(   str(consulta[8]))   )      #rut
                    
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
                            self.tableWidget_1.setItem(fila , 7 , QTableWidgetItem(   str(consulta[7]))   )      #DESPACHO
                            self.tableWidget_1.setItem(fila , 8 , QTableWidgetItem(    detalle['rut']  ) )      #rut
                

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
                self.tableWidget_1.setItem(fila , 7 , QTableWidgetItem(   str(consulta[7]))   )      #DESPACHO
                self.tableWidget_1.setItem(fila , 8 , QTableWidgetItem(   str(consulta[8]))   )      #rut

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
                self.tableWidget_1.setItem(fila , 7 , QTableWidgetItem(   str(consulta[7]))   )      #DESPACHO
                self.tableWidget_1.setItem(fila , 8 , QTableWidgetItem(    detalle['rut']  ) )      #rut

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
                    #print(f'row: {row}, column: {column}, item={item}')
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
                rut = self.tableWidget_1.item(seleccion, 8 ).text()
                print('ventana crear orden ... para: '+ aux_tipo_doc+ ' ' + str(nro_interno) + ' | rut: ' + rut )

                fecha = datetime.now().date()
                self.fecha.setCalendarPopup(True)
                self.r_facturar_1.setChecked(False) # ASI NO MUESTRA EL "POR FACTURAR" 
                self.r_uso_interno_1.setChecked(False) # ASI NO MUESTRA EL "uso interno" 

                self.fecha.setDate(fecha)    #FECHA ESTIMADA DE ENTREGA
                self.lb_interno.setText(str(nro_interno)) #nro interno
                self.nombre_2.setText('')
                self.telefono_2.setText('')
                self.contacto_2.setText('')
                self.oce_2.setText('')
                self.r_enchape.setChecked(False)
                self.r_despacho.setChecked(False)
                self.r_dim.setChecked(False)
                self.r_carp.setChecked(False)
                self.r_elab.setChecked(False)
                self.r_pall.setChecked(False)
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
                    #----- new 
                    datos_cliente = self.conexion.root.obtener_cliente(rut)
                    print(datos_cliente)
                    if datos_cliente:
                        print('cliente encontrado --> se rellenan sus datos')
                        self.nombre_2.setText(datos_cliente[1]) #nombre del cliente
                        self.telefono_2.setText(datos_cliente[3]) #telefono  del cliente 
                        self.contacto_2.setText(datos_cliente[4]) #correo (o contacto) del cliente 
                        
                    self.stackedWidget.setCurrentWidget(self.crear_orden)
                    #------------
                    #self.tabla_ordenes = self.conexion.root.obtener_clientes_de_ordenes('dimensionado')
                    #print('cantidad de datos: ' + str(len(self.tabla_ordenes)))
                    #ordenes = pd.DataFrame(self.tabla_ordenes)
                    #nombres = ordenes[1].tolist()
                    #print(ordenes.head())
                    #print(nombres[7])
                    #completer = QCompleter( nombres )
                    #self.nombre_2.setCompleter(completer)
                    

                except EOFError:
                    QMessageBox.about(self, 'ERROR', 'Se perdio la conexion con el servidor')
                
        else:
            QMessageBox.about(self,'ERROR', 'Seleccione una Fila antes de continuar')

    def asignar_despacho(self):
        seleccion = self.tableWidget_1.currentRow()
        if seleccion != -1:
            _item = self.tableWidget_1.item( seleccion, 0) 
            if _item:            
                interno = self.tableWidget_1.item(seleccion, 0 ).text()
                aux_tipo_doc = self.tableWidget_1.item(seleccion, 1 ).text()
                estado = self.tableWidget_1.item(seleccion, 7 ).text()
                nro_interno = int(interno)
                self.inter = nro_interno

                print('-------------- ASIGNAR DESPACHO DOMICILIO ... para: '+ aux_tipo_doc+ ' ' + str(nro_interno) + '----------------')
                dialog = InputDialog3(aux_tipo_doc, str(nro_interno), estado ,self)
                if dialog.exec():
                    print('dialog cerrado')
                    new_estado = dialog.getInputs()
                    aux = 'Estado actual: '+ estado +' | Estado Obtenido: ' + new_estado
                    print(aux)

                    if estado == new_estado:
                        print('MISMOS ESTADOS, NO SE DETECTO CAMBIO')
                        QMessageBox.about(self,'INFORMACIÓN', 'NO SE DETECTARON CAMBIOS\n' + aux )
                    else:
                        print('ESTADOS DISTINTOS, CAMBIO DETECTADO')
                        try:
                            if self.conexion.root.actualizar_despacho(aux_tipo_doc, nro_interno, new_estado):
                                print('ASIGNACION DEL DESPACHO EXITOSO')
                                self.tableWidget_1.setRowCount(0)
                                QMessageBox.about(self,'EXITO', 'DESPACHO A DOMICILIO ASIGNADO CORRECTAMENTE')
                            else:
                                print('ASIGNACION DEL DESPACHO ERRONEA, intentelo mas tarde o contacte al soporte')
                                QMessageBox.warning(self,'PELIGRO', 'ASIGNACION DEL DESPACHO ERRONEA, intentelo mas tarde o contacte al soporte')
                        except EOFError:
                            self.conexion_perdida()
                print('-----------------------------------------------------------')
        else:
            QMessageBox.about(self,'ERROR', 'Seleccione una Fila antes de continuar')    
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

    def rellenar_datos_cliente(self):
        print('------- QCOMPLETE ACTIVADO --------------')
        nombre_cliente = self.nombre_2.text()
        print('Nombre: ' + nombre_cliente + ' | LEN: ' + str(len(nombre_cliente)) )
        try:
            index = self.nombres.index(nombre_cliente)
        except ValueError:
            index = -1

        if index >= 0:
            #nombre_cliente = nombre_cliente.rstrip()
            #print('Nombre: ' + nombre_cliente + ' | LEN: ' + str(len(nombre_cliente)) )
            self.telefono_2.setText(str(self.telefonos[index]))
            self.contacto_2.setText(str(self.contactos[index]))
        print('------- QCOMPLETE FIN --------------')


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
                                if self.conexion.root.actualizar_despacho(self.tipo_doc, self.inter , despacho):
                                    print('ORDEN CON DESPACHO -> CONFIRMA NOTA VENTA con DESPACHO.')
                                else:
                                    print('ORDEN CON DESPACHO ->Erro al actualizar DESPACHO A NOTA VENTA')
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
                                self.inicializar_buscar_venta()
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
                                self.inicializar_buscar_venta()
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
                                self.inicializar_buscar_venta()
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
                                self.inicializar_buscar_venta()
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
    
    
        

    
        
# ------ Funciones de buscar orden de trabajo -------------
    def inicializar_buscar_orden(self):
        self.anterior = None #SABER SI VUELVE A ESTADISTICAS O AQUI
        self.anterior2 = None #SABER SI VUELVE A GESTION REINGRESOS O AKI
        self.r_facturar_1.setChecked(False) # ASI NO MUESTRA EL "POR FACTURAR" 
        self.txt_orden.setText('')
        self.txt_cliente.setText('')
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
                    self.conexion_perdida()

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
                    self.conexion_perdida()
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
            self.conexion_perdida()

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
                    self.conexion_perdida()

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
                    self.conexion_perdida()

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
            self.conexion_perdida()

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
            
#------ Funciones para MODIFICAR ORDEN ----------- 

    def inicializar_modificar_orden(self):
        self.limpiar_varibles()
        self.txt_vendedor_5.setText('')
        self.contacto_5.setText('') 
        self.oce_5.setText('')
        self.txt_interno_5.setEnabled(True) #solo lectura
        self.txt_nro_doc_5.setEnabled(True)
        self.date_venta_5.setEnabled(True)
        self.comboBox_5.setEnabled(True)
        self.txt_vendedor_5.setEnabled(True)
        self.r_enchape_5.setChecked(False)
        self.r_despacho_5.setChecked(False)

        seleccion = self.tb_buscar_orden.currentRow()
        if seleccion != -1:
            _item = self.tb_buscar_orden.item( seleccion, 0) 
            if _item:            
                orden = self.tb_buscar_orden.item(seleccion, 0 ).text()
                self.nro_orden = int(orden)
                tipo = self.lb_tipo_orden.text()
                print('----------- MODIFICAR ORDEN ... para: '+ tipo + ' -->' + str(self.nro_orden) + ' -------------')
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
                        print('ORDEN MANUAL DETECTADA')
                        self.manual = True
                        if resultado[8]:
                            self.txt_nro_doc_5.setText( str(resultado[8]) )       #NUMERO DOCUMENTO
                            self.nro_doc = int( resultado[8] )

                        else:
                            self.txt_nro_doc_5.setText( '0' )       #NUMERO DOCUMENTO

                        if resultado[7]:
                            if resultado[7] == 'BOLETA':
                                self.comboBox_5.addItem( resultado[7] )                 #TIPO DOCUMENTO
                                self.comboBox_5.addItem('FACTURA')
                                self.comboBox_5.addItem('GUIA')
                                self.tipo_doc = resultado[7]

                            elif resultado[7] == 'FACTURA':
                                self.comboBox_5.addItem( resultado[7] )                 #TIPO DOCUMENTO
                                self.comboBox_5.addItem('BOLETA')
                                self.comboBox_5.addItem('GUIA')
                                self.tipo_doc = resultado[7]
                            elif resultado[7] == 'GUIA':
                                self.comboBox_5.addItem(resultado[7])
                                self.comboBox_5.addItem('FACTURA')  
                                self.comboBox_5.addItem('BOLETA')
                                self.tipo_doc = resultado[7]
                            elif resultado[7] == 'NO ASIGNADO':
                                self.comboBox_5.addItem('NO ASIGNADO') 
                                self.comboBox_5.addItem('FACTURA')  
                                self.comboBox_5.addItem('BOLETA')
                                self.comboBox_5.addItem('GUIA')
                        else:
                            print('no existe su tipo doc') # else compatible con la version antgua donde tipo doc era null, actualmente tipo doc es 'no asignado'
                            self.comboBox_5.addItem('NO ASIGNADO') 
                            self.comboBox_5.addItem('FACTURA')  
                            self.comboBox_5.addItem('BOLETA')
                            self.comboBox_5.addItem('GUIA')
                        if resultado[2]:
                            aux3 = datetime.fromisoformat(str(resultado[2]) )
                            self.date_venta_5.setDate( aux3 )   #FECHA DE VENTA
                        if resultado[15]:
                            self.vendedor = resultado[15]   #VENDEDOR
                            self.txt_vendedor_5.setText( resultado[15] ) 
                              
                    else: #SI NO ES MANUAL
                        print('ORDEN COMÚN DETECTADA')
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
                    if resultado[20]:
                        self.lb_vinculo_5.setText(resultado[20])
                    else:
                        self.lb_vinculo_5.setText('NO CREADO')

                    vinc = resultado[20]
                    print('Vinculo:')
                    print(vinc)
                    #self.lb_planchas.setText( str(detalle["total_planchas"]) )
                    cantidades = detalle["cantidades"]
                    descripciones = detalle["descripciones"]
                    valores_neto = detalle["valores_neto"]
                    try:
                        self.aux_vendedor = detalle["creado_por"]
                    except:
                        print('orden antigua, sin creador asignado')
                        pass

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
                        print('ORDEN MANUAL DETECTADA')
                        self.manual = True
                        if resultado[5]:
                            self.txt_nro_doc_5.setText( str(resultado[5]) )       #NUMERO DOCUMENTO
                            self.nro_doc = int( resultado[5] )
                        else:
                            self.txt_nro_doc_5.setText( '0' )       #NUMERO DOCUMENTO

                        if resultado[6]:
                            if resultado[6] == 'BOLETA':
                                self.comboBox_5.addItem( resultado[6] )                 #TIPO DOCUMENTO
                                self.comboBox_5.addItem('FACTURA')
                                self.comboBox_5.addItem('GUIA')
                                self.tipo_doc = resultado[6]

                            elif resultado[6] == 'FACTURA':
                                self.comboBox_5.addItem( resultado[6] )                 #TIPO DOCUMENTO
                                self.comboBox_5.addItem('BOLETA')
                                self.comboBox_5.addItem('GUIA')
                                self.tipo_doc = resultado[6]
                            elif resultado[6] == 'GUIA':
                                self.comboBox_5.addItem(resultado[6])
                                self.comboBox_5.addItem('FACTURA')  
                                self.comboBox_5.addItem('BOLETA')
                                self.tipo_doc = resultado[6]
                            elif resultado[6] == 'NO ASIGNADO':
                                self.comboBox_5.addItem('NO ASIGNADO') 
                                self.comboBox_5.addItem('FACTURA')  
                                self.comboBox_5.addItem('BOLETA')
                                self.comboBox_5.addItem('GUIA')
                        else:
                            print('no existe su tipo doc') # else compatible con la version antgua donde tipo doc era null, actualmente tipo doc es 'no asignado'
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
                        print('ORDEN COMÚN DETECTADA')
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

                    if resultado[19]:
                        self.lb_vinculo_5.setText(resultado[19])
                    else:
                        self.lb_vinculo_5.setText('NO CREADO')

                    vinc = resultado[19]
                    print('Vinculo:')
                    print(vinc)

                    detalle = json.loads(resultado[11])                  #DETALLE
                    cantidades = detalle["cantidades"]
                    descripciones = detalle["descripciones"]
                    valores_neto = detalle["valores_neto"]
                    try:
                        self.aux_vendedor = detalle["creado_por"]
                    except:
                        print('orden antigua, sin creador asignado')
                        pass

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
        #print(self.vendedor)
        print(self.tipo_doc)
        if not os.path.isfile(abrir):
            
            datos = ( str(self.nro_orden) ,str(self.fecha_orden.strftime("%d-%m-%Y")),self.nombre_5.text(),self.telefono_5.text(), str((self.fecha_5.date().toPyDate()).strftime("%d-%m-%Y")),cantidades,descripciones,enchapado ,self.contacto_5.text(),self.oce_5.text() ,self.vendedor )
            self.crear_pdf(datos, tipo ,despacho)
            print('pdf no encontrado, pero se acaba de crear')
        else:
            print('El pdf si existe localmente')
        print('----------------------------------------------------------------------')
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
        #v5.6
        aux_tipo = self.tipo_doc
        aux_nro = self.nro_doc

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
                            "creado_por" : self.aux_vendedor
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
                                print('pdf actualizado')
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
                            if self.manual:
                                estado = self.lb_vinculo_5.text()
                                self.actualizar_vinculo_orden_manual(aux_tipo,aux_nro,self.tipo_doc,self.nro_doc,estado)
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
 

 #  Funcion para generar REINGRESO -------------
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
                self.combo_motivos.clear()
                self.combo_soluciones.clear()
                self.txt_descripcion_2.clear()
                self.grupo_motivos.hide() #V5.8 
                self.txt_solucion_2.hide()#v5.8
                parametros_reingreso = self.conexion.root.obtener_parametros_reingreso()
                if parametros_reingreso:
                    if parametros_reingreso[1] != None:
                        detalle = json.loads(parametros_reingreso[1])
                        try:
                            lista_motivos = detalle["lista_motivos"]
                            for item in lista_motivos:
                                self.combo_motivos.addItem(item)
                        except:
                            pass
                        
                        try:
                            lista_soluciones = detalle['lista_soluciones']
                            for item in lista_soluciones:
                                self.combo_soluciones.addItem(item)
                        except:
                            pass
                
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
        motivo = self.combo_motivos.currentText() #v5.8
        '''
        if self.r_cambio_2.isChecked():
            motivo = 'CAMBIO'
        elif self.r_devolucion_2.isChecked():
            motivo = 'DEVOLUCION'
        elif self.r_otro_2.isChecked():
            motivo = self.txt_otro_2.text()'''
        
        proceso = self.lb_proceso_2.text()
        descr = self.txt_descripcion_2.toPlainText()
        solucion = self.combo_soluciones.currentText()  #v5.8
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
                        "valores_neto": valores_neto,
                        "creado_por" : self.datos_usuario[8],
                        "compatibilidad": "version-5.8"
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
                        self.stackedWidget.setCurrentWidget(self.inicio)
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
 # FUNCIONES PARA GESTION DEL REINGRESO 
    def inicializar_buscar_reingreso(self):
        self.anterior2 = True
        self.r_reingreso_2.setChecked(True)
        self.dateEdit_2.setCalendarPopup(True)
        self.dateEdit_2.setDate(datetime.now())
        self.stackedWidget.setCurrentWidget(self.buscar_reingreso)
    
    def buscar_reingreso_1(self):
        print('buscando reingreso')
        self.tb_buscar_reingreso.setRowCount(0)
        if self.r_reingreso_2.isChecked():
            print('uscando x numero reingreso')
            numero = self.txt_reingreso_2.text()
            if numero != '':
                numero = int(numero)
                print('buscando ', numero)
                consulta = self.conexion.root.obtener_reingreso_x_numero(numero)
                print(consulta)
                if consulta:
                    fila = self.tb_buscar_reingreso.rowCount()
                    self.tb_buscar_reingreso.insertRow(fila)
                    self.tb_buscar_reingreso.setItem(fila , 0 , QTableWidgetItem(str(consulta[1]))) #FECHA
                    self.tb_buscar_reingreso.setItem(fila , 1 , QTableWidgetItem(str(consulta[0]) )) #NRO REINGRESO
                    self.tb_buscar_reingreso.setItem(fila , 2 , QTableWidgetItem(str(consulta[2])))      #TIPO DOCUMENTO
                    self.tb_buscar_reingreso.setItem(fila , 3 , QTableWidgetItem(str(consulta[3]  ))) #NRO DOCUMENTO

                    self.tb_buscar_reingreso.setItem(fila , 4 , QTableWidgetItem(str(consulta[7] )   ))      #PROCESO
                    detalle = json.loads(consulta[8])
                    try:
                        creador = detalle["creado_por"]
                    except KeyError:
                        creador = 'No registrado'

                    self.tb_buscar_reingreso.setItem(fila , 5 , QTableWidgetItem( str(consulta[4] ) ))             # nro orden

                    try:
                        version = detalle["compatibilidad"]
                    except KeyError:
                        version = 'version-5.7'

                    self.tb_buscar_reingreso.setItem(fila , 6 , QTableWidgetItem( creador ))             #Vendedor
                    try:
                        estado = detalle["estado"]
                    except KeyError:
                        estado = 'VALIDA'
                    self.tb_buscar_reingreso.setItem(fila , 7 , QTableWidgetItem(  estado  ))  #estado

                    self.tb_buscar_reingreso.setItem(fila , 8 , QTableWidgetItem( version ))  #version de compatibilidad

        elif self.r_fecha_2.isChecked():
            
            print('uscando x fecha')
            fecha = self.dateEdit_2.date()
            print(fecha)
            fecha = fecha.toPyDate()
            print(fecha)

            lista_consulta = self.conexion.root.obtener_reingreso_x_fecha(str(fecha))
            if lista_consulta:
                for consulta in lista_consulta:
                    fila = self.tb_buscar_reingreso.rowCount()
                    self.tb_buscar_reingreso.insertRow(fila)
                    self.tb_buscar_reingreso.setItem(fila , 0 , QTableWidgetItem(str(consulta[1]))) #FECHA
                    self.tb_buscar_reingreso.setItem(fila , 1 , QTableWidgetItem(str(consulta[0]) )) #NRO REINGRESO
                    self.tb_buscar_reingreso.setItem(fila , 2 , QTableWidgetItem(str(consulta[2])))      #TIPO DOCUMENTO
                    self.tb_buscar_reingreso.setItem(fila , 3 , QTableWidgetItem(str(consulta[3]  ))) #NRO DOCUMENTO

                    self.tb_buscar_reingreso.setItem(fila , 4 , QTableWidgetItem(str(consulta[7] )   ))      #PROCESO
                    detalle = json.loads(consulta[8])
                    try:
                        creador = detalle["creado_por"]
                    except KeyError:
                        creador = 'No registrado'

                    self.tb_buscar_reingreso.setItem(fila , 5 , QTableWidgetItem( str(consulta[4] ) ))             # nro orden

                    try:
                        version = detalle["compatibilidad"]
                    except KeyError:
                        version = 'version-5.7'

                    self.tb_buscar_reingreso.setItem(fila , 6 , QTableWidgetItem( creador ))             #Vendedor
                    try:
                        estado = detalle["estado"]
                    except KeyError:
                        estado = 'VALIDA'
                    self.tb_buscar_reingreso.setItem(fila , 7 , QTableWidgetItem(  estado  ))  #estado

                    self.tb_buscar_reingreso.setItem(fila , 8 , QTableWidgetItem( version ))  #version de compatibilidad


        elif self.r_tipo_reingreso_2.isChecked():
            print('buscando x tipo reingreso')
            tipo = self.comboBox_2.currentText()
            print('buscando x ', tipo)

    def vista_modificar_reingreso(self):
        print('REDIRECCIONAR A VISTA MODIFICAR REINGRESO')
        seleccion = self.tb_buscar_reingreso.currentRow()
        if seleccion != -1:
            _item = self.tb_buscar_reingreso.item( seleccion, 0) 
            if _item:            
                nro_reingreso = self.tb_buscar_reingreso.item( seleccion, 1).text() 
                
                print(self.version)
                estado = self.tb_buscar_reingreso.item( seleccion, 7).text() 
                print(estado)
                print('ventana  reingreso ... para: xd' , nro_reingreso )
                self.rellenar_datos_reingreso2(nro_reingreso)
                self.tabWidget_2.setCurrentIndex(1)

                self.btn_registrar_6.hide() 

                self.btn_guardar_6.show()
                if estado == 'VALIDA':
                    self.btn_validar_reingreso.hide()
                    self.btn_anular_reingreso.show()
                elif estado == 'ANULADA':
                    self.btn_anular_reingreso.hide()
                    self.btn_validar_reingreso.show()

                self.label_65.show()
                self.lb_reingreso_6.show()
                self.lb_reingreso_6.setText(str(nro_reingreso))


                self.stackedWidget.setCurrentWidget(self.ingreso_manual)
        else:
            QMessageBox.about(self,'ERROR', 'Seleccione una fila antes de continuar')
    def rellenar_datos_reingreso2(self, id):

        self.rellenar_datos_manual()
        print('rellenando datos reingreso para: ', id)
        try:
            resultado = self.conexion.root.obtener_reingreso_x_numero(id) 
            self.nro_reingreso = id #GUARDA EL ID DEL REINGRESO PARA SU POSIBLE USO POSTERIOR (ANULAR, ACTUALIZAR, VER PDF)
            print('id reingreso copiado')
            tipo_doc = resultado[2] # 0 no asignado - 1 boleta - 2 factura - 3 guia
            if tipo_doc:
                if tipo_doc == 'BOLETA':
                    self.comboBox_6.setCurrentIndex(1)
                elif tipo_doc == 'FACTURA':
                    self.comboBox_6.setCurrentIndex(2)
                elif tipo_doc == 'GUIA':
                    self.comboBox_6.setCurrentIndex(3)

            nro_doc = resultado[3]
            self.txt_nro_doc_6.setText(str(nro_doc))
            nro_orden =  resultado[4]
            self.txt_orden_6.setText(str(nro_orden))
            motivo = resultado[5]

            if motivo == 'CAMBIO':
                self.r_cambio_6.setChecked(True)

            elif motivo == 'DEVOLUCION':
                self.r_devolucion_6.setChecked(True)
            else:
                self.r_otro_6.setChecked(True)
                self.txt_otro_6.setText(motivo)

            descripcion = resultado[6]
            self.txt_descripcion_6.appendPlainText(descripcion)
            proceso = resultado[7]
            if proceso == 'DIMENSIONADO':
                self.r_d_6.setChecked(True)
            elif proceso == 'ELABORACION':
                self.r_e_6.setChecked(True)
            elif proceso == 'CARPINTERIA':
                self.r_c_6.setChecked(True)
            elif proceso == 'PALLETS':
                self.r_p_6.setChecked(True)

            detalle = json.loads(resultado[8])
            self.aux_detalle = detalle # COPIA DEL DETALLE, PARA LUEGO AÑADIR MAS OPCIONES SIN REALIZAR OTRA CONSULTA.

            cantidades = detalle["cantidades"]
            descripciones = detalle["descripciones"]
            netos = detalle["valores_neto"]
            j = 0
            while j < len( cantidades ):
                fila = self.tb_reingreso_manual.rowCount()
                self.tb_reingreso_manual.insertRow(fila)
                self.tb_reingreso_manual.setItem(fila , 0 , QTableWidgetItem( descripciones[j] ) ) 
                self.tb_reingreso_manual.setItem(fila , 1 , QTableWidgetItem( str( cantidades[j] )  ) )
                self.tb_reingreso_manual.setItem(fila , 2 , QTableWidgetItem( str( netos[j] )  ) )
                j+=1 

            solucion = resultado[9]
            self.txt_solucion_6.appendPlainText(solucion)
            try:
                compatibilidad = detalle['compatibilidad']
                self.version = compatibilidad
            except KeyError:
                self.version = 'version-5.7'

            print(self.version)   
            if self.version == 'version-5.8':
                    print('VERSION 5.8 compatible')
                    self.grupo_motivos_2.hide()
                    self.txt_solucion_6.hide()
                    self.combo_motivos_2.show()
                    self.combo_soluciones_2.show()
                    aux = self.combo_motivos_2.findText(motivo)
                    print(aux)
                    if aux != -1:
                        self.combo_motivos_2.setCurrentIndex(aux)
                    
                    aux2 = self.combo_soluciones_2.findText(solucion)
                    print(aux2)
                    if aux2 != -1:
                        self.combo_soluciones_2.setCurrentIndex(aux2)
            else:
                print('VERSION 5.7 - error de compatibilidad')
                self.version = 'version-5.7'
                self.combo_motivos_2.hide()
                self.combo_soluciones_2.hide()
                self.grupo_motivos_2.show()
                self.txt_solucion_6.show()


                
        except EOFError:
            self.conexion_perdida()   
    def anular_validar_reingreso(self, nuevo_estado):
        if nuevo_estado == 'ANULAR':
            print('anulando reingreso: ',self.nro_reingreso)
            estado = 'ANULADA'
        elif nuevo_estado == 'VALIDAR':
            print('VALIDANDO reingreso: ',self.nro_reingreso)
            estado = 'VALIDA'

        print(self.aux_detalle)
        nuevo_detalle = self.aux_detalle
        nuevo_detalle['estado'] = estado
        fecha = datetime.now()
        fecha = fecha.strftime('%d-%m-%Y %H:%M')
        item = {
            'vendedor': self.datos_usuario[8],
            'fecha' : str(fecha) ,
            'estado': estado
        }
        print(item)
        itemx = json.dumps(item)

        historial = None
        try:
            historial = nuevo_detalle['historial']
            print('Tiene registro de cambios de estado.')
            historial.append(item)
            
        except KeyError:
            print('no tiene registro de cambios de estado.')
            historial = []
            historial.append(item)

        nuevo_detalle['historial'] = historial
        nuevo_detalle = json.dumps(nuevo_detalle)
        print(nuevo_detalle)
        try:
            if self.conexion.root.actualizar_detalle_reingreso(nuevo_detalle, self.nro_reingreso):
                self.tb_buscar_reingreso.setRowCount(0)
                self.aux_detalle = None
                self.stackedWidget.setCurrentWidget(self.buscar_reingreso)
                QMessageBox.about(self,'EXITO',f'Reingreso N° {self.nro_reingreso} {nuevo_estado} correctamente')
            else:
                QMessageBox.about(self,'ERROR','Problemas al {nuevo_estado} reingreso.\nPosiblemente no se hayan detectado cambios.\nContacte al Soporte.')
        except EOFError:
            self.conexion_perdida()


    def verificar_pdf_reingreso(self):
        seleccion = self.tb_buscar_reingreso.currentRow()
        if seleccion != -1:
            _item = self.tb_buscar_reingreso.item( seleccion, 0) 
            if _item:            
                nro_reingreso = self.tb_buscar_reingreso.item( seleccion, 1).text()
                print('verificando existencia del pdff  reingreso ... para: xd' , nro_reingreso )
                self.nro_reingreso = nro_reingreso
            
                print('Obteniendo datos reales, creando PDF ...')
                try:
                    resultado = self.conexion.root.obtener_reingreso_x_numero(self.nro_reingreso)
                    fecha = resultado[1]
                    tipo_doc = resultado[2]
                    nro_doc = resultado[3]
                    motivo = resultado[5]
                    descr = resultado[6]
                    proceso = resultado[7]
                    solucion = resultado[9]
                    detalle = json.loads(resultado[8])
                    cantidades = detalle['cantidades']
                    descripciones = detalle['descripciones']
                    valores_neto = detalle['valores_neto']

                    datos = (self.nro_reingreso, str(fecha) , tipo_doc , nro_doc , motivo , descr , proceso , solucion, cantidades, descripciones, valores_neto)
                    self.crear_pdf_reingreso(datos)
                    sleep(1)
                    self.ver_pdf_reingreso()
                except EOFError:
                    self.conexion_perdida()
                except PermissionError:
                    QMessageBox.about(self,'ERROR', 'Otro programa tiene abierto el documento PDF. Intente cerrar el documento para poder volver a visualizarlo')
                

        else:
            QMessageBox.about(self,'ERROR', 'Seleccione una fila antes de continuar')

    def decidir_atras2(self):
        if self.anterior2: 
            self.stackedWidget.setCurrentWidget(self.buscar_reingreso)
        else:
            self.stackedWidget.setCurrentWidget(self.inicio)
# --------- Funciones de INGRESO MANUAL -----------------
    def inicializar_ingreso_manual(self):
        self.txt_obs_1.clear()
        self.anterior2 = False #para hacer que vuelva al inicio enves de gestion de reingresos
        self.version = 'version-5.8'
        self.btn_guardar_6.hide() #v.5.8
        self.btn_registrar_6.show() #v.5.8
        self.label_65.hide() #v.5.8 
        self.lb_reingreso_6.hide()#v.5.8

        self.grupo_motivos_2.hide() #v.5.8
        self.txt_solucion_6.hide() #v.5.8
        self.btn_anular_reingreso.hide()
        self.btn_validar_reingreso.hide()

        self.combo_motivos_2.show()
        self.combo_soluciones_2.show()

        if self.datos_usuario[4] == 'SI' :
            self.rellenar_datos_manual()
            self.stackedWidget.setCurrentWidget(self.ingreso_manual)
        else:
            dialog = InputDialog2('CLAVE:', True ,'INGRESE CLAVE',self)
            if dialog.exec():
                clave = dialog.getInputs()
                try:
                    resultado = self.conexion.root.obtener_clave('dinamica') #v
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
        self.tb_orden_manual.setColumnWidth(1,350)
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
        self.combo_motivos_2.clear()
        self.combo_soluciones_2.clear()
        try:
            parametros_reingreso = self.conexion.root.obtener_parametros_reingreso()
            if parametros_reingreso:
                if parametros_reingreso[1] != None:
                    detalle = json.loads(parametros_reingreso[1])
                    print(detalle)
                    try:
                        lista_motivos = detalle["lista_motivos"]
                        for item in lista_motivos:
                            self.combo_motivos_2.addItem(item)
                    except KeyError:
                        print('error key motivos')
                        lista_motivos = []
        
                    try:
                        lista_soluciones = detalle['lista_soluciones']
                        for item in lista_soluciones:
                            self.combo_soluciones_2.addItem(item)
                    except KeyError:
                        lista_soluciones = []
        except:
            pass

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
                                self.conexion.root.eliminar_clave(self.clave,'dinamica')
                                self.clave = None
                            self.nro_orden = self.buscar_nro_orden(resultado)
                            self.conexion.root.actualizar_orden_dim_obser(observacion , self.nro_orden)

                            datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, enchape, cont,oce, vendedor)
                            self.crear_pdf(datos,'dimensionado', despacho )
                            self.conexion.root.actualizar_vinculo_existente('dimensionado',self.nro_orden,'NO CREADO')
                            boton = QMessageBox.question(self, 'Orden de dimensionado registrada correctamente', 'Desea abrir la Orden?')
                            if boton == QMessageBox.Yes:
                                self.ver_pdf('dimensionado')
                            self.stackedWidget.setCurrentWidget(self.inicio)
                            
                        elif self.r_elab_1.isChecked():
                            self.conexion.root.registrar_orden_elaboracion( nombre,telefono,str(fecha_orden), str(fecha),self.nro_doc,self.tipo_doc,cont,oce, despacho, interno ,detalle,f_venta,vendedor)
                            if self.clave:
                                self.conexion.root.eliminar_clave(self.clave,'dinamica')
                                self.clave = None
                            resultado = self.conexion.root.buscar_orden_elab_interno(interno)
                            self.nro_orden = self.buscar_nro_orden(resultado)
                            self.conexion.root.actualizar_orden_elab_obser(observacion , self.nro_orden)
                            datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, 'NO', cont, oce,vendedor)
                            self.crear_pdf(datos , 'elaboracion', despacho)
                            self.conexion.root.actualizar_vinculo_existente('elaboracion',self.nro_orden,'NO CREADO')
                            boton = QMessageBox.question(self, 'Orden de elaboracion registrada correctamente', 'Desea abrir la Orden?')
                            if boton == QMessageBox.Yes:
                                self.ver_pdf('elaboracion')
                            self.stackedWidget.setCurrentWidget(self.inicio)

                        elif self.r_carp_1.isChecked():
                            self.conexion.root.registrar_orden_carpinteria( nombre,telefono,str(fecha_orden), str(fecha),self.nro_doc,self.tipo_doc,cont,oce, despacho, interno ,detalle, f_venta ,vendedor)
                            if self.clave:
                                self.conexion.root.eliminar_clave(self.clave,'dinamica')
                                self.clave = None
                            resultado = self.conexion.root.buscar_orden_carp_interno(interno)
                            self.nro_orden = self.buscar_nro_orden(resultado)
                            self.conexion.root.actualizar_orden_carp_obser(observacion , self.nro_orden)
                            datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, 'NO', cont, oce, vendedor)
                            self.crear_pdf(datos , 'carpinteria', despacho)
                            self.conexion.root.actualizar_vinculo_existente('carpinteria',self.nro_orden,'NO CREADO')
                            boton = QMessageBox.question(self, 'Orden de elaboracion registrada correctamente', 'Desea abrir la Orden?')
                            if boton == QMessageBox.Yes:
                                self.ver_pdf('carpinteria')
                            self.stackedWidget.setCurrentWidget(self.inicio)

                        elif self.r_pall_1.isChecked():
                            self.conexion.root.registrar_orden_pallets( nombre,telefono,str(fecha_orden), str(fecha),self.nro_doc,self.tipo_doc,cont,oce, despacho, interno ,detalle, f_venta,vendedor)
                            if self.clave:
                                self.conexion.root.eliminar_clave(self.clave,'dinamica')
                                self.clave = None
                            resultado = self.conexion.root.buscar_orden_pall_interno(interno)
                            self.nro_orden = self.buscar_nro_orden(resultado)
                            self.conexion.root.actualizar_orden_pall_obser(observacion , self.nro_orden)
                            datos = ( str(self.nro_orden) , str(fecha_orden.strftime("%d-%m-%Y")), nombre , telefono, str(fecha.strftime("%d-%m-%Y")) , cantidades, descripciones, 'NO', cont, oce,vendedor)
                            self.crear_pdf(datos , 'pallets',despacho)
                            self.conexion.root.actualizar_vinculo_existente('pallets',self.nro_orden,'NO CREADO')
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
    def rellenar_datos_cliente_manual(self):
        print('------- QCOMPLETE MANUAL ACTIVADO --------------')
        nombre_cliente = self.nombre_1.text()
        print('Nombre: ' + nombre_cliente + ' | LEN: ' + str(len(nombre_cliente)) )
        try:
            index = self.nombres.index(nombre_cliente)
        except ValueError:
            index = -1

        if index >= 0:
            #nombre_cliente = nombre_cliente.rstrip()
            #print('Nombre: ' + nombre_cliente + ' | LEN: ' + str(len(nombre_cliente)) )
            self.telefono_1.setText(str(self.telefonos[index]))
            self.contacto_1.setText(str(self.contactos[index]))
        print('------- QCOMPLETE MANUAL FIN --------------')

 #------ funciones reingreso manual
    def reingreso_manual(self, modo):
        nro_orden = self.txt_orden_6.text()           #NUMERO DE ORDEN
        tipo_doc = self.comboBox_6.currentText() #TIPO DE DOCUMENTO
        nro_doc = self.txt_nro_doc_6.text()        #NUMERO DE DOCUMENTO
        fecha = datetime.now().date()              #FECHA DE REINGRESO
        
        if self.version == 'version-5.8':
            motivo = self.combo_motivos_2.currentText()    #MOTIVO v5.8
            solucion = self.combo_soluciones_2.currentText()  #solucion v5.8
        else:
            motivo = ''
            if self.r_cambio_6.isChecked():
                motivo = 'CAMBIO'
            elif self.r_devolucion_6.isChecked():
                motivo = 'DEVOLUCION'
            elif self.r_otro_6.isChecked():
                motivo = self.txt_otro_6.text()
            
            solucion = self.txt_solucion_6.toPlainText()   #solucion
        
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
                if modo == 'registrar':
                    formato["compatibilidad"] = "version-5.8"
                elif modo == 'actualizar':
                    formato["compatibilidad"] = self.version
                    
                detalle = json.dumps(formato)
                try:
                    nro_orden = int(nro_orden)
                    nro_doc = int(nro_doc)
                    #MODO PARA REGISTRAR LOS REINGRESOS
                    if modo == 'registrar':
                        if self.conexion.root.registrar_reingreso( str(fecha), tipo_doc, nro_doc, nro_orden, motivo, descr, proceso, detalle,solucion):
                            resultado = self.conexion.root.obtener_max_reingreso()
                            self.nro_reingreso = resultado[0]
                            print('max nro reingreso: ' + str(resultado[0]) + ' de tipo: ' + str(type(resultado[0])))
                            datos = (resultado[0], str(fecha) , tipo_doc , nro_doc , motivo , descr , proceso , solucion, cantidades, descripciones, valores_neto)
                            self.crear_pdf_reingreso(datos)
                            if self.clave:
                                    self.conexion.root.eliminar_clave(self.clave,'dinamica')
                                    self.clave = None

                            boton = QMessageBox.question(self, 'Reingreso registrado correctamente', 'Desea ver el reingreso?')
                            if boton == QMessageBox.Yes:
                                self.ver_pdf_reingreso()
                            self.stackedWidget.setCurrentWidget(self.inicio)
                        else:
                            QMessageBox.about(self,'ERROR','404 NOT FOUND. Contacte con Don Huber ...problemas al registrar reingreso')
                    #MODO PARA ACTUALIZAR LOS REINGRERSOS
                    elif modo == 'actualizar':
                        print('actualizando xdxd')
                        id = self.lb_reingreso_6.text()
                        
                        if self.conexion.root.actualizar_reingreso( id , tipo_doc, nro_doc, nro_orden, motivo, descr, proceso, detalle,solucion):
                            print('actualizado correct')
                            self.tb_buscar_reingreso.setRowCount(0)
                            self.stackedWidget.setCurrentWidget(self.buscar_reingreso)
                            QMessageBox.about(self,'EXITO',f'Reingreso N° {id} actualizado')
                        else:
                            QMessageBox.about(self,'ERROR','Problemas al actualizar reingreso.\nPosiblemente no se hayan detectado cambios.\nIntente realizar un cambio y luego proceda a actualizar')

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

#---------- FUNCIONES DE INFORME ------------
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
            encabezado = ['NUMERO REINGRESO','FECHA REINGRESO','NUMERO DE ORDEN', 'TIPO DOCUMENTO', 'NUMERO DOCUMENTO', 'PROCESO','MERCADERIA','CANTIDAD','VALOR NETO','MOTIVO','DESCRIPCIÓN' ,'SOLUCIÓN','CREADO POR','ESTADO','HISTORIAL']
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

                    try:
                        estado = detalle["estado"]
                    except KeyError:
                        estado = 'VALIDA'
                    
                    try:
                        historial = json.dumps( detalle["historial"] )
                    except KeyError:
                        historial = 'SIN REGISTRO'

                    sol =  item[9] #SOLUCION  STR
                    j = 0
                    while j < len(cant):
                        fila = [ nro_re, fr, no, td, nd, proc, merc[j], cant[j], net[j], mot, descripcion ,sol, creador,estado,historial]
                        ws.append(fila)
                        j +=1

            filas_total = ws.max_row
            tab = Table(displayName="tabla1" , ref="A1:O"+ str(filas_total) )
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
                Popen([abrir], shell=True)
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
# ---------- FUNCIONES DE ESTADISTICAS ---------
    def inicializar_estadisticas(self):
        self.box_vendedores.clear()
        self.tb_orden_manual_2.setRowCount(0)
        self.tb_orden_manual_2.setColumnWidth(0,70)#tipo orden
        self.tb_orden_manual_2.setColumnWidth(1,70)#fecha orden
        self.tb_orden_manual_2.setColumnWidth(2,60)#nro orden
        self.tb_orden_manual_2.setColumnWidth(3,65)#interno
        self.tb_orden_manual_2.setColumnWidth(4,65)#documento
        self.tb_orden_manual_2.setColumnWidth(5,65)#nro_doc
        self.tb_orden_manual_2.setColumnWidth(6,150) #cliente

        self.stackedWidget.setCurrentWidget(self.estadisticas)
        if self.manuales:
            self.rellenar_tb_manuales()
        #cargar vendedores
        try:
            datos = self.conexion.root.obtener_vendedores_activos()
            print(self.datos_usuario[8])
            self.box_vendedores.addItem('todos')
            self.box_vendedores.addItem(self.datos_usuario[8])
            if datos:
                for item in datos:
                    if item[10] != self.datos_usuario[8]:
                        self.box_vendedores.addItem(item[10])

            else:
                print('No se encontraron usuarios activos')
        except EOFError:
            self.conexion_perdida()
        
    def buscar_manuales(self):
        self.tb_orden_manual_2.setRowCount(0)
        

        tipo_orden = self.box_tipo_orden.currentText()
        vendedor = self.box_vendedores.currentText()
        print(tipo_orden)
        print(vendedor)
        self.manuales = None
        try:
            otro = {} # DETERMINA SI SE BUSCAN TODAS LAS MANUALES O SOLO INCOMPLETAS
            if self.ch_omitir_carp.isChecked():
                otro['omitir_uso_carpinteria'] = True
            if self.ch_omitir_exito.isChecked():
                otro['omitir_vinculo_exitoso'] = True
                
            self.manuales = self.conexion.root.ordenes_manuales(tipo_orden,vendedor, otro) 
            if self.manuales:
                print('RELLENANDO TABLA CON ORDENES MANUALES PRICIPAL')
                for item in self.manuales:
                    fila = self.tb_orden_manual_2.rowCount()
                    self.tb_orden_manual_2.insertRow(fila)

                    date_string = str(item[1]) #fecha orden
                    date_object = datetime.strptime(date_string, "%Y-%m-%d")
                    aux = date_object.strftime('%d-%m-%Y')

                    self.tb_orden_manual_2.setItem(fila , 0, QTableWidgetItem( str(item[0]) ) ) #fecha orden
                    self.tb_orden_manual_2.setItem(fila , 1, QTableWidgetItem( str(aux) ) ) #fecha orden
                    self.tb_orden_manual_2.setItem(fila , 2, QTableWidgetItem( str(item[2]) ) ) #nro orden
                    if item[3] != None:
                        self.tb_orden_manual_2.setItem(fila , 3, QTableWidgetItem( str(item[3]) ) ) #NRO INTERNO
                    if item[4] != None:
                        self.tb_orden_manual_2.setItem(fila , 4, QTableWidgetItem( str(item[4]) ) ) # TIPO DOCUMENTO
                    if item[5] != None:
                        self.tb_orden_manual_2.setItem(fila , 5, QTableWidgetItem( str(item[5]) ) ) #NRO DOCUMENTO

                    self.tb_orden_manual_2.setItem(fila , 6, QTableWidgetItem( str(item[6]) ) ) # CLIENTE

                    if item[7] == None:
                        self.tb_orden_manual_2.setItem(fila , 7, QTableWidgetItem( 'NO CREADO' ) )
                    else:
                        self.tb_orden_manual_2.setItem(fila , 7, QTableWidgetItem( str(item[7]) ) ) #VINCULO

                    self.tb_orden_manual_2.setItem(fila , 8, QTableWidgetItem( str(item[8]) ) ) #VENDEDOR

            self.aux_tabla = self.tb_orden_manual_2

        except EOFError:
            self.conexion_perdida()

    def rellenar_tb_manuales(self):
        self.tb_orden_manual_2.setRowCount(0)
        if self.manuales:
                print('RELLENANDO TABLA CON ORDENES MANUALES PROCESO INTERNO')
                for item in self.manuales:
                    fila = self.tb_orden_manual_2.rowCount()
                    self.tb_orden_manual_2.insertRow(fila)

                    date_string = str(item[1]) #fecha orden
                    date_object = datetime.strptime(date_string, "%Y-%m-%d")
                    aux = date_object.strftime('%d-%m-%Y')

                    self.tb_orden_manual_2.setItem(fila , 0, QTableWidgetItem( str(item[0]) ) ) #fecha orden
                    self.tb_orden_manual_2.setItem(fila , 1, QTableWidgetItem( str(aux) ) ) #fecha orden
                    self.tb_orden_manual_2.setItem(fila , 2, QTableWidgetItem( str(item[2]) ) ) #nro orden
                    if item[3] != None:
                        self.tb_orden_manual_2.setItem(fila , 3, QTableWidgetItem( str(item[3]) ) ) #NRO INTERNO
                    if item[4] != None:
                        self.tb_orden_manual_2.setItem(fila , 4, QTableWidgetItem( str(item[4]) ) ) # TIPO DOCUMENTO
                    if item[5] != None:
                        self.tb_orden_manual_2.setItem(fila , 5, QTableWidgetItem( str(item[5]) ) ) #NRO DOCUMENTO

                    self.tb_orden_manual_2.setItem(fila , 6, QTableWidgetItem( str(item[6]) ) ) # CLIENTE

                    if item[7] == None:
                        self.tb_orden_manual_2.setItem(fila , 7, QTableWidgetItem( 'NO CREADO' ) )
                    else:
                        self.tb_orden_manual_2.setItem(fila , 7, QTableWidgetItem( str(item[7]) ) ) #VINCULO

                    self.tb_orden_manual_2.setItem(fila , 8, QTableWidgetItem( str(item[8]) ) ) #VENDEDOR
        

    def filtrar_solo_incompletas(self):
        
        if self.check_incompletas.isChecked():
            self.rellenar_tb_manuales()
            print('mostrando solo incompletas')
            if self.aux_tabla != None:
                self.tb_orden_manual_2 = self.aux_tabla
                remover = []
                column = 4 #COLUMNA DEL tipo documento
                # rowCount() This property holds the number of rows in the table
                for row in range(self.tb_orden_manual_2.rowCount()): 
                    # item(row, 0) Returns the item for the given row and column if one has been set; otherwise returns nullptr.
                    tipo_doc = self.tb_orden_manual_2.item(row, column) 
                    if tipo_doc:            
                        val_tipo_doc = self.tb_orden_manual_2.item(row, column).text()
                        val_nro_doc = self.tb_orden_manual_2.item(row, column + 1).text()
                        
                        vinculo = self.tb_orden_manual_2.item(row, 7 ).text() # VINCULO 
                        #print(vinculo)
                        if vinculo == 'CREADO' :
                            remover.append(row)
                print(remover)
                k = 0
                for i in remover:
                    self.tb_orden_manual_2.removeRow(i - k)
                    k += 1

        else:
            self.rellenar_tb_manuales()
            print('mostrando todas')
    
    def continuar_orden_manual(self):
        self.anterior = "ESTADISTICAS"
        self.limpiar_varibles()
        self.txt_vendedor_5.setText('')
        self.contacto_5.setText('') 
        self.oce_5.setText('')
        self.txt_interno_5.setEnabled(True) #solo lectura
        self.txt_nro_doc_5.setEnabled(True)
        self.date_venta_5.setEnabled(True)
        self.comboBox_5.setEnabled(True)
        self.txt_vendedor_5.setEnabled(True)
        self.r_enchape_5.setChecked(False)
        self.r_despacho_5.setChecked(False)
        # fecha_orden,nro_orden,interno,tipo_doc,nro_doc,nombre,vinc_existente 
        seleccion = self.tb_orden_manual_2.currentRow()
        if seleccion != -1:
            _item = self.tb_orden_manual_2.item( seleccion, 0) 
            if _item:            
                orden = self.tb_orden_manual_2.item(seleccion, 2 ).text() #nro orden
                self.nro_orden = int(orden)
                tipo = self.tb_orden_manual_2.item(seleccion, 0 ).text() #nro orden
                print('----------- MODIFICAR ORDEN MANUAL de estadisticas... para: '+ tipo + ' -->' + str(self.nro_orden) + ' -------------')
                self.tb_modificar_orden.setColumnWidth(0,80)
                self.tb_modificar_orden.setColumnWidth(1,430)
                self.tb_modificar_orden.setColumnWidth(2,85)

                self.rellenar_datos_orden(tipo)
                self.stackedWidget.setCurrentWidget(self.modificar_orden)
        
        else:
            QMessageBox.about(self,'ERROR', 'Seleccione una FILA antes de continuar')

    def crear_grafico(self):
        for i in reversed(range(self.box_grafico.count())): 
            self.box_grafico.itemAt(i).widget().setParent(None)

        tipo = self.tipo_estadistica.currentText()
        
        tipo_grafico = self.tipo_grafico.currentText()
        titulo = ''
        datos = None
        if tipo == 'GENERALES':
            titulo = 'CANTIDAD DE ORDENES POR ÁREA'
            datos = self.conexion.root.estadisticas_total_ordenes(False)
        elif tipo == 'ORDENES MANUALES':
            titulo = 'CANTIDAD DE ORDENES MANUALES POR ÁREA'
            datos = self.conexion.root.estadisticas_total_ordenes(True)
        elif tipo == 'ORDENES X VENDEDOR':
            
            tipo_orden = self.tipo_orden_trabajo.currentText()
            titulo = "ORDENES DE " + tipo_orden + ' X VENDEDOR'
            tipo_orden = tipo_orden.lower()
            datos = self.conexion.root.estadisticas_generales(tipo_orden)
        if self.tipo_grafico.currentText() == 'SCATTER PLOT':
            titulo = 'Grafico de visualizacion previa, en desarrollo...'
        if datos:
            x_list = []
            y_list = []
            for item in datos:
               x_list.append(item[0]) 
               y_list.append(item[1]) 

            grafico = Grafico(y_list,x_list,tipo_grafico,titulo)
            self.box_grafico.addWidget(grafico)

# Sondeos
    def iniciar_sondeo(self):
        
        if self.sondeo_params['estado_sondeo_manual'] == False:
            print('iniciando sondeo manual')
            self.label_71.setText('ACTIVO')
            self.label_73.setText( str( self.sondeo_params['intervalo_sondeo_manual'] ))
            self.sondeo_params['estado_sondeo_manual'] = True
            self.executor.submit(self.sondeo_manuales)
        else:
            print('El sondeo ya fue iniciado')

    def sondeo_manuales(self):
        sleep(1)
        self.obtener_datos_sondeo(False)
        tiempo = 1
        while self.sondeo_params['estado_sondeo_manual']:
            if tiempo % self.sondeo_params['intervalo_sondeo_manual'] == 0 :
                fecha = datetime.now()
                print('sondeo realizado: ' + str(fecha))
                self.obtener_datos_sondeo(False)

                tiempo = 1
            else:
                tiempo += 1
            sleep(1)
        print('Sondeo ordenes manuales Finalizado')
    def obtener_datos_sondeo(self,actualizar_tabla):
        fecha_1 = datetime.now()
        fecha_1 =  fecha_1.strftime('%Y-%m-%d')
        otros_parametros = {
            "omitir_uso_carpinteria" :  self.sondeo_params['omitir_uso_carpinteria'] ,
            "omitir_vinculo_exitoso" : self.sondeo_params['omitir_vinculo_exitoso'] ,
            "fecha_unica": True,
            "fecha_1": str(fecha_1) 
        }
        
        self.manuales = self.conexion.root.ordenes_manuales('todos','todos',otros_parametros)
        cantidad = str(len(self.manuales))
        self.btn_notificacion1.setText(cantidad)

        if actualizar_tabla:
            self.rellenar_tb_manuales()


    def actualizar_sondeo_config(self):
        intervalo = self.txt_intervalo_sondeo_manual.text()
        print(intervalo)
        try:
            val = int(intervalo)
            if val > 30:
                self.sondeo_params['intervalo_sondeo_manual'] = val
                self.label_73.setText(intervalo)
            else:
                QMessageBox.about(self,'ERROR' ,'Intervalo minimo de 30 segundos')

        except:
            QMessageBox.about(self,'ERROR' ,'Ingrese solo numeros sobre el intervalo')

        if self.ch_omitir_uso_carp.isChecked():
            self.sondeo_params['omitir_uso_carpinteria'] = True
            self.label_75.setText('SI')
        else:
            self.sondeo_params['omitir_uso_carpinteria'] = False
            self.label_75.setText('NO')

        if self.ch_omitir_vinculo.isChecked():
            self.sondeo_params['omitir_vinculo_exitoso'] = True
            self.label_76.setText('SI')
        else:
            self.sondeo_params['omitir_vinculo_exitoso'] = False
            self.label_76.setText('NO')

    def finalizar_sondeo(self): 
        print('Finalizar sondeo')
        self.label_71.setText('INACTIVO')
        self.sondeo_params['estado_sondeo_manual'] = False

    def easter_egg(self):
        r = lambda: randint(0,255)
        print('#%02X%02X%02X' % (r(),r(),r()))
        color = '#{:02x}{:02x}{:02x}'.format(r(), r(), r())
        self.btn_notificacion1.setStyleSheet("background-color: " + color  +"")

 # -------- funcion para generar clave ------------
    def generar_clave(self):
        dialog = InputDialog2('CLAVE:', False, 'REGISTRAR CLAVE',self)
        if dialog.exec():
            clave = dialog.getInputs()
            if clave != '':
                try:
                    if self.conexion.root.registrar_clave(clave,'dinamica'):
                        QMessageBox.about(self,'EXITO' ,'CLAVE: ' + clave +' REGISTRADA')
                    else:
                        QMessageBox.about(self,'ERROR' ,'ERROR AL REGISTRAR LA CLAVE')
                except EOFError:
                    self.conexion_perdida()
            else:
                QMessageBox.about(self,'ERROR' ,'Ingrese una clave antes de continuar')

#----- funciones generales -----------
    def cargar_clientes(self, tipo_orden, tipo_estado):
        datos = []
        print('obteniendo clientes de: ' + tipo_orden)
        datos = self.conexion.root.obtener_clientes_de_ordenes(tipo_orden)
        print(str(len(datos)))
        print(str(type(datos)))
        print('orden: '+ tipo_orden + ' | len: ' + str(len(datos)) + ' | type: '+ str(type(datos)))
        #self.txt_detalle.setText('orden: '+ tipo_orden + ' | len: ' + str(len(self.datos)) + ' | type: '+ str(type(self.datos)))
        ordenes = pd.DataFrame(datos) #([0]nombre, [1]telefono, [2]contacto)
        self.nombres = ordenes[0].tolist()
        self.telefonos = ordenes[1].tolist()
        self.contactos = ordenes[2].tolist()
        model = QStringListModel(self.nombres)
        print(str(type(model)))
   
        if tipo_estado == 'normal':
            self.completer.setModel(model)
            self.completer.setMaxVisibleItems(7)
            self.completer.setCaseSensitivity(0) # 0: no es estricto con mayus | 1: es estricto con mayus
            print(self.completer.filterMode()) #filtermode podria ser: coincidencias de inicio, fin o que basta que contenga.
            self.nombre_2.setCompleter(self.completer)
        elif tipo_estado == 'manual':
            self.completer_manual.setModel(model)
            self.completer_manual.setMaxVisibleItems(7)
            self.completer_manual.setCaseSensitivity(0) # 0: no es estricto con mayus | 1: es estricto con mayus
            print(self.completer_manual.filterMode()) #filtermode podria ser: coincidencias de inicio, fin o que basta que contenga.
            self.nombre_1.setCompleter(self.completer_manual)
    
    def decidir_atras(self):
        if self.anterior: 
            self.stackedWidget.setCurrentWidget(self.estadisticas)
        else: 
            self.stackedWidget.setCurrentWidget(self.buscar_orden)

    def crear_vinculo(self,tipo):
        vinc = {
            "tipo" : tipo,
            "folio" : self.nro_orden }
        vinculo = json.dumps(vinc)
        # BUSCA SI EXISTEN VINCULOS.
        if self.conexion.root.añadir_vinculo_orden_a_venta(self.tipo_doc, vinculo,self.nro_doc):
            print('vinculo  de orden de '+ tipo +' a venta creado exitosamente')
        else:
            print('no se encontro el documento o sucedio un error')

    def crear_vinculo_v2(self,tipo):
        vinc = {
            "tipo" : tipo,
            "folio" : self.nro_orden }
        vinculo = json.dumps(vinc)
        # BUSCA SI EXISTEN VINCULOS.
        if self.conexion.root.añadir_vinculo_orden_a_venta(self.tipo_doc, vinculo,self.nro_doc):
            print('vinculo V2 de orden de trabajo '+tipo +' a venta creado exitosamente')
            self.lb_vinculo_5.setText('CREADO')
        else:
            print('no se encontro el documento o sucedio un error')
    def actualizar_vinculo_orden_manual(self,aux_tipo,aux_nro,tipo_doc,nro_doc,estado):
        x = 0
        print('------------ ACTUALIZANDO VINCULO A ORDEN MANUAL -----------')
        print('Estado vinculo: ' + estado +' - tipo orden: ' + self.tipo +' - nro_orden: '+ str(self.nro_orden))
        print(f'ANTIGUOS: - tipo: {aux_tipo} , nro_doc : {str(aux_nro)} | NUEVOS: - tipo: {tipo_doc} , nro_doc: {str(nro_doc)}')
        if nro_doc != 0:
            print('NRO ACEPTABLE...')
            if estado == 'NO CREADO':
                print('creando vinculo ......')
                new_tipo = self.tipo.lower()
                print(new_tipo)
                if tipo_doc != None or tipo_doc != 'NO ASIGNADO':
                    print('DOCUMENTO VALIDO: '+ tipo_doc)
                    if (aux_tipo == tipo_doc) and (aux_nro == nro_doc):
                        print('DOCUMENTO IDENTICO AL ANTERIOR, se cancela la vinculacion')
                    else:
                        print('DOCUMENTOS DISTINTO AL ANTERIOR, se procede a generar el vinculo')
                        self.crear_vinculo_v2(new_tipo)
                else:
                    print('TIPO DE DOCUMENTO NO VALIDO, VINCULO NO CREADO.')

            elif estado == 'CREADO':
                print('creando vinculo ......')
                if tipo_doc != None or tipo_doc != 'NO ASIGNADO':
                    print('DOCUMENTO VALIDO: '+ tipo_doc)
                    if (aux_tipo == tipo_doc) and (aux_nro == nro_doc):
                        print('DOCUMENTO IDENTICO AL ANTERIOR, se cancela la vinculacion')
                    else:
                        print('DOCUMENTOS DISTINTO AL ANTERIOR, se procede a generar el vinculo')
                        new_tipo = self.tipo.lower()
                        print(new_tipo)
                        self.crear_vinculo_v2(new_tipo)
                else:
                    print('TIPO DE DOCUMENTO NO VALIDO, VINCULO NO CREADO.')

                print('ELIMINANDO vinculo anterior...')
        else:
            print("NRO DOC NO VALIDO = 0")
            new_tipo = self.tipo.lower()
        print('-------------------------------------------------------------')


    def crear_pdf(self, lista, tipo, despacho): #PDF DE ORDEN DE TRABAJO
        ruta = ( self.carpeta +'/ordenes/' + tipo +'_' + lista[0] + '.pdf' )  #NRO DE ORDEN  
        formato = self.carpeta +"/formatos/" + tipo +".jpg"
        agua = self.carpeta + "/formatos/despacho.png"
        uso_interno = self.carpeta + "/formatos/uso interno.png"
        hojas = 2
        if tipo == 'carpinteria' or tipo == 'pallets' or tipo == 'elaboracion':
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

                    
                    if self.tipo_doc: # si existe el tipo de documento
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
                    if lista[10] != None:
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
            print('creando pdf reingreso...')
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
                
                #print(len( datos[5] ))
                k = 0 
                j = 0 
                for item in lista:
                    documento.drawString(20* mm, (92.5 + salto - k)*mm , lista[j] )  #descripcion del problema
                    #documento.drawString(20* mm, (92.5 + salto - k )*mm , descr )  #descripcion del problema
                    k += 6
                    j += 1
                
                lista = self.separar2( datos[7] , 85) #SOLUCION
                #print(len(datos[7]))
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
                if self.vendedor: #si no existe es orden manual antigua. ordenes manuales acutales tienen vendedor asociado al usuario sagot.
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
                claves = self.conexion.root.obtener_clave('dinamica')
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
                        self.conexion.root.eliminar_clave(aux_clave,'dinamica')
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
        Popen([ruta], shell=True)

    def ver_pdf(self, tipo):
        abrir = self.carpeta+ '/ordenes/' + tipo.lower() +'_' +str(self.nro_orden) + '.pdf'
        Popen([abrir], shell=True)

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
        #print('----------------------------------------------------')
        #print(cadena)
        #print('espacios necesarios: ' + str(iter))
        i = 0
        while len(cadena)> long:
        
            #print('long > '+ str(long) +':')
            aux = cadena[0:long]
            index = aux[::-1].find(' ')
        
            aux = aux[:(long-index)]
            #print('Iteracion: '+str(i)+ ': '+ aux)
            lista.append(aux)
            cadena = cadena[long - index :]
            i += 1

        if len(cadena) > 0 :
            vacias = cadena.count(' ')
            if vacias == len(cadena):
                z  = 0 
                #print('item vacio')
                #print('----------------------------------------------------')
            else:
                #print('fin long < '+ str(long) +':' + cadena)
                lista.append(cadena)
                #print('----------------------------------------------------')

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

    def limpiar_varibles(self):
        self.vendedores = None
        self.aux_tabla = None
        self.bol_fact = None # boletas y facturas encontradas
        self.guias = None    # guias encontradas
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
        self.aux_vendedor = None

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
                self.btn_inicio.show()
                self.btn_buscar.setText('Notas de venta')
                self.btn_modificar.setText('Ordenes de trabajo')
                self.btn_orden_manual.setText('Ingreso manual')
                self.btn_generar_clave.setText('Generar clave')
                self.btn_informe.setText('Generar informe')
                self.btn_atras.setText('Cerrar sesión')
                self.btn_estadisticas.setText('Estadisticas')
                self.btn_configuracion.setText('Configuración')
                self.btn_reingreso.setText(' Reingreso')
            else:
                self.btn_inicio.hide()
                self.btn_buscar.setText('')
                self.btn_modificar.setText('')
                self.btn_orden_manual.setText('')
                self.btn_generar_clave.setText('')
                self.btn_informe.setText('')
                self.btn_atras.setText('')
                self.btn_estadisticas.setText('')
                self.btn_configuracion.setText('')
                self.btn_reingreso.setText('')
                
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
            print(event.oldState())
            if event.oldState() and Qt.WindowMinimized:
                
                print('MINIMIZADA ')
                print("Window restored (to normal or maximized state)!")

            elif event.oldState() == Qt.WindowNoState and self.windowState() == Qt.WindowMaximized:
                print("Window Maximized!")
                # ----- SE ADAPTAN TODAS LAS TABLAS  -------
                #Buscar venta
                self.tableWidget_1.setColumnWidth(0,80) #interno
                self.tableWidget_1.setColumnWidth(1,80) #documento
                self.tableWidget_1.setColumnWidth(2,80) #nro doc
                self.tableWidget_1.setColumnWidth(3,125) #fecha venta
                self.tableWidget_1.setColumnWidth(4,200) #cliente
                self.tableWidget_1.setColumnWidth(5,200) #vendedor
                self.tableWidget_1.setColumnWidth(6,100) #total
                self.tableWidget_1.setColumnWidth(7,120) #despacho domicilio
                #Crear venta
                #Buscar orden
                #Modificar orden
                #reingreso
                #ORDEN MANUAL
                #REINGRESO MANUAL
                #estadisticas
                self.tb_orden_manual_2.setColumnWidth(5,320)

    def closeEvent(self, event):
        self.finalizar_sondeo()
        print('Cerrando aplicacion')
        event.accept()

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

class InputDialog3(QDialog):
    def __init__(self,tipo_doc, folio, estado , parent=None ):
        super().__init__(parent)
        uic.loadUi('upd_domicilio.ui',self)
        self.btn_guardar.clicked.connect(self.guardar)
        self.btn_cancelar.clicked.connect(self.cancelar)
        self.lb_folio.setText(folio)
        self.lb_doc.setText(tipo_doc)
        self.lb_estado.setText(estado)


    def guardar(self):
        self.accept()
    def cancelar(self):
        self.reject()
    def getInputs(self):
        return self.comboBox.currentText()

class Canvas(FigureCanvas):
    def __init__(self,parent,x,y, tipo_grafico,titulo):
        self.fig , self.ax = plt.subplots(figsize=(6, 4), constrained_layout=True )
        
        super().__init__(self.fig)
        self.setParent(parent)
        self.x = x
        self.y = y
        print(x) #labels
        print(y) #cantidades
        # MATPLOTLIB SCRIPT
        self.ax.set_title(titulo, fontsize=12)
        self.dibujar_grafico(tipo_grafico)

    def dibujar_grafico(self,tipo):
        if tipo == 'BARRAS':
            print('DIBUJANDO GRAFICO DE BARRAS')
            self.ax.barh(self.x, self.y)
        elif tipo == 'CIRCULAR':
            print('DIBUJANDO GRAFICO CIRCULAR')
            self.ax.pie(self.y ,  labels=self.x, autopct='%1.1f%%',shadow=True, startangle=90)
            self.ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
        elif tipo == 'LINEAL':
            print('DIBUJANDO GRAFICO LINEAL')
            self.ax.plot(self.x,self.y)
            plt.xticks(rotation=90)            
        elif tipo == 'STEP':
            self.ax.stem(self.x, self.y) 
            plt.xticks(rotation=90) 
        elif tipo == "SCATTER PLOT":
            np.random.seed(19680801)

            # Compute areas and colors
            N = 150
            r = 2 * np.random.rand(N)
            theta = 2 * np.pi * np.random.rand(N)
            area = 200 * r**2
            colors = theta

            self.ax = self.fig.add_subplot(projection='polar')
            c = self.ax.scatter(theta, r, c=colors, s=area, cmap='hsv', alpha=0.75) 
        
class Grafico(QWidget):
    def __init__(self, x, y, tipo_grafico,title):
        super().__init__()
        chart = Canvas(self, x , y,tipo_grafico,title)


    
if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('icono_imagen/icono_barv3.png'))
    myappid = 'madenco.personal.area' # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid) 
    vendedor = Vendedor()
    vendedor.show()
    sys.exit(app.exec_())
