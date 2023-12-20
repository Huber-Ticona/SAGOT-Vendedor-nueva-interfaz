from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QPixmap , QIcon
import sys
import ctypes
from vendedor import Vendedor
from app.modulos.helpers import Imagen
if __name__ == '__main__':
    print("( CREANDO APP ... )")
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon('app/icono_imagen/icono_barv3.png'))
    myappid = 'madenco.personal.area' # arbitrary string
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid) 
    print("( INICIANDO APP VENDEDOR ... )")
    vendedor = Vendedor()
    vendedor.show()
    sys.exit(app.exec_())  
    """ 

    name = "192.168.1.94"
    Imagen.imprimir_voucher(name) """
