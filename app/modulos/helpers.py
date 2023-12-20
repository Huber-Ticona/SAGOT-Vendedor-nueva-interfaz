
from PIL import Image, ImageFont, ImageDraw, ImageOps, ImageWin
import win32ui
from datetime import datetime

class Imagen():
    valor = 222
    @staticmethod #No puede acceder a atributos de clase ().
    def hola():
        print('hola desde imagen')
    @classmethod #Puede acceder a los atributos de clase (self).
    def hola2(self):
        print('hola22222222 desde imagen, val: ' , self.valor)

    @staticmethod
    def crear_voucher(datos):
        """ Parametros voucher Despacho """
        """ params = dict(
            nombre = "jose gutierrez",
            direccion = "11 de septiembre, 4560",
            referencia = "don guti",
            telefono = "+56 19028822",
            contacto = 'gitu@gmalin.com',
            fecha_estimada = "2023-01-22"
            ) """
        params = datos
        fecha_estimada = datetime.fromisoformat(str(datos["fecha_estimada"])) 
        fecha_estimada = str(fecha_estimada.strftime("%d-%m-%Y %H:%M:%S"))
        params["fecha_estimada"] = fecha_estimada

        print("|--Params: " ,params)
        print(f"|-- Len: {len(params)}")
        for key,value in params.items():
            print(f"|  key: {key} | Value: {value} ")


        dir_imagenes = "app/icono_imagen"
        nombre_imagen_final = "test.png"

        #imagen = Image.open("app/icono_imagen/madenco logo.png")
        #INSERTANDO FECHA,tipo y folio
        #draw = ImageDraw.Draw(imagen)

        mode = 'RGBA'
        size = (900, 550)
        background_color = (255, 255, 255, 255)  # Color blanco para imágenes RGBA
        img = Image.new(mode, size, background_color)
        
        """ Logo madenco """
        logo = Image.open(f"{dir_imagenes}/ENCO_Log.png").convert("RGBA")
        margin_left = 30
        logo = logo.resize( (300,150) )
        # Definir la posición de pegado
        posicion_pegado = (10,10)  # Coordenadas (x, y) de la esquina superior izquierda
        
        # Pegar la imagen dentro de la imagen principal
        img.paste(im= logo, box= posicion_pegado , mask= logo.split()[3])

        """ Codigo QR v6.08 
        qr_code = Image.open(f"{dir_imagenes}/qr_madenco.png").convert("RGBA")
        
        qr_code = qr_code.resize( ( int(qr_code.size[0] /6 ) , int(qr_code.size[1] /6) ) )
        posicion_pegado = (700,200)  # Coordenadas (x, y) de la esquina superior izquierda
        img.paste(im= qr_code, box= posicion_pegado , mask= qr_code.split()[3])
        """
        # Se toma la plantilla para poder editarla con objetos.(canvas) texto u imagenes.
        draw = ImageDraw.Draw(img)
        # Definir las coordenadas del rectángulo
        x1, y1 = 0, 0  # Esquina superior izquierda
        x2, y2 = size[0] , size[1]  # Esquina inferior derecha

        # Definir el color del borde del rectángulo (negro en este caso)
        border_color = (0, 0, 0)    
        grosor_rectangulo = 5
        # Dibujar el rectángulo
        draw.rectangle([x1, y1, x2, y2], outline=border_color, width=grosor_rectangulo)

        font = ImageFont.truetype(dir_imagenes + '/Helvetica.ttf', 30)
        # draw.text((x, y),"Sample Text",(r,g,b))
        margin_salto_linea = 60

        # Titulo
        draw.text( (330, margin_salto_linea ),f"DESPACHO A DOMICILIO", (0,0,0) ,font=font)
        dim_texto = font.getsize("DESPACHO A DOMICILIO")
        print("font, size texto: ", font.getsize("DESPACHO A DOMICILIO"))
        pos_subrayado = [(330,margin_salto_linea + dim_texto[1] + 3 ), (330 + dim_texto[0] , margin_salto_linea + dim_texto[1] + 3) ]
        draw.line(pos_subrayado, fill ="black", width = 5)

        # CONTENIDO
        margin_salto_linea = 200
        
        for key,value in params.items():
            draw.text( (margin_left, margin_salto_linea ),f"{key.upper()}:  {value.upper()}", (0,0,0) ,font=font)
            margin_salto_linea += 40


        filename = f"{dir_imagenes}/{nombre_imagen_final}"
        img = img.rotate(-90,expand=True) #Se rota
        img.save(filename)
        print("|-- Imagen final guardada ...")

    @staticmethod
    def crear_qr(datos):
        x = 0

    @staticmethod
    def imprimir_voucher(nombre_impresora):
        print("Imprimiendo...")
        
        file_name = "app/icono_imagen/test.png"
        #printer_name = "192.168.1.94"
        printer_name = nombre_impresora
        print(f"(NOMBRE IMPRESORA): {printer_name}")

        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)

        printable_area = (576, 900) #tamaño original x 1.5
        printer_size = (576,900)
        printer_margins = (0,0)

        print("|Area pintable: ",printable_area)
        print("|Tamaño de la impresora: ",printer_size)

        # Abra la imagen, gírela si es más ancha que
        # es alto, y calcule cuánto multiplicar 
        # por cada píxel para hacerlo lo más grande posible en 
        # la página sin distorsionar. 
        # 

        bmp = Image.open(file_name) #se toma una imagen vertical donde claramente y > x.
        print(bmp.size)
        #if bmp.size[0] > bmp.size[1]: # X > Y
        #  bmp = bmp.rotate (90)

        ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
        scale = min (ratios)

        #
        ## Inicie el trabajo de impresión y dibuje el mapa de bits en 
        # el dispositivo de impresión en el tamaño escalado. 

        hDC.StartDoc (file_name)
        hDC.StartPage ()

        dib = ImageWin.Dib (bmp)
        scaled_width, scaled_height = [int (scale * i) for i in bmp.size]

        x1 = int ((printer_size[0] - scaled_width) / 2)
        y1 = int ((printer_size[1] - scaled_height) / 2)
        x2 = x1 + scaled_width
        y2 = y1 + scaled_height

        dib.draw (hDC.GetHandleOutput (), (x1, y1, x2, y2))

        hDC.EndPage ()
        hDC.EndDoc ()
        hDC.DeleteDC ()
        print("Impresion finalizada")
        return True , "Voucher Impreso con exito."
        """ except Exception as e:
            print("ERROR: ",e)
            return False , str(e) """
        

class Funciones():

    @staticmethod
    def validar_telefono(code,numero):
        # Eliminar espacios en blanco al inicio y al final de la cadena
        code = code.strip()
        print("| Validando codigo telefonico:", str(code))

        # Verificar si el código tiene el formato correcto
        if not code.startswith("+"):
            print("| Error: El código debe comenzar con '+'")
            return False , "Error: El código debe comenzar con '+'"

        # Verificar que solo hay números después del signo más y como máximo 3 dígitos
        numeros_despues_del_mas = code[1:]
        if not numeros_despues_del_mas.isdigit() or len(numeros_despues_del_mas) > 2:
            print("| Error: Números de país como maximo de 2 digitos.")
            return False , "Error: Números de país como maximo de 2 digitos."

        # La validación fue exitosa
        print("| Código de país válido")
        if not numero.isdigit():
            print("| Error: El numero telefonico deben ser solo numeros.")
            return False, "Error: El numero telefonico deben ser solo numeros."
        
        return True , "Telefono validado correctamente."

