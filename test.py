import qrcode
from PIL import Image
from app.modulos.helpers import Imagen

def crear_qrcode():
    # URL para el código QR
    url = "http://madenco.site"

    # Crear el código QR
    qr = qrcode.QRCode(
        version=1, 
        error_correction=qrcode.constants.ERROR_CORRECT_L, 
        box_size=30, 
        border=5)

    qr.add_data(url)
    qr.make(fit=True)

    # Crear una imagen del código QR
    qr_img = qr.make_image(fill_color="black", back_color="white")
    print(type(qr_img))
    # Cargar el logo de la empresa
    logo_path = "app/icono_imagen/qr_logo_enco_2.png"  # Ruta al archivo de imagen del logo
    logo = Image.open(logo_path).convert("RGBA")
    logo = logo.resize( (150,150) )
    # Obtener las dimensiones del código QR y el logo
    qr_width, qr_height = qr_img.size
    print("|(qr size): ", qr_img.size)
    logo_width, logo_height = logo.size
    print("|(logo size): ", logo.size)

    # Calcular la posición para colocar el logo en el centro del código QR
    position = ((qr_width - logo_width) // 2, (qr_height - logo_height) // 2)

    # Pegar el logo en el centro del código QR
    qr_img.paste(logo, position)

    # Guardar la imagen con el código QR y el logo
    output_path = "app/icono_imagen/qrcode.png"
    qr_img.save(output_path)

    print("Código QR con logo generado y guardado en:", output_path)
def test_voucher_vale():

    params = dict(
            nombre = "jose gutierrez",
            direccion = "11 de septiembre, 4560",
            referencia = "don guti",
            telefono = "+56 19028822",
            contacto = 'gitu@gmalin.com',
            fecha_estimada = "2023-01-22"
            ) 
    Imagen.crear_voucher(params)
    
if __name__ == "__main__": 
    x = 0
    test_voucher_vale()
