### Criado por Daniloeus ###
### script para gerar QRcode ###
import qrcode

qr_information = "inserir informações do QR code"
image = qrcode.make(qr_information)
image.save("QR_code001.png", "PNG")