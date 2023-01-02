#Librerías para tener en cuenta
import tkinter.messagebox
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from webdriver_manager.firefox import GeckoDriverManager
import time as tm

#Archivo excel para mensajes de texto
message_control = r'C:\project\python\Envio_masivo_mensajes_texto\message_control.xlsx'
df = pd.read_excel(message_control, sheet_name='message_simple')

# Configuraciones varias y opcionales
print(datetime.now(),'Gracias por ejecutar este script creado por Sebastián Beltrán de HDTC.')
print('Puede cerrar el proceso en cualquier momento con combinación de teclas Ctrl + C.')
print('VER: 1.0.0 Script para el envío masivo de menasjería de texto a dispositivos')
print('móviles. Se requiere contar con crédito antes de realizar esta prueba')
print('Este código puede ser usado para otros propósitos bajo su responsabilidad')
print(datetime.now(),'Script ejecutado por el usuario')

tkinter.messagebox.showinfo('Servicio de envío masivo de mensajes GSM - Antes de comenzar:','Se recomienda tener instalado, configurado y abierto la aplicación de mensajería GSM Google Messages en el dispositivo móvil para el escaneo del código QR que genere la página web. Cuenta con 15 segundos promedio para la configuración del servicio. Se recomienda contar con buen servicio de internet, gracias. Ver tutorial en el video')

# URL Google Messages
url1 = 'https://messages.google.com/web/authentication'
print(datetime.now(), 'Abriendo página web', url1)
broswer = webdriver.FirefoxOptions()

# Opcones del navegador para la apertura en modo privado
broswer.add_argument('--Private')
print(datetime.now(), 'Preparando navegador')
driver = webdriver.Firefox(executable_path=GeckoDriverManager().install(), options=broswer)
driver.maximize_window()

# Definir sesión activa
def sesion_active():
    driver.find_element(By.CLASS_NAME, 'mat-mdc-button-touch-target')

    

# Definir existencia código QR
def codigoQR():
    try:
        driver.find_elements(By.CLASS_NAME, 'qr-code ng-trigger ng-trigger-fadeIn ng-tns-c94-0 ng-star-inserted')
    except driver:
        return False
    return True
    

# Apertura página web messages Google
def message_google():
    driver.get(url1)
    tm.sleep(5)

    wait = True

    while wait:
        print(datetime.now(),'Por favor escanea el código QR presentado en el navegador web. Cuenta con 15 segundos')
        tm.sleep(15)
        active = sesion_active()
        if wait != active:
            print(datetime.now(), 'Sistema emparejado en celular - página web. Se están cargando los elementos')
            break


message_google()

# Cajón para mensajes
def send_GSM():
    for i in df.index:
        
        # Digitar el número de celular de 10 dígitos y el contenido del mensaje
        numero_cel = str(df['Contacto'][i])
        messages = str(df['Message'][i])

        ## Parámetros para el procedimiento del mensaje
        print(datetime.now(),'Preparando mensaje al número de celular', numero_cel)
        # Seleciones proceso general del chat. Fecha actualizado: 23/12/2022
        bolsa_chat = '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-main-nav/div/mw-fab-link/a/span[2]/div'
        contacto = '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/mw-new-conversation-sub-header/div/div[2]/mw-contact-chips-input/div/mat-chip-listbox/span/input'
        seleccion_contacto = '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-new-conversation-container/div/mw-contacts-list/div/mw-contact-row[1]/div'
        enviar_mensaje_1 = '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/div[2]/div/mws-autosize-textarea/textarea'
        enviar_mensaje_2 = '/html/body/mw-app/mw-bootstrap/div/main/mw-main-container/div/mw-conversation-container/div/div[1]/div/mws-message-compose/div/mws-message-send-button/div/mw-message-send-button/button/span[5]'
        driver.find_element(By.XPATH, bolsa_chat).click()
        tm.sleep(1)
        print(datetime.now(),'Seleccionando chat')
        # Selección número de celular
        driver.find_element(By.XPATH, contacto).send_keys(numero_cel)
        tm.sleep(3)
        driver.find_element(By.XPATH, seleccion_contacto).click()
        tm.sleep(4)
        # Enviar mensaje
        driver.find_element(By.XPATH, enviar_mensaje_1).send_keys(messages)
        tm.sleep(1)
        driver.find_element(By.XPATH, enviar_mensaje_2).click()
        print(datetime.now(),'Mensaje procesado del archivo excel de manera exitosa al usuario',numero_cel,', se recomienda tener crédito disponible con su operador móvil para el envío final')
        # Programar el tiempo estimado para el envío del mensaje
        # Se debe tener en cuenta el cálculo del tiempo estimado
        # menos el tiempo programado en el cajón de mensajes
        # Para en este caso, se restaron 7 segundos para el
        # mejor desempeño del script con la página web y PC.
        # Se debe considerar la conexión a internet a la fecha.
        tm.sleep(49) # <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
        # Este es el tiempo para configurar de acuerdo con los
        # ítems dados anteriormente. Todos los tiempos los puede
        # ajustar de acuerdo con el rendimiento general del PC.

send_GSM()
print(datetime.now(), 'Proceso realizado correctamente, gracias por la ejecución de este código script con responsabilidad. Grupo HD TEC-CO S.A.S.')

driver.close()

# Código proporcionado por Grupo HD TEC-CO S.A.S. y es líbre para su uso.