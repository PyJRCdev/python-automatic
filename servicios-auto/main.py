from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.mime.text import MIMEText

# Ruta a EdgeDriver
edgedriver_path = './edgedriver/msedgedriver.exe'
download_dir = './Download'  # Cambia esta ruta al directorio donde se descargará el archivo
url = 'https://portal.ismgroup.es/'

# Función para iniciar sesión y navegar
def download_excel_with_selenium(email, password):
    # Configuración de Microsoft Edge en modo headless (sin interfaz gráfica)
    options = webdriver.EdgeOptions()
    options.add_argument('--headless')
    options.add_argument('--disable-gpu')  # Recomendado para el modo headless
    options.add_argument('--no-sandbox')
    prefs = {"download.default_directory": download_dir}
    options.add_experimental_option("prefs", prefs)

    # Iniciar Edge en modo headless
    service = Service(executable_path=edgedriver_path)
    driver = webdriver.Edge(service=service, options=options)
    driver.get(url)

    # Esperar a que cargue la página y hacer el inicio de sesión
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'email')))
    
    email_input = driver.find_element(By.NAME, 'email')
    email_input.send_keys(email)
    
    password_input = driver.find_element(By.NAME, 'password')
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)

    # Esperar a que cargue la página después de iniciar sesión
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//button[text()="Matic"]')))
    
    # Navegar por los botones
    driver.find_element(By.XPATH, '//button[text()="Matic"]').click()
    time.sleep(2)

    driver.find_element(By.XPATH, '//button[text()="Seguimiento"]').click()
    time.sleep(2)
    
    driver.find_element(By.XPATH, '//button[text()="Niveles de servicio"]').click()
    time.sleep(2)

    # Seleccionar la opción "Proyecto" del selector
    project_selector = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//select[@name="proyecto"]'))
    )
    project_selector.click()
    project_option = driver.find_element(By.XPATH, '//option[text()="Proyecto"]')
    project_option.click()

    # Descargar el archivo Excel
    download_button = driver.find_element(By.XPATH, '//button[text()="Descargar"]')
    download_button.click()

    # Esperar a que se complete la descarga
    time.sleep(10)  # Ajusta según el tiempo que toma la descarga
    driver.quit()

# Función para enviar el correo
def send_email(smtp_server, port, sender_email, password, receiver_email, subject, body, attachment):
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with open(attachment, 'rb') as f:
        mime = MIMEBase('application', 'octet-stream')
        mime.set_payload(f.read())
        encoders.encode_base64(mime)
        mime.add_header('Content-Disposition', f'attachment; filename={os.path.basename(attachment)}')
        msg.attach(mime)

    with smtplib.SMTP(smtp_server, port) as server:
        server.starttls()
        server.login(sender_email, password)
        server.send_message(msg)

email = input("Introduce tu email: ")
password = getpass.getpass("Introduce tu contraseña: ")

# Variables del proceso
email = 'tu_email@empresa.com'
password = 'tu_contraseña'
smtp_server = 'smtp.tuempresa.com'
port = 587
sender_email = 'tu_email@empresa.com'
receiver_email = 'destinatario@empresa.com'
subject = 'Reporte de niveles de servicio'
body = 'Adjunto encontrarás el reporte de niveles de servicio.'
attachment = os.path.join(download_dir, 'archivo.xlsx')  # Cambia el nombre según corresponda

# Ejecutar el proceso
download_excel_with_selenium(email, password)
send_email(smtp_server, port, sender_email, password, receiver_email, subject, body, attachment)

print("El correo ha sido enviado con éxito.")
