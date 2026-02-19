"""
Tarea para realizar pedidos de robots a partir de un fichero CSV, guardar los resultados en PDFs y crear un ZIP con los PDFs generados.
Páginas de interés:
- https://sema4.ai/docs/automation/courses/build-a-robot-python 
- https://pypi.org/project/robocorp-browser/
- https://pypi.org/project/robocorp/
- https://sema4.ai/docs/automation/python/robocorp/robocorp-browser/api
- https://sema4.ai/docs/automation/python/rpa-framework

"""

from robocorp.tasks import task
from robocorp import browser

from RPA.HTTP import HTTP
from RPA.PDF import PDF

import csv
import os
import zipfile

@task
def order_robots():
    """
    Realiza pedidos de robots a https://robotsparebinindustries.com/#/robot-order a partir de un fichero CSV con el formato:
    Order number,Head,Body,Legs,Address
    Pasos:
    1. Leer el fichero CSV.
    2. Para cada línea del CSV, realizar un pedido a la web.
    3. Guardar el resultado de cada pedido en un fichero PDF.
    4. Guarda un pantallazo del robot pedido.
    5. Embebe el pantallazo en el PDF junto con los detalles del pedido.
    6. Repite el proceso para cada línea del CSV.
    7. Crea un fichero ZIP con todos los PDFs generados.
    """
    browser.configure(
        slowmo=100,
        viewport_size={1600, 1200}
    )
    open_robots_website()
    orders = download_and_read_csv("https://robotsparebinindustries.com/orders.csv")
    for order in orders:
        place_order(order)
        store_receipt_as_pdf(order['Order number'])
        take_screenshot(order['Order number'])
        save_pdf(order['Order number'])
        next_order()
    create_zip("orders_pdfs.zip") 

def open_robots_website():
    print("Abriendo la página de pedidos de robots...")
    browser.goto("https://robotsparebinindustries.com/#/robot-order")

def download_and_read_csv(file_url):
    """ Descargamos y leemos el fichero CSV con los pedidos de robots. """
    print("Descargando el archivo CSV con los datos de pedidos de robots...")
    http = HTTP()
    file_name = file_url.split('/')[-1]
    http.download(url=file_url, target_file=f"./output/{file_name}", overwrite=True)

    file_path = f"./output/{file_name}"

    print(f"Leyendo el fichero CSV: {file_path}")
    orders = []
    with open(file_path, newline='') as csvfile:
        reader = csv.DictReader(csvfile)
        counter = 0
        for row in reader:
            orders.append(row)
            counter += 1
        print(f"Total de pedidos leídos: {counter}")
    return orders

def place_order(order):
    print(f"Realizando pedido: {order['Order number']}")
    # Campos Order number,Head,Body,Legs,Address
    page = browser.page()
    page.click("button:text('I guess so...')") # Aceptar cookies
    page.select_option("#head", str(order["Head"]))
    page.click("#id-body-"+str(order["Body"]))
    page.fill("//input[@placeholder='Enter the part number for the legs']", str(order["Legs"]))
    page.fill("#address", str(order["Address"]))
    page.click("button:text('Preview')")

    for i in range(5):  # Intentar hacer clic en el botón "ORDER" hasta 5 veces
        try:
            page.click("button:text('ORDER')")
            page.wait_for_selector("DIV #order-completion", timeout=1000)
            print(f"Pedido completado correctamente para el pedido {order['Order number']}")
            break
        except Exception as e:
            print(f"Intento {i+1}: Error al hacer el pedido, reintentando... {e}")

def store_receipt_as_pdf(order_number):
    print(f"Guardando el recibo del pedido {order_number} como PDF...")

    page = browser.page()
    receipt_html = page.locator("DIV #receipt").inner_html()

    receipt_title = f"<h1>Recibo del pedido {order_number}</h1>"
    

    pdf = PDF()
    pdf.html_to_pdf(receipt_title + receipt_html, f"output/receipt_{order_number}.pdf")

def take_screenshot(order_number):
    print("Tomando pantallazo del pedido...")
    screenshot_path = f"./output/screenshot_{order_number}.png"
    browser.page().locator("DIV #robot-preview-image").screenshot(path=screenshot_path)
    return screenshot_path

def save_pdf(order_number):
    pdf = PDF()

    list_of_files = [
        f'./output/receipt_{order_number}.pdf',
        f'./output/screenshot_{order_number}.png:align=center'
    ]
    pdf.add_files_to_pdf(
        files=list_of_files,
        target_document=f"output/receipt_complete_{order_number}.pdf"
    )

def next_order():
    print("Preparando para el siguiente pedido...")
    page = browser.page()
    page.click("button:text('Order another robot')")

def create_zip(zip_name):
    print(f"Creando el archivo ZIP: ./output/{zip_name} con los PDFs generados...")
    with zipfile.ZipFile(f"./output/{zip_name}", 'w') as zipf:
        for foldername, subfolders, filenames in os.walk('output/'):
            for filename in filenames:
                if filename.startswith('receipt_complete_') and filename.endswith('.pdf'):
                    file_path = os.path.join(foldername, filename)
                    zipf.write(file_path, os.path.relpath(file_path, 'output/'))
    print(f"Archivo ZIP '{zip_name}' creado exitosamente.")