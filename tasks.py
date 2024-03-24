from robocorp.tasks import task
from robocorp import browser

from RPA.PDF import PDF
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files
from RPA.Tables import Tables
from RPA.Archive import Archive

import time
import os, shutil

@task
def order_robot_from_robotsparepin():
    """
        Order robots from a csv file from RobotSpareBin Industries Inc.
        Save the order HTML result as a PDF and embed a screenshot for each order.
        Generate a ZIP file with the receipts and the images.
    """

    browser.configure(
        slowmo=100
    )

    open_robot_order_website()
    get_csv_data()
    process_csv_data()
    archive_receipts()
    clean_output_folder()

def open_robot_order_website():
    browser.goto('https://robotsparebinindustries.com/#/robot-order')


def fill_order_form(order):
    page = browser.page()

    page.click('//*[contains(text(), "OK")]')

    page.select_option('#head', order['Head'])
    page.click('#id-body-' + order['Body'])
    page.fill('//html/body/div/div/div[1]/div/div[1]/form/div[3]/input', order['Legs'])
    page.fill('//html/body/div/div/div[1]/div/div[1]/form/div[4]/input', order['Address'])
    page.click('#order')
    time.sleep(0.7)


    while True:
        can_order_another = page.query_selector('#order-another')
        if not can_order_another:
            page.click('#order')
        else:
            time.sleep(0.7)
            take_screenshot(order['Order number'])
            generate_pdf(order['Order number'])
            embed_screenshot_to_receipt(order['Order number'])
            page.click('#order-another')
            break




def get_csv_data():
    http = HTTP()
    http.download(url='https://robotsparebinindustries.com/orders.csv', overwrite=True)

def process_csv_data():
    csv = Tables()
    data = csv.read_table_from_csv(path='orders.csv', header=True)
    for row in data:
        fill_order_form(row)


def generate_pdf(order_number):
    page = browser.page()
    pdf = PDF()
    html_receipt = page.locator('#receipt').inner_html()
    pdf.html_to_pdf(html_receipt, f'output/receipts/receipt_{str(order_number)}.pdf', margin=15)

def take_screenshot(order_number):
    page = browser.page()
    page.screenshot(path=f'output/receipts/ss_{str(order_number)}.png')


def clean_output_folder():

    folder = './output/receipts'
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print('Failed to delete %s. Reason: %s' % (file_path, e))

def embed_screenshot_to_receipt(order_number):
    pdf = PDF()
    pdf.add_files_to_pdf(
        [f'output/receipts/ss_{str(order_number)}.png'],
        target_document=f'output/receipts/receipt_{str(order_number)}.pdf',
        append=True
    )

def archive_receipts():
    archive = Archive()
    archive.archive_folder_with_zip('./output/receipts', 'receipts.zip')
