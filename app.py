import os
import openpyxl
import pandas as pd
import smtplib
import mimetypes
from email.message import EmailMessage
from analyticskit import *

pd.options.display.float_format = '{:,.2f}'.format # Display option to include 2 decimal places on all floats
pd.set_option('display.max_columns', None) # Display option to show all columns
pd.set_option('display.max_rows', None) # Display option to show all rows

while True:
    directory = single_input("What's the path to the directory where ProductData.xlsx, MasterInvoice.pdf, and Order files are located?")
    if directory[-1] != '\\':
        directory = directory + '\\'
    try:
        filenames = os.listdir(directory)
    except:
        filenames = []
    if "ProductData.xlsx" in filenames and "MasterInvoice.pdf" in filenames and "Order1.xlsx" in filenames and "Order2.xlsx" in filenames and "Order3.xlsx" in filenames and "Order4.xlsx" in filenames:
        break
    else:
        show_message("Error", "The filepath you entered is invalid. Please try again.")

product_descriptions = {}
product_prices = {}

df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Descriptions')
for index, row in df.iterrows():
    product_descriptions[row['Unit number']] = row['Description']

df = pd.read_excel(directory + 'ProductData.xlsx', sheet_name='Prices')
for index, row in df.iterrows():
    product_prices[row['Unit number']] = row['Price per unit']

def get_order_data(invoice_num, filename, directory):
    #open excel file
    template = openpyxl.load_workbook(directory + filename)
    selected_sheet = template['OrderForm']
    # get date and format it correctly as string
    date = selected_sheet['C9'].value
    date = date.strftime('%m/%d/%y')

    #get name of customer
    name = selected_sheet['C14'].value
    # get customer email
    email = selected_sheet['C13'].value

    #dictionary with order data
    order_data = {}
    order_data['InvoiceNum'] = str(invoice_num)
    order_data['Date'] = date
    order_data['CustomerAddress1'] = selected_sheet['C16'].value
    order_data['CustomerName'] = name
    order_data['CustomerCompany'] = selected_sheet['C15'].value
    order_data['CustomerAddress2'] = f"{selected_sheet['C17'].value}, {selected_sheet['C18'].value} {selected_sheet['C19'].value}"

    total = 0

    if selected_sheet['B24'].value is not None:
        order_data['Price0'] = f'${product_prices[selected_sheet['B24'].value]:,.2f}'
        order_data['Quantity0'] = selected_sheet['C24'].value
        amount = product_prices[selected_sheet['B24'].value] * selected_sheet['C24'].value
        order_data['Amount0'] = f'${amount:,.2f}'
        order_data['Description0'] = product_descriptions[selected_sheet['B24'].value]
        total += amount

    if selected_sheet['B25'].value is not None:
        order_data['Price1'] = f'${product_prices[selected_sheet['B25'].value]:,.2f}'
        order_data['Quantity1'] = selected_sheet['C25'].value
        amount = product_prices[selected_sheet['B25'].value] * selected_sheet['C25'].value
        order_data['Amount1'] = f'${amount:,.2f}'
        order_data['Description1'] = product_descriptions[selected_sheet['B25'].value]
        total += amount

    if selected_sheet['B26'].value is not None:
        order_data['Price2'] = f'${product_prices[selected_sheet['B26'].value]:,.2f}'
        order_data['Quantity2'] = selected_sheet['C26'].value
        amount = product_prices[selected_sheet['B26'].value] * selected_sheet['C26'].value
        order_data['Amount2'] = f'${amount:,.2f}'
        order_data['Description2'] = product_descriptions[selected_sheet['B26'].value]
        total += amount

    order_data['Total'] = f'{total:,.2f}'

    return order_data, name, email, date


def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', f'{directory}Invoice {invoice_num}.pdf', order_data)

create_invoice(8515, get_order_data(8515, 'Order1.xlsx', directory)[0], directory)


gmail_password_path = select_file('Please select the gmail password file')
gmail_password = open(gmail_password_path).read()

def send_email(invoice_num, name, email, date, directory):
    # Enter your gmail user name and password

    gmail_user = "josephrather7@gmail.com"

    # List of recipients

    recipients = [email]

    # Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

    # Body of the email

    body = f"""Dear {name},
        
Attached is your invoice dated {date}.

Thank you for your business. 

PythonTech Co.
10 Python Drive
Pleasant Grove, UT 84062
Phone: (801) 555-2372
Email: sales@pythontechco.com"""

    # Initialize the email message

    message = EmailMessage()
    message['From'] = gmail_user
    message['To'] = recipients
    message['Subject'] = subject
    message.set_content(body)

    # Add attachment

    path_to_file = f"{directory}Invoice {invoice_num}.pdf"
    message = add_attachment(path_to_file, message)
            
    # Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

    # Send the email

    mail_server.send_message(message)
    mail_server.quit()


filenames = os.listdir(directory)
invoice_num = 8515

for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        data = get_order_data(invoice_num, filename, directory)
        create_invoice(invoice_num, data[0], directory)
        send_email(invoice_num, data[1], data[2], data[3], directory)

        invoice_num += 1


