while True:
    directory = single_input("Please input the directory where the ProductData.xlsx, MasterInvoice.pdf, and Order xlsx files 1-4 are located.", mask=False, width=400, height=200)
    if directory.endswith('\\'):
        pass
    else:
        directory= directory+'\\'
    filenames= os.listdir(directory)
    if 'ProductData.xlsx' in filenames and 'MasterInvoice.pdf' in filenames and 'Order1.xlsx' in filenames and 'Order2.xlsx' in filenames and 'Order3.xlsx' in filenames and 'Order4.xlsx' in filenames:
        print(directory)
        break 
    else: 
        show_message('Error', 'Directory Does Not Contain Required Files. Please Re-Enter Directory.')

product_descriptions={}
descr_df = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Descriptions')
for index,row in descr_df.iterrows():
    product_descriptions[row['Unit number']]=row['Description']

product_prices={}
price_df = pd.read_excel(directory+'ProductData.xlsx', sheet_name='Prices')
for index,row in price_df.iterrows():
    product_prices[row['Unit number']]=row['Price per unit']

product_descriptions

def get_order_data(invoice_num,filename,directory):
    order_data={}
    total= 0
    template = openpyxl.load_workbook(directory+filename, data_only=True)
    selected_sheet = template['OrderForm']
    order_date = selected_sheet['C9'].value
    month= order_date.month
    day= order_date.day
    year=order_date.year
    date= (f'{month}/{day}/{year}')
    email= selected_sheet['C13'].value
    name= selected_sheet['C14'].value
    company_name= selected_sheet['C15'].value
    street= selected_sheet['C16'].value
    city= selected_sheet['C17'].value
    state= selected_sheet['C18'].value
    zipcode= selected_sheet['C19'].value

    order_data['InvoiceNum']= invoice_num
    order_data['Date']= date
    order_data['CustomerAddress1']= street
    order_data['CustomerName']= name
    order_data['CustomerCompany']= company_name
    order_data['CustomerAddress2']= f'{city}, {state}, {zipcode}'

    unit_number0= selected_sheet['B24'].value
    quantity0= selected_sheet['C24'].value
    product_descr0= product_descriptions[unit_number0]
    product_price0= product_prices[unit_number0]
    product_total0= quantity0*product_price0
    price0formatted= f'${product_price0:,.2f}'
    product0_totalformatted= f'${product_total0:,.2f}'
    total= total + product_total0

    order_data['Price0']= price0formatted
    order_data['Quantity0']= quantity0
    order_data['Description0']= product_descr0
    order_data['Amount0']= product0_totalformatted
    

    unit_number1= selected_sheet['B25'].value
    if unit_number1 is not None and unit_number1 != '':
        quantity1= selected_sheet['C25'].value
        product_descr1= product_descriptions[unit_number1]
        product_price1= product_prices[unit_number1]
        product_total1= quantity1*product_price1
        price1formatted= f'${product_price1:,.2f}'
        product1_totalformatted= f'${product_total1:,.2f}'
        total= total+product_total1

        order_data['Price1']= price1formatted
        order_data['Quantity1']= quantity1
        order_data['Amount1']= product1_totalformatted
        order_data['Description1']= product_descr1
    

    
    unit_number2= selected_sheet['B26'].value
    if unit_number2 is not None and unit_number2 != '':
        quantity2= selected_sheet['C26'].value
        product_descr2= product_descriptions[unit_number2]
        product_price2= product_prices[unit_number2]
        product_total2= quantity2*product_price2
        price2formatted= f'${product_price2:,.2f}'
        product2_totalformatted= f'${product_total2:,.2f}'
        total= total + product_total2

        order_data['Price2']= price2formatted
        order_data['Quantity2']= quantity2
        order_data['Amount2']= product2_totalformatted
        order_data['Description2']= product_descr2

    total_formatted= f'${total:,.2f}'
    order_data['Total']= total_formatted
        

    return order_data, name, email, date

invoice_num= 8515
filename= 'Order1.xlsx'
order_data, name, email, date= get_order_data(invoice_num,filename,directory)
print(order_data)
print(name)
print(email)
print(date)

def create_invoice(invoice_num, order_data, directory):
    create_pdf(directory + 'MasterInvoice.pdf', directory + f'Invoice {invoice_num}.pdf', order_data)

invoice_num= 8515

create_invoice(invoice_num, order_data, directory)

def send_email(invoice_num, name, email, date, directory):
    gmail_user = 'emmagstone7@gmail.com'
    gmail_password = open('gmail_password.txt').read()

# List of recipients

    recipients = email

# Email subect

    subject = f'Your Order from PythonTech Co. - Invoice {invoice_num}'

# Body of the email

    body = f"""Dear {name},
    
    Attached is your invoice dated {date}.

    Thank you for your business.

    PythonTech Co. 
    10 Python Dr. 
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

    path_to_file = directory+f'Invoice {invoice_num}.pdf'
#path_to_file = os.getcwd()+'/Invoice Template.xlsx' # MAC
    message = add_attachment(path_to_file, message) # This function is available from the analyticskit package
        
# Log in to the Gmail server

    mail_server = smtplib.SMTP_SSL('smtp.gmail.com')
    mail_server.login(gmail_user, gmail_password)

# Send the email

    mail_server.send_message(message)
    mail_server.quit()


send_email(invoice_num, name, email, date, directory)

invoice_num= 8515
filenames= os.listdir(directory)
for filename in filenames:
    if 'Order' in filename and 'xlsx' in filename:
        order_data, name, email, date= get_order_data(invoice_num,filename,directory)
        #have to use file insead of filename so it uses each new file each time it runs
        create_invoice(invoice_num, order_data, directory)
        send_email(invoice_num, name, email, date, directory)
        invoice_num= invoice_num+1