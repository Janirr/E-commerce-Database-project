import requests
import mysql.connector
import logging
import xlsxwriter

# All the warnings will go into app.log file
logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')


# Class Currency will have 3 letters code [EUR,USD,...] and with that we can get any rate from NBP
class Currency:
    def __init__(self, name, rate):
        self.name = name
        self.rate = rate

    def updateRate(self):
        URL = "http://api.nbp.pl/api/exchangerates/rates/a/" + self.name.lower() + "/"
        response = requests.get(URL)
        # If we have connection
        if response.status_code == 200:
            # If there exists a rate
            if response.json()['rates'][0]['mid']:
                rate = round(response.json()['rates'][0]['mid'], 2)
                print("{}/PLN exchange rate: {}".format(self.name, rate))
                # Change the rate to the rate from the URL
                self.rate = rate
            else:
                logging.warning("Problem with finding rate inside the url")
        else:
            logging.warning("Couldn't connect to:", URL)
            return "Couldn't connect to:", URL


# Values so they won't be highlighted
cursor = ''
ID_and_Price = []
db = ''

# Connecting to the database
try:
    db = mysql.connector.connect(
        host="localhost",
        user="root",
        password="1234",
        database="mydb"
    )
    cursor = db.cursor()
except mysql.connector.Error as e:
    logging.warning(e)
    exit()
try:
    cursor.execute("SELECT ProductID,UnitPrice FROM Product")
    ID_and_Price = cursor.fetchall()
except mysql.connector.Error as err:
    logging.warning(err)
    exit()

# Create USD and EUR Currency class
USD = Currency("USD", 0)
EUR = Currency("EUR", 0)
# Update their rates to their actual rates
USD.updateRate()
EUR.updateRate()
# count_rows will count how many records(s) were changed(affected)
# by updating their new USD/EUR rate

count_rows = 0
for x in ID_and_Price:
    # ProductID is the 1st value and its the best to separate them by
    # unique ID and compare them with WHERE from MySQL
    ProductID = x[0]
    try:
        cursor.execute("UPDATE Product SET UnitPriceUS = '{}' WHERE ProductID = '{}'".format(USD.rate, ProductID))
        cursor.execute("UPDATE Product SET UnitPriceEuro = '{}' WHERE ProductID = '{}'".format(EUR.rate, ProductID))
        # Commit it to the database
        db.commit()
        # cursor.rowcount will be either 0 or 1, count_rows will collect those values
    except mysql.connector.Error as err:
        logging.warning(err)
        exit()
    count_rows += cursor.rowcount
print(count_rows, "record(s) affected")


answer = int(input("Do you want to create Excel file with the list of all products? [y/n]"))
# Creating Product.xlsx file with the values from Produkt table
if answer == "y":
    workbook = xlsxwriter.Workbook('Product.xlsx')
    worksheet = workbook.add_worksheet()
    # Variables, col stands for number of specific column, same goes for row
    row = 0
    col = 0
    columns = []
    select_values = []
    # Get the headers [columns] from Product
    try:
        cursor.execute('SHOW columns FROM Product')
        columns = cursor.fetchall()
    except mysql.connector.Error as err:
        logging.warning(err)
        exit()
    for i in columns:
        # Write them into the first row in xls [A0,A1,A2,...]
        worksheet.write(row, col, i[0])
        col += 1
    col = 0
    row += 1
    try:
        cursor.execute("SELECT * FROM Product")
        select_values = cursor.fetchall()
    except mysql.connector.Error as err:
        logging.warning(err)
        exit()
    for i in select_values:
        for j in i:
            if type(j) is bytes:
                j = str(j)
            # Write them into the next row and columns in xls [B0,B1,B2,C0,C1,C2,...]
            worksheet.write(row, col, j)
            col += 1
        col = 0
        row += 1
    # Close the file
    workbook.close()
