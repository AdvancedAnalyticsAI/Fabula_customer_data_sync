# Install required packages
!pip install pymssql gspread gspread-dataframe oauth2client

# Import libraries
import pandas as pd
import pymssql
import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials

# Database connection
conn = pymssql.connect(
    server='172.16.20.2',
    user='sa',
    password='Admin1',
    database='FabulaCoatings'
)

# Query: Last 12 months of customer-product purchase history
query = """
SELECT 
    FORMAT(i.InvDate, 'yyyy-MM') AS PurchaseMonth,
    c.Name AS CustomerName,
    s.StockLink,
    s.Description_1 AS ProductDescription,
    SUM(il.fQuantity) AS TotalQuantityPurchased,
    SUM(il.fQuantity * il.fUnitPriceExcl) AS TotalAmountSpent
FROM 
    InvNum i
JOIN 
    _btblInvoiceLines il ON i.AutoIndex = il.iInvoiceID
JOIN 
    Client c ON i.AccountID = c.DCLink
JOIN 
    StkItem s ON il.iStockCodeID = s.StockLink
WHERE 
    i.InvDate >= DATEADD(MONTH, -12, GETDATE())
GROUP BY 
    FORMAT(i.InvDate, 'yyyy-MM'),
    c.Name,
    s.StockLink,
    s.Description_1
ORDER BY 
    PurchaseMonth DESC,
    c.Name,
    TotalAmountSpent DESC;
"""

# Run query
df = pd.read_sql(query, conn)
conn.close()

# Google Sheets API setup
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
client = gspread.authorize(creds)

# Open the target spreadsheet
spreadsheet = client.open('Fabula Customer Purchases')
worksheet = spreadsheet.worksheet('12_Month_Sales')
worksheet.clear()

# Upload data
set_with_dataframe(worksheet, df)

print("✅ Exported to Google Sheets successfully.")