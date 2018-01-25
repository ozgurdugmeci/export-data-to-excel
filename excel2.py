import pandas as pd
import sqlite3
from openpyxl import load_workbook


connection = sqlite3.connect("C:\\sqlite\\testo.db")
cursor = connection.cursor()
cursor.execute("SELECT cinsiyet, isim FROM tablo_customer where id between 8 and 500 ")
report=cursor.fetchall()


df= pd.DataFrame(report)
df.columns=['cinsiyet','isim']
#print (df)
file = "C:\\Users\\Ozgur\\Desktop\\htm\\graph2.xlsx"

path=file
book = load_workbook(path)
writer = pd.ExcelWriter(path, engine = 'openpyxl')
writer.book = book
writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
df.to_excel(writer, "ss")

writer.save()
writer.close()
