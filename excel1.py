import time 
import sqlite3
from win32com.client.gencache import EnsureDispatch
from win32com.client import constants

connection = sqlite3.connect("C:\\sqlite\\testo.db")
cursor = connection.cursor()
cursor.execute("SELECT cinsiyet, isim FROM tablo_customer where id between 1 and 300 ")
report=cursor.fetchall()



file = "C:\\Users\\Ozgur\\Desktop\\htm\\graph.xlsm"
satr=2

xl = EnsureDispatch('Excel.Application')


wb2=xl.Workbooks.Open(file)
wb2.DisplayScreenAlerts=False
ws=wb2.Worksheets("data")

ws.Cells(1,1).Value= "Cinsiyet"
ws.Cells(1,2).Value= "isim"


ws.Range("A2:B20000").ClearContents()

for i in report:
 ws.Cells(satr,1).Value= i[0] 
 ws.Cells(satr,2).Value= i[1]
 satr=satr+1


wb2.RefreshAll()
time.sleep(2)

wb2.Save()
wb2.Close()
xl.Quit()
del xl
