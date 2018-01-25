import shutil
import os
from os import path
from win32com.client.gencache import EnsureDispatch
from win32com.client import constants


for root, dirs, files in os.walk("C:\\Users\\Ozgur\\Desktop\\htm"):
 
 for dir in dirs:
   try:
    print(dir)
    shutil.rmtree("C:\\Users\\Ozgur\\Desktop\\htm\\"+dir)  
     #os.unlink("C:\\Users\\Ozgur\\Desktop\\htm\\"+file)
   except:
    continue




for root, dirs, files in os.walk("C:\\Users\\Ozgur\\Desktop\\htm"):
 for file in files:
  if file.endswith('.htm'):
   try:
    os.remove("C:\\Users\\Ozgur\\Desktop\\htm\\"+file)
   except:
    continue   



 
for root, dirs, files in os.walk("C:\\Users\\Ozgur\\Desktop\\htm"):
 
 for file in files:
   if file.endswith('.xlsm'):
    try:
     pathy=os.path.abspath(file)
     print(file)
     print (pathy)
    
     print(pathy)    
     #os.unlink("C:\\Users\\Ozgur\\Desktop\\htm\\"+file)
    
     xl = EnsureDispatch('Excel.Application')
     wb = xl.Workbooks.Open("C:\\Users\\Ozgur\\Desktop\\htm\\"+file)
     wb.RefreshAll()
     wb.Save()      
     file=file.replace("xlsm","htm")
     new= "C:\\Users\\Ozgur\\Desktop\\htm\\"+file
     print(new) 
    
     wb.SaveAs(new, constants.xlHtml)
     xl.Workbooks.Close()
     
    except:
     continue 
     
     xl.Quit()
     del xl     
