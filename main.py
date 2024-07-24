import os
import win32com.client as client

excel = client.Dispatch("excel.application")
for file in os.listdir(os.getcwd()+"/oldVersion/"):
    filename,fileextension = os.path.splitext(file)
    wb = excel.Workbooks.Open(os.getcwd() + "/oldVersion/" + file)
    output_path  = os.getcwd()+ "/newVersion/" + filename
    wb.SaveAs(output_path,51)
    wb.Close()
excel.Quit()
