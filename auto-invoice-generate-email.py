import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font
import win32com.client as win32
import os
from datetime import date
import time
import shutil
today = str(date.today())
dirc = "C:\\Users\\logan\\OneDrive\\Documents\\Work\\SGS\\Automate Invoiceing\\"

invoice_data = pd.read_excel(f"{dirc}\\invoice_data.xlsx")


#Make a set of uniqe companies matched with their email
comp_data = []
for index, row in invoice_data.iterrows():
    comp_data.append((row.Company,row.Email))
comp_data = list(set(comp_data))


# Create folder "to send"
if not os.path.exists(f"{dirc}\\to_send"):
        os.makedirs(f"{dirc}\\to_send")


#Create invoices by using uniqe company list, finding all pairs and putting that data onto inovicing template
for i in range(len(comp_data)):
    temp_data = []
    temp_comp = comp_data[i][0]


    for _,row in invoice_data.iterrows():
        # add values in a tupple in the format (invoice no, date, outstanding balance)
        if temp_comp == row.Company:
            temp_data.append((row.Invoice_No,str(row.Date)[:10], f"${row.Outstanding_balance}"))
            
    #sort the data based on date    
    temp_data = sorted(temp_data, key=lambda x: x[1])   
    
    #Open and activate the excel
    invoice_template = load_workbook(f"{dirc}\\invoice_template.xlsx")
    sheet = invoice_template.active  # Assuming you are working with the active sheet
    
    roll_bal = 0 #sum of outstanding balance for the company
    for i, val in enumerate(temp_data):
        sheet[f'B{7+i}'] = val[0]  # Invoice No.
        sheet[f'D{7+i}'] = val[1]  # Date
        sheet[f'F{7+i}'] = val[2]  # Outstanding balance
        roll_bal += int(val[2][1:])
        
    sheet[f"D{7+i+2}"] = "Total Balance"
    sheet[f"D{7+i+2}"].font = Font(bold=True)
    
    sheet[f"F{7+i+2}"] = f'${roll_bal}'
    sheet[f"F{7+i+2}"].font = Font(bold=True)
    
    sheet[f"D{7+i+3}"] = "Date"
    sheet[f"D{7+i+3}"].font = Font(bold=True)
    
    sheet[f"F{7+i+3}"] = today
    sheet[f"F{7+i+3}"].font = Font(bold=True)
    
#O(n^2)
    #Save file
    invoice_template.save(f"{dirc}\\to_send\\{temp_comp}_{today}.xlsx")

    
#Create sent foldder
if not os.path.exists(f"{dirc}\\sent"):
    os.makedirs(f"{dirc}\\sent")

#Itterate through the folder we made, and send of each excel
for i, filename in enumerate(os.listdir(f"{dirc}\\to_send")):
    file_path = os.path.join(f"{dirc}\\to_send", filename)
    
    
    temp_comp = comp_data[i][0]
    temp_email = comp_data[i][1]
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0) #0 just mean basic email type 
    mail.Subject = f'SGS Outstanding invoices for {temp_comp}'
    mail.To = temp_email
    mail.Body = f"""Dear {temp_comp},

Our records show that you have at least one outstanding invoice. Please see attached for information. If you have any questions or concerns, please let us know.

Thank you
Logan
"""
    mail.Attachments.Add(file_path)


#     mail.Send()
    shutil.move(file_path,f"{dirc}\\sent\\")
