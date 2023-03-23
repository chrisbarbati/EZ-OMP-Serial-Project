#Christian Barbati
#Serial Logger for EZ-OMP
#Capture serial data in real time & create an Excel Spreadsheet

#Imports
import openpyxl #Documentation: https://openpyxl.readthedocs.io/en/stable/

#Ask user for the name of the spreadsheet and store it.
print("Please input the name of the spreadsheet: ")
spreadsheetName = input()

#Instantiate a new workbook object, and select the sheet
workbook1 = openpyxl.Workbook()
worksheet = workbook1.active

#Save the spreadsheet to a file, named per the input collected earlier
workbook1.save(spreadsheetName+".xlsx")