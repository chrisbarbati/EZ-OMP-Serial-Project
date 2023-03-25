#Christian Barbati
#Serial Logger for EZ-OMP
#Capture serial data in real time & create an Excel Spreadsheet

#Goal is to create a serial logger that will parse serial data from COM port directly into .xlsx file. 
#Will start with basic spreadsheet functionality, then move to parsing existing text file captured from serial, THEN try to add serial read functionality.

#Imports
import openpyxl #Documentation: https://openpyxl.readthedocs.io/en/stable/
import os #Documentation: 

#Ask user for the name of the spreadsheet and store it.
print("Please input the name of the spreadsheet: ")
spreadsheetName = input()

#Instantiate a new workbook object, and select the sheet
workbook1 = openpyxl.Workbook()
worksheet = workbook1.active

#Open the text file and read the contents
sourceText = open("CAPTURE.TXT")
sourceContent = sourceText.read()

#Split the contents of the text file into a list of lines.
sourceContent.split()

#Iterate over each line to print the contents
for line in sourceContent:
    print(sourceContent)

#Add headings:
worksheet["A1"] = "Time (ms):"
worksheet["B1"] = "Target:"
worksheet["C1"] = "Valve:"
worksheet["D1"] = "Difference:"

#Save the spreadsheet to a file, named per the input collected earlier
workbook1.save(spreadsheetName+".xlsx")