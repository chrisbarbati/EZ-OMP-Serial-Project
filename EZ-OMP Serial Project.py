#Christian Barbati
#Serial Logger for EZ-OMP
#Capture serial data in real time & create an Excel Spreadsheet

#Goal is to create a serial logger that will parse serial data from COM port directly into .xlsx file. 
#Will start with basic spreadsheet functionality, then move to parsing existing text file captured from serial, THEN try to add serial read functionality.

#Imports
import openpyxl #Documentation: https://openpyxl.readthedocs.io/en/stable/
import os

#Ask user for the name of the spreadsheet and store it.
print("Please input the name of the spreadsheet: ")
spreadsheetName = input()

#Instantiate a new workbook object, and select the sheet
workbook1 = openpyxl.Workbook()
worksheet = workbook1.active

#Open the text file
sourceText = open("CAPTURE.TXT", "r")

i = 1
endOfFile = False

#Iterate over each line to print the contents, ending when an empty line is returned
while(not endOfFile):
    sourceContent = sourceText.readline(i)
    print(sourceContent)
    i += 1

    if(sourceContent == ""):
        endOfFile = True

#Add headings:
worksheet["A1"] = "Time (ms):"
worksheet["B1"] = "Target:"
worksheet["C1"] = "Valve:"
worksheet["D1"] = "Difference:"

#Save the spreadsheet to a file, named per the input collected earlier
workbook1.save(spreadsheetName+".xlsx")