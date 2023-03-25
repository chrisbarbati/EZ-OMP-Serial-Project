#Christian Barbati
#Serial Logger for EZ-OMP
#Capture serial data in real time & create an Excel Spreadsheet

#Goal is to create a serial logger that will parse serial data from COM port directly into .xlsx file. 
#Will start with basic spreadsheet functionality, then move to parsing existing text file captured from serial, THEN try to add serial read functionality.

#Need to add: Filter out nan values, outliers, and blanks, then remove all blank cells and calculate difference.

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

#i iterates over each line in the txt file, j iterates at 1/3 the speed for each line output in the spreadsheet
i = 1
j = 2
endOfFile = False

#Add headings:
worksheet["A1"] = "Time (ms):"
worksheet["B1"] = "Target:"
worksheet["C1"] = "Valve:"
worksheet["D1"] = "Difference:"

#Iterate over each line to print the contents, ending when an empty line is returned
while(not endOfFile):

    sourceContent = sourceText.readline(i)

    #Check if we have reached the end of the file, and if so change the boolean to true and re-evaluate
    if(sourceContent == ""):
        endOfFile = True
        continue

    print(sourceContent)

    #Check the first character of the content, and send it to the appropriate cell
    if(not sourceContent[0] == "T") and (not sourceContent[0] == "F"):
        worksheet["A" + str(j + 1)] = sourceContent
    if(sourceContent[0] == "T"):
        worksheet["B" + str(j)] = sourceContent
    if(sourceContent[0] == "V"):
        worksheet["C" + str(j)] = sourceContent

    i += 1

    #For every three rows of data from the txt file, we need to write one row of the spreadsheet, so j increases at 1/3 the rate of i
    if(i % 3 == 0):
        j += 1

#Save the spreadsheet to a file, named per the input collected earlier
workbook1.save(spreadsheetName+".xlsx")