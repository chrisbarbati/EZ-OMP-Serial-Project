#Christian Barbati
#Serial Logger for EZ-OMP
#Capture serial data in real time & create an Excel Spreadsheet

#Goal is to create a serial logger that will parse serial data from COM port directly into .xlsx file. 
#Will start with basic spreadsheet functionality, then move to parsing existing text file captured from serial, THEN try to add serial read functionality.

#Need to add: Filter out nan values, outliers, and blanks, then remove all blank cells and calculate difference.

#Imports
import openpyxl #Documentation: https://openpyxl.readthedocs.io/en/stable/
import os
import serial

#Ask user for desired COM port (temporarily set as 5)
#print("What COM port # should be read from?")
#comPort = input()
comPort = 5

#Open a serial port, baud rate 9600
ser1 = serial.Serial('COM' + str(comPort), 9600)

#Simple test of readline() function. Remove me later4
while(True):
    data = str(ser1.readline()) #Arduino is configured to add a carriage return at end of each line
    data = data.lstrip("b'")
    data = data.rstrip("'")
    data = data.rstrip("\\r\\n") #Strip out the b'' from the serial library, and escape the escape characters to remove \r and \n
    print(data)

#Ask user for the name of the spreadsheet and store it.
#print("Please input the name of the spreadsheet: ")
spreadsheetName = "Test"#input()

#Instantiate a new workbook object, and select the sheet
workbook1 = openpyxl.Workbook()
worksheet = workbook1.active

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

    sourceContent = sourceContent.rstrip("\n")

    #Check the first character of the content, and send it to the appropriate cell
    if(not sourceContent[0] == "T") and (not sourceContent[0] == "F"):
        worksheet["A" + str(j + 1)] = sourceContent
    if(sourceContent[0] == "T"):
        worksheet["B" + str(j)] = sourceContent[2:] #Strips the T: from in front of the target valve position
    if(sourceContent[0] == "V"):
        worksheet["C" + str(j)] = sourceContent[2:] #Strips the V: from the measured valve position

    i += 1

    #For every three rows of data from the txt file, we need to write one row of the spreadsheet, so j increases at 1/3 the rate of i
    if(i % 3 == 0):
        j += 1

#First and last few rows sometimes display errant data, so delete them. 30 rows per second, so -25 rows is only losing about 1 second of data.
worksheet.delete_rows(2, 10)
worksheet.delete_rows(j-15, j)

#Iterate over the spreadsheet to remove NaN, outliers. Then remove all rows with a blank value, shift rows up to fill empty space, and calculate difference.
columns = ["B", "C"]

for column in columns:
    for row in range (2,j):

        #Replace all nan values and outliers (less than 5%, greater than 95%)
        if(worksheet[str(column) + str(row)].value == None):
            worksheet[str(column) + str(row)] = ""
        if(worksheet[str(column) + str(row)].value == "nan"):
            worksheet[str(column) + str(row)] = ""

#Iterate over the spreadsheet and convert all strings to floats:
columns = ["A", "B", "C"]

for column in columns:
    for row in range (2,j):
        #Only attempt to convert to float IF the value is a string type, and if it is not blank
        if((type(worksheet[str(column) + str(row)].value) == str) and not (worksheet[str(column) + str(row)].value == "")):
            worksheet[str(column) + str(row)] = float(worksheet[str(column) + str(row)].value)
        

#Save the spreadsheet to a file, named per the input collected earlier
workbook1.save(spreadsheetName+".xlsx")