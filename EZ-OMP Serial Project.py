#Christian Barbati
#Serial Logger for EZ-OMP
#Capture serial data in real time & create an Excel Spreadsheet

#Goal is to create a serial logger that will parse serial data from COM port directly into .xlsx file. 
#Will start with basic spreadsheet functionality, then move to parsing existing text file captured from serial, THEN try to add serial read functionality.

#Imports
import openpyxl #Documentation: https://openpyxl.readthedocs.io/en/stable/
import os
import serial
from pytimedinput import timedKey #Allows a timeout on the input function, which I will use to interrupt the loop.
import pyinputplus as pyip #To validate the user input filename and COM port #
import re

print("What COM port # should be read from?")
comPort = pyip.inputInt()

#Open a serial port, baud rate 9600
ser1 = serial.Serial('COM' + str(comPort), 9600)

#Variable to hold whether user wants to stop reading (default False)
stopReading = False

#i counts the current iteration of the loop, j counts every third iteration to move to the next row
i = 1
j = 1

#Regex & custom input function to validate that spreadsheet name does not contain non-alphanumeric characters
sheetNameCheck = re.compile("\w")

def inputCheck(spreadsheetName):
    if(not sheetNameCheck.search(spreadsheetName)):
        raise Exception("Invalid input. Alphanumeric values only.")
    return spreadsheetName

#Ask user for the name of the spreadsheet and store it.
print("Please input the name of the spreadsheet: ")
spreadsheetName = pyip.inputCustom(inputCheck)

#Instantiate a new workbook object, and select the sheet
workbook1 = openpyxl.Workbook()
worksheet = workbook1.active

#Add headings:
worksheet["A1"] = "Time (ms):"
worksheet["B1"] = "Target:"
worksheet["C1"] = "Valve:"
worksheet["D1"] = "Difference:"

#Simple test of readline() function. Remove me later4
while(not stopReading):

    data = str(ser1.readline()) #Arduino is configured to add a carriage return at end of each line

    data = data.lstrip("b'")
    data = data.rstrip("'")
    data = data.rstrip("\\r\\n") #Strip out the b'' from the serial library, and escape the escape characters to remove \r and \n

    print(data)

    #Check the first character of the content, and send it to the appropriate cell
    if(not data[0] == "T") and (not data[0] == "V"):
        worksheet["A" + str(j + 1)] = float(data)
    if(data[0] == "T"):
        data = data.lstrip("T:")
        worksheet["B" + str(j + 1)] = float(data)
    if(data[0] == "V"):
        data = data.lstrip("V:")
        worksheet["C" + str(j)] = float(data)

    i += 1

    if(i % 3 == 0):
        j += 1

    userKeystroke, timedOut = timedKey(prompt="", timeout=.001) #Set timeout at .001 so execution isn't blocked

    if(timedOut):
        continue
    elif(not userKeystroke == ""): #If the user strikes any key, userKeystroke will not be empty, and the loop will break
        stopReading = True
    else:
        continue

worksheet.delete_rows(j-2, j)

#Iterate over the sheet and calculate the difference between commanded valve position and actual valve position

for row in range (2, j-2):
    target = float(worksheet["B" + str(row)].value)
    actual = float(worksheet["C" + str(row)].value)

    difference = abs(target-actual) #Sign is not important, only magnitude, so we'll take the absolute value

    worksheet["D" + str(row)] = difference


#Save the spreadsheet to a file, named per the input collected earlier
workbook1.save(spreadsheetName+".xlsx")