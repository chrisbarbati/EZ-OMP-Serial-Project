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
from pytimedinput import timedKey #Allows a timeout on the input function, which I will use to interrupt the loop.

#Ask user for desired COM port (temporarily set as 5)
#print("What COM port # should be read from?")
#comPort = input()
comPort = 5

#Open a serial port, baud rate 9600
ser1 = serial.Serial('COM' + str(comPort), 9600)

#Variable to hold whether user wants to stop reading (default False)
stopReading = False

i = 1
j = 1

#Ask user for the name of the spreadsheet and store it.
#print("Please input the name of the spreadsheet: ")
spreadsheetName = "Test2"#input()

#Instantiate a new workbook object, and select the sheet
workbook1 = openpyxl.Workbook()
worksheet = workbook1.active

#Add headings:
worksheet["A1"] = "Time (ms):"
worksheet["B1"] = "Target:"
worksheet["C1"] = "Valve:"

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
        break
    else:
        continue



# #First and last few rows sometimes display errant data, so delete them. 30 rows per second, so -25 rows is only losing about 1 second of data.
# worksheet.delete_rows(2, 10)
# worksheet.delete_rows(j-15, j)

# #Iterate over the spreadsheet and convert all strings to floats:
# columns = ["A", "B", "C"]

# for column in columns:
#     for row in range (2,j):
#         #Only attempt to convert to float IF the value is a string type, and if it is not blank
#         if((type(worksheet[str(column) + str(row)].value) == str) and not (worksheet[str(column) + str(row)].value == "")):
#             worksheet[str(column) + str(row)] = float(worksheet[str(column) + str(row)].value)
        

#Save the spreadsheet to a file, named per the input collected earlier
workbook1.save(spreadsheetName+".xlsx")