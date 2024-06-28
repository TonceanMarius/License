import serial  #import that enable the serial communication port
import time    #import that enable the time-related commands
import pyodbc  #import that enable the manipulation of the Microsoft Access database

arduino = serial.Serial(port='COM8', baudrate=115200, timeout=.1)  #establishe the connection between arduino and PC
time.sleep(1.5)   #delay


def write_read(x):    #This function will send the data over the serial port and verify if the data is correctly sent
    arduino.write(bytes(x, 'utf-8'))
    time.sleep(0.05)
    data = arduino.readline()
    return data

pyodbc.pooling=False  # disable the ODBC connection pooling

conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=E:\drive gabi\licenta\licenta 3\tabel.accdb;')  #define the batabase path
cursor = conn.cursor()  #set the cursor in the database
cursor.execute('select id, Availability from tracking_sales')  #extract data from the "Availability" field

fetchlist = {}     #define a new list

for row in cursor.fetchall():   #Extract data from database and store in the list
    fetchlist[row.id] = row.Availability

for i in sorted(fetchlist.keys()): #Sort the data from the list
    temp=write_read(fetchlist[i])  #Send the sorted data through the serial port
    print(fetchlist[i])            #Display the list elements in console

time.sleep(1)           #Delay

arduino.flushInput()    #Empty the input serial port buffer
arduino.flushOutput()   #Empty the output serial port buffer

time.sleep(0.5)         #Delay

cursor.close()          #delete the cursor from the database
arduino.close()         #close the serial port communication

