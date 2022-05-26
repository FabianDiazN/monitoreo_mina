import serial
import xlwt
from datetime import datetime

class Guardar_datos:

    def __init__(self,port,speed):
        
        self.port = port
        self.speed = speed

        self.wb = xlwt.Workbook()
        self.ws = self.wb.add_sheet("Datos de arduino",cell_overwrite_ok=True)
        self.ws.write(0, 0, "Datos leidos desde arduino y python")
        self.columns = ["Fecha y hora"]
        self.number = 100
        

    def setColumns(self,col):
        self.columns.extend(col)
        
    def setRecordsNumber(self,number):
        self.number = number
        
    def readPort(self):
        ser = serial.Serial(self.port, self.speed, timeout=1)
        c = 0
        for col in self.columns:
            self.ws.write(1, c, col)
            c = c + 1
        self.fila = 2

        i = 0
        while(i<self.number):
            line = str(ser.readline())
            if(len(line) > 0):
                now = datetime.now()
                date_time = now.strftime("%m/%d/%Y, %H:%M:%S")
                print(date_time,line)
                if(line.find(",")):
                    c = 1
                    self.ws.write(self.fila, 0, date_time)
                    columnas = line.split(",")
                    for col in columnas:
                        self.ws.write(self.fila, c, col)
                        c = c + 1

                i = i + 1
                self.fila = self.fila + 1
    
    def writeFile(self,archivo):
    	self.wb.save(archivo)

    