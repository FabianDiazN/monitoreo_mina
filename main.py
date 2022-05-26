from datetime import date, datetime
import string
from excel import Guardar_datos


class estilo():

    AZUL = '\33[34m'
    ROJO = '\33[31m'
    VERDE = '\033[32m'


print(estilo.VERDE+"""                                                
+-----------------------------------------------------+
|  __  __                _   _                        |             
| |  \/  |  ___   _ _   (_) | |_   ___   _ _          |              
| | |\/| | / _ \ | ' \  | | |  _| / _ \ | '_|         |            
| |_|  |_| \___/ |_||_| |_|  \__| \___/ |_|    V.0.1  |
|                                                     |
+-----------------------------------------------------+                                                                                                                                                                             
""")
puerto = "COM4"
baudios = 9600
formato = ".xls"

nombre_archivo = input(
    estilo.AZUL+"Ingresa el nombre con el que se guardara el archivo:")

nombre_final = nombre_archivo+formato

numero_muestras = int(input(estilo.ROJO+"Numero de muestras:"))

serialToExcel = Guardar_datos(puerto, baudios)

columnas = ["Lectura"]

serialToExcel.setColumns(["Lectura"])
serialToExcel.setRecordsNumber(numero_muestras+2)
serialToExcel.readPort()
serialToExcel.writeFile(nombre_final)
print(estilo.AZUL+"_____________________________________________________________________________________________________")
print(estilo.VERDE+"El archivo:"+nombre_archivo+estilo.VERDE+" se ha guardado")
