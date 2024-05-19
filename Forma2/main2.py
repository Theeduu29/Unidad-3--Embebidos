import serial
import xlsxwriter

com_port = "COM5"  #puerto usb arduino conectado
baud_rate = 9600  #el beguin que esta funcionando

# Abre el puerto serial
ser = serial.Serial(com_port, baud_rate, timeout=1)  #abre el serial del arduino
ser.flushInput()  #limpia los datos al final

#crea los archivos exel para guardar los datos
workbook1 = xlsxwriter.Workbook('Promedio.xlsx')
worksheet1 = workbook1.add_worksheet()

workbook2 = xlsxwriter.Workbook('Mediana.xlsx')
worksheet2 = workbook2.add_worksheet()

workbook3 = xlsxwriter.Workbook('ValorMayor.xlsx')
worksheet3 = workbook3.add_worksheet()

workbook4 = xlsxwriter.Workbook('ValorMenor.xlsx')
worksheet4 = workbook4.add_worksheet()

# Escribe los encabezados en cada archivo Excel
#A1(numero de casilla donde se coloca el texto) texto
#worksheet1.write('A1', 'Promedio')

#worksheet2.write('A1', 'Mediana')

#worksheet3.write('A1', 'ValorMayor')

#worksheet4.write('A1', 'ValorMenor')


row = 1  # Inicializa el contador de filas para cada archivo

while True:
    try:
        data = ser.readline().decode('utf-8').rstrip()  #lee y decodifca los datos del arduino
        if data:
            values = data.split(',')  # se utiliza para dividir la linea de datos como espaciador
            worksheet1.write(row, 0, values[0])  #row columna x y values valor especifico
            print(values[0])

            worksheet2.write(row, 0, values[1])
            print(values[1])

            worksheet3.write(row, 1, values[2])
            print(values[2])

            worksheet4.write(row, 1, values[3])
            print(values[3])

            row += 1  # Incrementa el contador de filas para el siguiente valor

    #viene siendo un sino
    except KeyboardInterrupt:
        print("ha ocurrido un fallo")
        break

#creacion de los archivos xlsx exel
workbook1.close()
workbook2.close()
workbook3.close()
workbook4.close()

#cierra el serial y la conexion
ser.close()
