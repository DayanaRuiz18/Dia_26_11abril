#Importar panda, openpyxl y numpy

import pandas
import openpyxl
import numpy

#Abrir archivo1

m = openpyxl.load_workbook('Archivo1.xlsx',data_only = True)

#Activar la hoja del archivo1 en la que se encuentran los datos

hoja1 = m.active

#Acceder a las celdas que continen la información a utilizar

celdas = hoja1['A1':'C3']

#Recorrer y obtener el contenido de las celdas para añadirlo a una lista

matrizm = []

for fila in celdas:
  a = [celda.value for celda in fila] 
  matrizm.append(a)

#print(matrizm)

#Cerrar archivo 1

m.close()

#Convertir la lista obtenida en un dataframe

df1 = pandas.DataFrame(matrizm)

#Añadir los nombres de las celdas en la que se encuentran los datos

df12 = pandas.DataFrame(matrizm,index=[1,2,3], columns=["A","B","C"])
#print(str(df12))

#Abrir archivo2

n = openpyxl.load_workbook('Archivo2.xlsx',data_only = True)

#Activar la hoja del archivo1 en la que se encuentran los datos

hoja2 = n.active

#Acceder a las celdas que continen la información a utilizar

celdas2 = hoja2['A1':'C3']

#Recorrer y obtener el contenido de las celdas para añadirlo a una lista

matrizn = []

for fila in celdas2:
  b = [celda.value for celda in fila] 
  matrizn.append(b)

#print(matrizn)

#Cerrar archivo 2

n.close()

#Convertir la lista obtenida en un dataframe

df2 = pandas.DataFrame(matrizn)

#print(str(df2))

#Añadir los nombres de las celdas en la que se encuentran los datos

df22 = pandas.DataFrame(matrizn,index=[1,2,3], columns=["A","B","C"])

#Crear el archivo 3 en donde se van a guardar las respuestas del archivo 1

e = open('Archivo3.xlsx','w+')

#Crear el archivo 4 en donde se van a guardar las respuestas del archivo 2

f = open('Archivo4.xlsx','w+')

##Trabajo en el archivo 3

#Escribir el contenido del archivo 1 en el archivo 3

e.write("El contenido del archivo 1 es:\n")
e.write(str(matrizm))

#Escribir como un dataframe el contenido del archivo 1 en el archivo 3 

e.write("\n\nEl dataframe con el contenido del archivo 1 es:\n")
e.write(str(df1))

#Escribir el dataframe con indices del contenido del archivo 1 en el archivo 3 

e.write("\n\nEl dataframe con el contenido del archivo 1 con sus respectivos nombres de celdas es:\n")
e.write(str(df12))

#Cambiar dos datos del dataframe1

df12.iloc[1,1] = -2
df12.iloc[2,1] = 2

#Escribir el nuevo dataframe en el archivo 3

e.write("\n\nSe realizaron unos cambios en los valores de algunas celdas, el nuevo dataframe es:\n")
e.write(str(df12))

#Cerrar y guardar la infromación del archivo 3

e.close()

##Trabajo en el archivo 4

#Escribir el contenido del archivo 2 en el archivo 4

f.write("El contenido del archivo 2 es:\n")
f.write(str(matrizn))

#Escribir como un dataframe el contenido del archivo 2 en el archivo 4 

f.write("\n\nEl dataframe con el contenido del archivo 2 es:\n")
f.write(str(df2))

#Escribir el dataframe con indices del contenido del archivo 2 en el archivo 4 

f.write("\n\nEl dataframe con el contenido del archivo 2 con sus respectivos nombres de celdas es:\n")
f.write(str(df22))

#Crear una nueva matriz

o = numpy.array([[4,7,1],[2,5,1],[1,9,4]])

#Crear un nuevo dataframe

a = pandas.DataFrame(o,index=[4,5,6], columns=["A","B","C"]) 

#Escribir el dataframe a en el archivo 4

f.write("\n\nSe creó un dataframe nueva llamada a, la cual es la siguiente:\n")
f.write(str(a))

#Sumar el dataframe del archivo 2 con el dataframe a

d = pandas.concat([df22,a])
#print(d)

#Escribir la suma anterior en el archivo 4

f.write("\n\nSi sumamos el dataframe del archivo 2 con el nuuevo dataframe a obtenemos el siguiente dataframe:\n")
f.write(str(d))

#Cerrar y guardar la infromación del archivo 4

f.close()