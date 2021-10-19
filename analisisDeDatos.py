import pandas as pd
import os 
import matplotlib.pyplot as plt
import numpy as np
RUTA = "C:/Ajax,json/Recursos/613_Tabla_Estadistica_Motos_2019.xlsx"#Ruta de excel o nombre
def main():
    os.system ("cls")
    PASSWORD = addUser()
    
    #Se limpia la consola para mejor legibilidad
    os.system ("cls")
    print("Welcome user")
    #Ingresamos nuestro contraseña
    passw = str(input("Password: "))
    option = 1
    #Validación de contraseña sea igual a la integrada anteriormente
    if(passw == PASSWORD):
        #Menu de opciones
        os.system("cls")
        print("         Menu")
        while(option != 99):
            print("Options")
            print("1.-Read Excel document")
            print("2.-Analyze excel document")
            print("3.-Graficar")
            print("99.-Exit")
            option = int(input("Option >> "))
            if(option == 1):
                print(readExcel())
            elif(option == 2):
                print("Analyze excel document")
                print("¿Cual fue la venta con mas utilidad?")
                mayorU()
                print("\n\n\n¿Cual fue la venta con menos utilidad?")
                menorU()
                graficar()
            elif(option == 3):
                print("GRaficar completa")
                graficar()
            elif(option==99):
                print("Thanks")
                break
            else:
                print("Invalid option")

#Funciones del programa
#Funcion para crear password
def addUser():
    password = str(input("Your password is: "))
    return password

#Funcion leer el documento excel completo
def readExcel():
    df = pd.read_excel(RUTA, sheet_name="consulta")
    return df

#Funcion que retorna la utilidad mas grande generada
def mayorU():
    df = pd.read_excel(RUTA)
    valores=df[["Utilidad"]]
    arr = np.array([valores])
    x = np.amax(arr)
    print(f"Utilidad = {x}")
    for agente in df.values:
        if(x in agente):
            print(agente)
    return x
#Funcion que retorna la utilidad menor generada
def menorU():
    df = pd.read_excel(RUTA)
    valores=df[["Utilidad"]]
    arr = np.array([valores])
    x = np.amin(arr)
    print(f"Utilidad = {x}")
    for agente in df.values:
        if(x in agente):
            print(agente)
            
    
#Funcion que da la grafica del analisis del excel
def graficar():
    df = pd.read_excel(RUTA)
    valores=df[["Agente","Utilidad"]]
    ax = valores.plot.bar(x="Agente",y="Utilidad", rot=0)
    plt.show()

if __name__ == '__main__':
    try:
        main()
    except:
        exit()