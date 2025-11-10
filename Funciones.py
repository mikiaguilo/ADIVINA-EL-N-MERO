import random
import pandas as pd
import openpyxl
import os
from getpass import getpass
from openpyxl import load_workbook
from tabulate import tabulate   # si quieres formato bonito (opcional)

#Definiremos una función que nos cree un menú para nuestro juego.
def pedir_nombre_solitario():

    nombre = input("Hola! Cómo te llamas=")
    return nombre
#
def pedir_nombre_parejas():

    nombre_1 = input('¿Cómo se llama el primer integrante?')
    nombre_2 = input('¿Cómo se llama el segundo integrante?')
    return nombre_1,nombre_2
#Función para pedir el nivel del jugador en modo solitario
def pedir_nivel_solitario(nombre):

    opcion_valida=False
    print(f'Hola {nombre}, ha llegado la hora de decidir la dificultad del juego.')
    while opcion_valida == False:
        print("\n--- Nivel de dificultad ---")
        print("1. Pulsa 1 si quieres jugar el modo fácil y tener 20 intentos.")
        print("2. Pulsa 2 si quieres jugar el modo intermedio y tener 12 intentos.")
        print("3. Pulsa 3 si quieres jugar el modo difícil y tener 5 intentos.")

        opcion = int(input("Elige una opción: "))

        if opcion == 1:
            dificultad = 1
            numero_intentos = 20
            print(f'Has elegido la dificultad fácil, tienes {numero_intentos} intentos')
            opcion_valida=True
        elif opcion == 2:
            dificultad = 2
            numero_intentos = 12
            print(f'Has elegido la dificultad intermedia, tienes {numero_intentos} intentos')
            opcion_valida=True
        elif opcion == 3:
            dificultad = 3
            numero_intentos = 5
            print(f'Has elegido la dificultad intermedia, tienes {numero_intentos} intentos')
            opcion_valida=True
        else:
            print("❌ Opción inválida, intenta de nuevo.")
    return numero_intentos, dificultad
#Función para pedir el nivel del jugador en modo pareja
def pedir_nivel_pareja(nombre_1,nombre_2):

    opcion_valida=False
    print(f'Hola {nombre_1} y {nombre_2} , ha llegado la hora de decidir la dificultad del juego.')
    while opcion_valida == False:
        print("\n--- Nivel de dificultad ---")
        print("1. Pulsa 1 si queréis jugar el modo fácil y tener 20 intentos.")
        print("2. Pulsa 2 si queréis jugar el modo intermedio y tener 12 intentos.")
        print("3. Pulsa 3 si queréis jugar el modo difícil y tener 5 intentos.")

        opcion = int(input("Elige una opción: "))

        if opcion == 1:
            dificultad = 1
            numero_intentos = 20
            print(f'Habéis elegido la dificultad fácil, teneis {numero_intentos} intentos')
            opcion_valida=True
        elif opcion == 2:
            dificultad = 2
            numero_intentos = 12
            print(f'Hebéis elegido la dificultad intermedia, teneis {numero_intentos} intentos')
            opcion_valida=True
        elif opcion == 3:
            dificultad = 3
            numero_intentos = 5
            print(f'Habéis elegido la dificultad intermedia, tenéis {numero_intentos} intentos')
            opcion_valida=True
        else:
            print("❌ Opción inválida, intentadlo de nuevo.")
    return numero_intentos,dificultad
def resultado_jugadores_solitario(nombre,dificultad, encierto):
    archivo = "Resultados de adivina el número.xlsx"

    wb = load_workbook(archivo)
    ws = wb["Resultado jugadores"]
    fila=0
    nombre_encontrado = False

    for celda in ws["A"]:
        fila += 1 
        if celda.value == nombre:
            nombre_encontrado = True
            break
    if nombre_encontrado:
        ws.cell(row=fila, column=5).value += 1
        if dificultad == 1:
            ws.cell(row=fila, column=4).value += 1
        elif dificultad == 2:
            ws.cell(row=fila, column=3).value += 1        
        elif dificultad == 3:
            ws.cell(row=fila, column=2).value += 1
        if encierto:
            ws.cell(row=fila, column=7).value += 1
        if not encierto:
            ws.cell(row=fila, column=8).value += 1
        num = (ws.cell(row=fila, column=7).value + ws.cell(row=fila, column=8).value)
        den = (ws.cell(row=fila, column=5).value + ws.cell(row=fila, column=6).value)

        ws.cell(row=fila, column=9).value = num / den
        valor_col9 = ws.cell(row=fila, column=9).value   # columna H
        ws.cell(row=fila, column=10).value = 1 - valor_col9
    else:
        if dificultad == 1 and encierto:
            nueva_fila = [nombre, 0, 0, 1, 1, 0, 1, 0, 1,0]
        if dificultad == 1 and not encierto:
            nueva_fila = [nombre, 0, 0, 1, 1, 0, 0, 0, 0,0]
        elif dificultad == 2 and encierto:
            nueva_fila = [nombre, 0, 1, 0, 1, 0, 1, 0, 1,0]
        elif dificultad == 2 and not encierto:
            nueva_fila = [nombre, 0, 1, 0, 1, 0, 0, 0, 0,0]
        elif dificultad == 3 and encierto:
            nueva_fila = [nombre, 1, 0, 0, 1, 0, 1, 0, 1,0]
        elif dificultad == 3 and not encierto:
            nueva_fila = [nombre, 1, 0, 0, 1, 0, 0, 0, 0,0]
        ws.append(nueva_fila)
        ws.cell(row=fila, column=9).number_format = "0.00%"
        ws.cell(row=fila, column=10).number_format = "0.00%"  
def solitario():
    path='Resultados de adivina el número.xlsx'
    nombre = pedir_nombre_solitario()
    wb = load_workbook(path)
    niveles={1:'Fácil', 2:'Intermedio',3:'Difícil'}
    resultado={True:'Victoria', False:'Derrota'}
    numero_intentos,dificultad = pedir_nivel_solitario(nombre)
    print(f"\n🔢 El juego empieza con {numero_intentos} intentos.")
    encierto = False  #Esta variable nos indica que no se ha encontrado el num secreto. Si se encuentra, encierto pasara a ser True y se romperá el bucle siguiente.
    numero_elegido = random.randint(1,1001) #Se genera un número aleatorio entre el 1 y el 1000.
    i = 0  #Es el contador.
    while i in range(numero_intentos) and encierto == False:
        intento=int(input(f'Tu intento número {i+1} es:'))
        if intento == numero_elegido:
            os.system("afplay 'success-fanfare-trumpets-6185.mp3'")
            encierto=True
            print(f'Has acertado en tu intento {i+1}! El número secreto era {numero_elegido}!')
        if intento<numero_elegido:
            os.system("afplay 'mixkit-wrong-answer-bass-buzzer-948.wav'")
            print(f'El número {intento} es menor que el númerop secreto... Te quedan {numero_intentos-(i+1)} intentos!')
            i += 1
        if intento>numero_elegido:
            os.system("afplay 'mixkit-wrong-answer-bass-buzzer-948.wav'")
            print(f'El número {intento} es mayor que el númerop secreto... Te quedan {numero_intentos-(i+1)} intentos!')
            i += 1
    if encierto == False:
        print(f"Lo siento, no has adivinado el número secreto que era {numero_elegido}... Inténtalo otra vez!")
    ws=wb['Solitario']
    nueva_fila=[nombre,niveles[dificultad],resultado[encierto],i,numero_elegido,intento]
    ws.append(nueva_fila)
    wb.save(path)

def pareja():
    path='Resultados de adivina el número.xlsx'
    wb = load_workbook(path)
    niveles = {1:'Fácil', 2:'Intermedio',3:'Difícil'}
    resultado = {True:'Victoria', False:'Derrota'}
    nombre_1, nombre_2 = pedir_nombre_parejas()
    numero_intentos, dificultad = pedir_nivel_pareja(nombre_1, nombre_2)
    print(f"\n🔢 El juego empieza con {numero_intentos} intentos.")
    encierto = False  #Esta variable nos indica que no se ha encontrado el num secreto. Si se encuentra, encierto pasara a ser True y se romperá el bucle siguiente.
    numero_elegido = random.randint(1,1001) #Se genera un número aleatorio entre el 1 y el 1000.
    i = 0  #Es el contador.
    while i in range(numero_intentos) and encierto == False:
        intento_1=int(getpass(f'El intento número {i+ 1} de {nombre_1} es:'))
        if intento_1 == numero_elegido:
            os.system("afplay 'success-fanfare-trumpets-6185.mp3'")
            print(f'Enhorabuena {nombre_1}! Has acertado en tu intento {i+ 1}! El número secreto era {numero_elegido}!')
            ganador=nombre_1
            perdedor=nombre_2
            encierto = True
            break
        if intento_1<numero_elegido:
            os.system("afplay 'mixkit-wrong-answer-bass-buzzer-948.wav'")
            print(f'El número que me has dado es menor que el númerop secreto... Te quedan {numero_intentos-(i+1)} intentos!')
        if intento_1>numero_elegido:
            os.system("afplay 'mixkit-wrong-answer-bass-buzzer-948.wav'")
            print(f'El número que me has dado es mayor que el númerop secreto... Te quedan {numero_intentos-(i+ 1)} intentos!')
        intento_2=int(getpass(f'El intento número {i+1} de {nombre_2} es:'))
        if intento_2 == numero_elegido:
            os.system("afplay 'success-fanfare-trumpets-6185.mp3'")
            print(f'Enhorabuena {nombre_2}! Has acertado en tu intento {i+ 1}! El número secreto era {numero_elegido}!')
            ganador = nombre_2
            perdedor = nombre_1
            encierto = True
            break
        if intento_2<numero_elegido:
            os.system("afplay 'mixkit-wrong-answer-bass-buzzer-948.wav'")
            print(f'El número que me has dado es menor que el númerop secreto... Te quedan {numero_intentos-(i+ 1)} intentos!')
            i += 1
        if intento_2>numero_elegido:
            os.system("afplay 'mixkit-wrong-answer-bass-buzzer-948.wav'")
            print(f'El número que me has dado es mayor que el númerop secreto... Te quedan {numero_intentos-(i+ 1)} intentos!')
            i += 1
    if encierto == False:
        print(f"Lo siento, no has adivinado el número secreto, que era {numero_elegido}... Inténtalo otra vez!")
    ws=wb['Pareja']
    nueva_fila=[nombre_1,nombre_2,ganador,i,i,numero_elegido,f'{intento_1}-{intento_2}']
    ws.append(nueva_fila)
    wb.save(path)
def estadisticas():
    archivo = "Resultados de adivina el número.xlsx"

    dfs = pd.read_excel(archivo, sheet_name=["Solitario", "Pareja", "Resultado jugadores"])

    for nombre, df in dfs.items():
        print(f"\n========== {nombre} ==========")
        print(tabulate(df, headers='keys', tablefmt='psql'))      


 