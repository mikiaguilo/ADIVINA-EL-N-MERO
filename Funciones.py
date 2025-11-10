import random
import pandas as pd
import openpyxl
import os
import getpass
from openpyxl import load_workbook
#Definiremos una función que nos cree un menú para nuestro juego.
def introduccion():
    print(
    "Hola! Bienvenidos al juego de 'Adivina el número'. Como bien indica el título de este juego,\n"
    "el objetivo de este será adivinar un número generado aleatoriamente por la máquina entre el 1 y el 1000.\n\n"
    "Habrá tres dificultades:\n"
    "  1) Nivel fácil (20 intentos)\n"
    "  2) Nivel intermedio (12 intentos)\n"
    "  3) Nivel difícil (5 intentos)\n\n"
    "También disponemos del modo de juego solitario o en pareja.\n"
    "Además, se creará un archivo Excel donde se guardarán las estadísticas "
    "dependiendo de los diferentes modos de juego."
)
def menu():
    while True:
        print("\n--- MENÚ PRINCIPAL ---")
        print("1. Pulsa 1 si quieres jugar en modo solitario.")
        print("2. Pulsa 1 si queréis jugar en modo pareja.")
        print("3. Pulsa 3 para ver las estadísticas del juego")
        print("4. Pulsa 4 para salir.")

        opcion = input("Elige una opción: ")

        if opcion == "1":
            juegos.juego_1()
        elif opcion == "2":
            juegos.juego_2()
        elif opcion == "3":
            juegos.juego_3()
        elif opcion == "4":
            print("\n👋 ¡Gracias por jugar!")
            break
        else:
            print("❌ Opción inválida, intenta de nuevo.")

def pedir_nombre_solitario():

    nombre = input("Hola! Cómo te llamas=")
    return nombre

def pedir_nombre_parejas():

    nombre_1 = input('¿Cómo se llama el primer integrante?')
    nombre_2 = input('¿Cómo se llama el segundo integrante?')
    return nombre_1,nombre_2
#Función para pedir el nivel del jugador en modo solitario
def pedir_nivel_solitario():

    opcion_valida=False
    nombre = pedir_nombre_solitario()
    print(f'Hola {nombre}, ha llegado la hora de decidir la dificultad del juego.')
    while opcion_valida == False:
        print("\n--- Nivel de dificultad ---")
        print("1. Pulsa 1 si quieres jugar el modo fácil y tener 20 intentos.")
        print("2. Pulsa 2 si quieres jugar el modo intermedio y tener 12 intentos.")
        print("3. Pulsa 3 si quieres jugar el modo difícil y tener 5 intentos.")

        opcion = int(input("Elige una opción: "))

        if opcion == "1":
            numero_intentos = 20
            print(f'Has elegido la dificultad fácil, tienes {numero_intentos} intentos')
            opcion_valida=True
        elif opcion == "2":
            numero_intentos = 12
            print(f'Has elegido la dificultad intermedia, tienes {numero_intentos} intentos')
            opcion_valida=True
        elif opcion == "3":
            print(f'Has elegido la dificultad intermedia, tienes {numero_intentos} intentos')
            opcion_valida=True
        else:
            print("❌ Opción inválida, intenta de nuevo.")
    return numero_intentos
#Función para pedir el nivel del jugador en modo pareja
def pedir_nivel_pareja():

    opcion_valida=False
    nombre_1, nombre_2 = pedir_nombre_parejas()
    print(f'Hola {nombre_1} y {nombre_2} , ha llegado la hora de decidir la dificultad del juego.')
    while opcion_valida == False:
        print("\n--- Nivel de dificultad ---")
        print("1. Pulsa 1 si queréis jugar el modo fácil y tener 20 intentos.")
        print("2. Pulsa 2 si queréis jugar el modo intermedio y tener 12 intentos.")
        print("3. Pulsa 3 si queréis jugar el modo difícil y tener 5 intentos.")

        opcion = int(input("Elige una opción: "))

        if opcion == 1:
            numero_intentos = 20
            print(f'Habéis elegido la dificultad fácil, teneis {numero_intentos} intentos')
            opcion_valida=True
        elif opcion == 2:
            numero_intentos = 12
            print(f'Hebéis elegido la dificultad intermedia, teneis {numero_intentos} intentos')
            opcion_valida=True
        elif opcion == 3:
            print(f'Habéis elegido la dificultad intermedia, tenéis {numero_intentos} intentos')
            opcion_valida=True
        else:
            print("❌ Opción inválida, intentadlo de nuevo.")
    return numero_intentos
def solitario():
    path='Resultados de adivina el número.xlsx'
    nombre=pedir_nombre_solitario()
    wb = load_workbook(path)
    niveles={1:'Fácil', 2:'Intermedio',3:'Difícil'}
    resultado={True:'Victoria', False:'Derrota'}
    numero_intentos = pedir_nivel_solitario()
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
    niveles={1:'Fácil', 2:'Intermedio',3:'Difícil'}
    resultado={True:'Victoria', False:'Derrota'}
    numero_intentos = pedir_nivel_pareja()
    nombre_1, nombre_2 = pedir_nombre_parejas
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
    nueva_fila=[nombre_1,nombre_2,ganador,j,j,f'{intento_1}-{intento_2}']
 