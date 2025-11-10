#Definiremos una función que nos cree un menú para nuestro juego.
import Funciones

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

        opcion = int(input("Elige una opción: "))

        if opcion == 1:
            Funciones.solitario()
        elif opcion == 2:
            Funciones.pareja()
        elif opcion == 3:
            Funciones.estadisticas()
        elif opcion == 4:
            print("\n👋 ¡Gracias por jugar!")
            break
        else:
            print("❌ Opción inválida, intenta de nuevo.")
menu()