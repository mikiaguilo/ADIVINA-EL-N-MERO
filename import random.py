
import pygame

pygame.init()
screen = pygame.display.set_mode((600, 400))
pygame.display.set_caption("Texto en la ventana")
font = pygame.font.Font(None, 28)  # Fuente por defecto
texto = [
    "Hola! Bienvenidos al juego 'Adivina el número'.",
    "Objetivo: adivinar un número entre 1 y 1000.",
    "Dificultades: 1) Fácil (20)  2) Intermedio (12)  3) Difícil (5)",
    "Modos: Solitario o en pareja."
]

running = True
while running:
    for event in pygame.event.get():
        if event.type == pygame.QUIT:
            running = False

    screen.fill((25, 28, 35))  # fondo
    y = 20
    for linea in texto:
        surface = font.render(linea, True, (230, 230, 230))
        screen.blit(surface, (20, y))
        y += surface.get_height() + 6  # espacio entre líneas

    pygame.display.flip()

pygame.quit()