import pygame

# Dimensiones y configuración de la window
WIDTH = 800
HEIGHT = 600
SIZE_CELL = 50
COLOR_CELL = (30, 50, 60)
COLOR_BORDER = (0, 0, 0)

# Inicializar Pygame
pygame.init()

# Crear la window
window = pygame.display.set_mode((WIDTH, HEIGHT))
pygame.display.set_caption("Ejemplo de Matriz con Pygame")

# Dimensiones y estado inicial de la matriz
row = 5
col = 5
matriz = [[0] * col for _ in range(row)]

# Función para dibujar la matriz en la window
def dibujar_matriz():
  for i in range(row):
    for j in range(col):
        pygame.draw.rect(window, COLOR_CELL,(j * SIZE_CELL, i * SIZE_CELL,SIZE_CELL, SIZE_CELL))
        pygame.draw.rect(window, COLOR_BORDER,(j * SIZE_CELL, i * SIZE_CELL,SIZE_CELL, SIZE_CELL), 1)

# Bucle principal
exe = True
clock = pygame.time.Clock()

while exe:
  for evento in pygame.event.get():
    if evento.type == pygame.QUIT:
      exe = False
    elif evento.type == pygame.KEYDOWN:
      if evento.key == pygame.K_UP:
        row += 1
        matriz.append([0] * col)
      elif evento.key == pygame.K_DOWN:
        if row > 0:
            row -= 1
            matriz.pop()
      elif evento.key == pygame.K_LEFT:
        if col > 0:
            col -= 1
            for fila in matriz:
                fila.pop()
      elif evento.key == pygame.K_RIGHT:
        col += 1
        for fila in matriz:
            fila.append(0)

    # Limpiar la window
  window.fill((0, 0, 0))

    # Dibujar la matriz en la window
  dibujar_matriz()

    # Actualizar la window
  pygame.display.flip()

    # Controlar la velocidad de fotogramas
  clock.tick(60)

# Salir de Pygame
pygame.quit()
