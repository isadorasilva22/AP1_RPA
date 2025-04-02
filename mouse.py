# Descobrindo a posição do mouse
import pyautogui
import time

print("Mova o mouse para descobrir a posição do cursos.")
try:
    while True:
        x, y = pyautogui.position()
        print(f"Posição atual: x={x}, y={y}", end="\r")
        time.sleep(0.5)
except KeyboardInterrupt:
    print("\nCaptura interrompida")

