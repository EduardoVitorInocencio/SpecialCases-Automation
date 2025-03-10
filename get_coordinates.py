import pyautogui
import time

print("Posicione o mouse onde deseja e pressione 'Ctrl + C' para sair.")

try:
    while True:
        # Captura a posição atual do mouse
        x, y = pyautogui.position()
        print(f"Posição do mouse: X={x}, Y={y}", end="\r", flush=True)  # Atualiza a mesma linha no terminal
        time.sleep(0.1)  # Pequena pausa para não sobrecarregar o terminal
except KeyboardInterrupt:
    print("\nCaptura de coordenadas encerrada.")