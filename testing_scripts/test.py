import pyautogui
import keyboard

print("Hover over the desired location and press 'a' to get the coordinates.")

while True:
    if keyboard.is_pressed('a'):
        print(pyautogui.position())
        break
