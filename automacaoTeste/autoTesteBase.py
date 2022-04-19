import pyautogui
import time

def openGoogle():
    pyautogui.alert("Você entrou no modo de Automação. "
                    "Por favor não utilize o mouse nem o teclado!!")
    pyautogui.press("winleft")
    pyautogui.write("chrome")
    pyautogui.press("enter")
    pyautogui.sleep(4)

def toAreaDeTrabalho():
    pyautogui.hotkey("winleft", "d")


def openWhatsapChrome():
    openGoogle()
    pyautogui.write("https://web.whatsapp.com/")
    pyautogui.press("enter")

def openYoutubeChrome():
    openGoogle()
    pyautogui.moveTo(x=407, y=482)
    pyautogui.mouseDown()
    time.sleep(2)
    pyautogui.leftClick()

def localizationPositionMouse():
    pyautogui.alert("Você entrou no modo de Automação. "
                    "Te daremos 5 seconds para colocar o mouse na posição desejada!!")
    pyautogui.sleep(5)
    print(pyautogui.position())
    
def openTrello():
    openGoogle()
    time.sleep(3)
    pyautogui.moveTo(x=58, y=77)
    pyautogui.leftClick()
    pyautogui.moveTo(x=82, y=112)
    pyautogui.leftClick()

