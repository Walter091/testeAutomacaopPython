import AutoExplorer
import AutoBase
import pyautogui as py
import time
# obj = AutoExplorer.TesteAutomacao(numArquivo=None, nomeArquivo=None, status=None)
# obj.Index()
# AutoBase.localizationPositionMouse()


py.moveTo(x=88, y=93)
py.leftClick()
py.rightClick()
time.sleep(1)
py.press("down", 20)
time.sleep(1)

py.press("right")
time.sleep(1)
py.press("down")
time.sleep(1)
py.press("enter")
time.sleep(10)

py.moveTo(x=88, y=93)
py.leftClick()
py.rightClick()
time.sleep(1)
py.press("down", 20)
time.sleep(1)

py.press("right")
time.sleep(1)
py.press("enter")
time.sleep(1)
py.write("commitMsg")
time.sleep(1)
py.press("tab")
py.press("enter")
time.sleep(10)

py.moveTo(x=88, y=93)
py.leftClick()
py.rightClick()
time.sleep(1)
py.press("down", 20)
time.sleep(1)
py.press("right")

py.moveTo(x=608, y=293)
py.leftClick()
time.sleep(1)
py.moveTo(x=717, y=569)
py.leftClick()


