import pyautogui
import time

pyautogui.moveTo(269,1066,2)
pyautogui.doubleClick()
time.sleep(1)
pyautogui.typewrite('Hello')
pyautogui.typewrite(['enter']) #[]안에 넣으면 그 키를 넣어줌
pyautogui.typewrite(['a','b','c']) # abc 출력됨
pyautogui.typewrite('hi')