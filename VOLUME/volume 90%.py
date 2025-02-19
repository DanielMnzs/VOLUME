import pyautogui
import pandas as pd
import time 



caminho = r"//depo0903/gpa$/PAC/Daniel Menezes/Python//VOLUME/Package.xlsx"
leitura = pd.read_excel(caminho)
df = pd.DataFrame(leitura)
print (df)


pyautogui.press("win")
time.sleep(1)
pyautogui.write("opera")
time.sleep(1)
pyautogui.press("enter")
time.sleep(1)
pyautogui.write("manhattan")
time.sleep(1)
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
pyautogui.press("down")
time.sleep(1)
pyautogui.press("enter")
time.sleep(2)
pyautogui.leftClick(52, 87)
time.sleep(2)
pyautogui.write("itens")
time.sleep(2)
pyautogui.leftClick(123, 228)
time.sleep(2)
pyautogui.leftClick(795, 121)
time.sleep(2)
pyautogui.leftClick(946, 133)

x = 0
y = 0

while True:
    plu = df.iloc[x , 0]
    volume = df.iloc[x , 2]
    pyautogui.doubleClick(242, 193)    
    time.sleep(2)
    pyautogui.write(str(plu))
    pyautogui.leftClick(1097, 191)
    time.sleep(2)
    pyautogui.leftClick(73, 299)
    time.sleep(2)
    pyautogui.leftClick(71, 703)
    time.sleep(2)
    pyautogui.leftClick(201, 225)
    time.sleep(2)
    pyautogui.leftClick(79, 695)
    time.sleep(2)
    pyautogui.doubleClick(236, 333)
    time.sleep(1)
    pyautogui.hotkey("backspace")
    time.sleep(2)
    pyautogui.write("0,")
    time.sleep(1)
    pyautogui.write(volume)    
    time.sleep(2)
    pyautogui.leftClick(1237, 695)
    time.sleep(2)
    pyautogui.leftClick(62, 143)
    time.sleep(2)
    print(plu)
    time.sleep(1)
    print(volume)
    time.sleep(1)
    x = x + 1
    print(x)
    print(plu)
    print("DEV: DANIEL MENEZES")