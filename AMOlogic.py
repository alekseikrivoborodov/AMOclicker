from pynput.mouse import Button, Controller
import time
import win32com.client

import AMOdata_laptop

mouse = Controller()
shell = win32com.client.Dispatch('WScript.Shell')

time.sleep(5)
print('Курсор найден здесь {0}'.format(
    mouse.position))


def clickerBotWorking(script, per):
    for all_work in range(per):
        for index in range(len(script)):
            if index in AMOdata_laptop.easy_index:
                time.sleep(AMOdata_laptop.delays_laptop[0])
            elif index in AMOdata_laptop.middle_index:
                time.sleep(AMOdata_laptop.delays_laptop[1])

            mouse.position = script[index]
            print('Курсор сейчас здесь {0}'.format(mouse.position))
            mouse.press(Button.left)
            mouse.release(Button.left)

            if index == 3:
                time.sleep(2)
                mouse.press(Button.left)
                mouse.release(Button.left)
                shell.SendKeys('^v')
                print("Текст вставлен")



def clickerBuyPython(script, per):
    for all_work in range(per):
        for index in range(len(script)):
            if index in AMOdata_laptop.easy_index:
                time.sleep(AMOdata_laptop.delays_laptop[0])
            elif index in AMOdata_laptop.middle_index:
                time.sleep(AMOdata_laptop.delays_laptop[1])

            mouse.position = script[index]
            print('Курсор сейчас здесь {0}'.format(mouse.position))
            mouse.press(Button.left)
            mouse.release(Button.left)

            if index == 3:
                time.sleep(2)
                mouse.press(Button.left)
                mouse.release(Button.left)
                shell.SendKeys('^v')
                print("Текст вставлен")

            elif index == 10:
                time.sleep(2)
                mouse.press(Button.left)
                print("Начало скролл")

            elif index == 11:
                mouse.press(Button.left)
                mouse.release(Button.left)
                print("Конец скролл")


def amazingClicker(function, script, per):
    function(script, per)


if __name__ == "__main__":
    time.sleep(3)
    # amazingClicker(clickerBuyPython(AMOdata_laptop.script_buy_python_laptop, 28))
    amazingClicker(clickerBotWorking(AMOdata_laptop.script_bot_working_laptop, 34))
