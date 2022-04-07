from pynput.mouse import Button, Controller
import time
import win32com.client

import AMOdata_laptop
import AMOdata_desktop

mouse = Controller()
shell = win32com.client.Dispatch('WScript.Shell')

time.sleep(5)
print('Курсор найден здесь {0}'.format(
    mouse.position))


def clickerBotWorking(script, per, device):
    if device == "desktop":
        device = AMOdata_desktop
    else:
        device = AMOdata_laptop

    for all_work in range(per):
        for index in range(len(script)):
            if index in device.easy_index:
                time.sleep(device.delays[0])
            elif index in device.middle_index:
                time.sleep(device.delays[1])

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


def clickerBuyPython(script, per, device):
    if device == "desktop":
        device = AMOdata_desktop
    else:
        device = AMOdata_laptop

    for all_work in range(per):
        for index in range(len(script)):
            if index in device.easy_index:
                time.sleep(device.delays[0])
            elif index in device.middle_index:
                time.sleep(device.delays[1])

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


def clickerFirstContact(script, per, device):
    if device == "desktop":
        device = AMOdata_desktop
    else:
        device = AMOdata_laptop
    
    for all_work in range(per):
        for i in range(len(script)):
            if i in device.easy_index:
                time.sleep(device.delays[0])
            elif i in device.middle_index:
                time.sleep(device.delays[1])

            mouse.position = script[i]
            print('Курсор сейчас здесь {0}'.format(mouse.position))
            mouse.press(Button.left)
            mouse.release(Button.left)

            if i == 3:
                time.sleep(2)
                # mouse.press(Button.left)
                # mouse.release(Button.left)
                shell.SendKeys('^v')
                print("Текст вставлен")



if __name__ == "__main__":
    time.sleep(3)
    clickerFirstContact(AMOdata_laptop.script_first_contact_laptop, 5, "laptop")
