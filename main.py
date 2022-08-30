import os
import pandas as pd


def check_input():
    for in_file in os.listdir("/mydir"):
        if in_file.endswith(".xlxs"):
            return 1
    return 0


def check_output():
    for out_file in os.listdir("/mydir"):
        if out_file.endswith(".pdf"):
            return 1
    return 0


def error_message(flag):
    if flag == 1:
        error_file = open("НЕТ ЭКСЕЛЬ ФАЙЛА.txt", "w+")
        error_file.write("Отсутствует эксель файл для расчета данных и заполнения в пдф!\n ")
        error_file.write("Поместите недостающий файл в папку к исполняемому файлу(count.exe) и "
                         "запустите программу заново \n")
        error_file.write("Спасибо ;)")
        error_file.close()
    if flag == 2:
        error_file = open("НЕТ ПДФ ФАЙЛА.txt", "w+")
        error_file.write("Отсутствует шаблонный пдф файл для вывода корректного результата!\n ")
        error_file.write("Поместите шаблон пдф файла в папку к исполняемому файлу(count.exe) и "
                         "запустите программу заново \n")
        error_file.write("Спасибо ;)")
        error_file.close()
    if flag == 3:
        error_file = open("НЕТ ЭКСЕЛЬ И ПДФ ФАЙЛОВ.txt", "w+")
        error_file.write("Отсутствует эксель файл и шаблон пдф\n ")
        error_file.write("Поместите недостающие файлы в папку к исполняемому файлу(count.exe) и "
                         "запустите программу заново \n")
        error_file.write("Спасибо ;)")
        error_file.close()
    if flag == 4:
        return 1


def read_xlxs_file():
    xlxs_file = pd.read_excel('sales.xlsx')
    data = pd.DataFrame(xlxs_file, columns=['Sales Date'])
    return data
