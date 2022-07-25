import time  # для создания эффекта печатания текста с задержкой
# секундомер. Необходимо установить (pip install stopwatch.py)
from stopwatch import Stopwatch
import msoffcrypto  # для работы с офисными файлами
import io  # для работы с файлами с расширениями docx и xlsx
# для созлания разноцветного текста в программе. Необходимо установить (pip install colorama)
from colorama import init, Fore, Back, Style
init()


def console_picture():
    print(Style.BRIGHT + Fore.CYAN)
    print("   ########   ########  ##      ##   #######   ########    ######## ")
    print("   ########   ########  ##      ##   ##    ##  ##          ######## ")
    print("   ##    ##   ##    ##  ##     ###   ########  ########       ##  ")
    print("   ##    ##   ########  ##  ##  ##   ##    ##  ##             ##  ")
    print("   ##    ##   ##        ## #    ##   ########  ########       ##  ")
    print("   ##    ##   ##        ##      ##   #######   ########       ##  ")
    print()
    print()
    print("   ########      ###    ##      ##   ########   ##      ##  ########     @@ ")
    print("   ##           #####   ##      ##   ##    ##   ##      ##  ########     @@ ")
    print("   ######      ##   ##  ##########   ##    ##   ##      ##     ##        @@ ")
    print("   ##    ##   ##     ## ##      ##   ##    ##   ##    ####     ##        @@ ")
    print("   ########   ######### ##      ##   ##    ##   ##  ##  ##     ##        @@ ")
    print("   ######     ##     ## ##      ##  ##########  ## #    ##     ##           ")
    print("                                    ##      ##  ##      ##     ##        @@ ")
    print()
    print()
    print()


console_picture()


wordTitle = '          Программа для взлома запароленных документов (word, exel, ppt) '
# создаем эффект печатания текста (текст выводиться с задержкой)
for i in wordTitle:
    print(i.upper(), end="")
    time.sleep(0.03)

stopwatch = Stopwatch(2)  # 2 это десятична точность для секундомера
decrypted = io.BytesIO()  # для декодирования файлов с расширением docx, xlsx


def crack_password(password_list, file_for_breaking):

    indx = 0
    cnt = len(list(open(password_list, 'rb')))
    # открываем файл (with open() as file: - пишем так, чтобы потом не писать комманду закрытия файла (close()). rb открытие в двоичном режиме )
    with open(password_list, 'rb') as file:
        stopwatch.restart()
        for line in file:
            for word in line.split():
                # вычисляем в процентном соотношении количество пребранных паролей
                x = (indx+1)/cnt * 100
                # отсекаем цифры после запятой до 2-х, чтоб не получалось вроде такого: 0.9834539503%
                x = float('{:.2f}'.format(x))
                print(
                    f'Количество перебранных паролей {indx} ----- Процент перебранных паролей {x} ---- Прошло времени {str(stopwatch)}\r', end="")  # подсчитываем количество перебранных паролей. Вывод текста в одну строку с затиранием предыдущего
                try:
                    indx += 1
                    # password если файл зашифрован, передайте пароль в этом аргументе (по умолчанию: None) .
                    file_for_breaking.load_key(password=word.decode('utf8'))
                    # для работы с файлами с расширениями docx и xlsx
                    file_for_breaking.decrypt(decrypted)
                    print("\n")
                    print(Style.BRIGHT + Fore.GREEN)
                    print("Пароль найден в строке: ", indx)
                    # Декодирует байтстроку в строку.
                    print("Пароль: ", word.decode())
                    stopwatch.reset()  # сбрасываем счетчик на 0
                    # После нахождения пароля, спрашиваем про желание продолжить взламывать пароли
                    print(Style.BRIGHT + Fore.YELLOW)
                    continue_work = input(
                        "Хотите продолжить? Если да, то нажмите букву 'д'")
                    if (continue_work == 'д'):
                        main_data()
                        # return True
                    else:
                        return True

                except:

                    continue
    return False


def main_data():
    print(Style.BRIGHT + Fore.YELLOW)
    code_file = input("\nВведите адрес запароленного документа ")
    # # делаем проверку рассширенния взламываемого файла
    # if not any(map(code_file.endswith, ('.doc', '.docx', '.xlsx', '.rtf'))):
    #     print(Style.BRIGHT + Fore.RED)
    #     print(
    #         "Вы указали неверный файл для взлома. Файл не является документом MSOffice")
    #     main_data()
    # делаем проверку рассширенния взламываемого файла
    if any(map(code_file.endswith, ('.docx', '.xlsx', '.rtf', '.pptx'))):
        print(Style.BRIGHT + Fore.RED)
        print(
            """Файлы с таким расширением очень устойчивы к атакам грубой силы, 
поэтому перебор паролей будет происходить крайне низко!!!""")
    print(Style.BRIGHT + Fore.YELLOW)
    password_list = input("Введите адресс словаря ")

    # Инициализируем
    code_file = open(code_file, 'rb')
    file_for_breaking = msoffcrypto.OfficeFile(code_file)
    # print(file_for_breaking.file)

    # подсчитываем количесвто слов в словаре
    cnt = len(list(open(password_list, 'rb')))

    print("Количество паролей в данном словаре ", cnt)

    if crack_password(password_list, file_for_breaking) == False:
        print(Style.BRIGHT + Fore.RED)
        print("\nПароль не найден. Попробуйте другой словарь ")
        main_data()


main_data()
