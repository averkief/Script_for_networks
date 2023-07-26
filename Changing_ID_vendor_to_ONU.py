import os
import openpyxl
#import pexpect   # Нет в Winwows аттрибута spawn
from telnetlib import Telnet
import time
import keyboard

# Глобальные переменные
MAC_READY = list()  # MAC устройств которые препрошились
FILES_TABLES = list()  # список файлов таблиц
MAC_PASS = dict()  # Словарь содержащий пару MAC и пароль
DEVICE_IP = '192.168.101.1'
DEVICE_PORT = 23 # порт telnet
NAME_VENDOR = '---NEW VENDOR---'  # Указать название вендора

# Step 1. Пассинг файлов в текущем каталоге, отобрать xlsx, xls, ods, csv
# Библиотека только для файлов xlsx!
def find_files_table():
    '''Наполняем список файлами с таблицами'''
    global FILES_TABLES
    for file in os.listdir('.'):
        #if file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.ods'):
        if file.endswith('.xlsx'): # библиотека работает только с xlsx!!!
            FILES_TABLES.append(file)
            print(f'Найден файл: {file}')

# Step 2. Создать словарь с парами MAC пароль
def create_mac_pass():
    '''Наполняем словарь парами мак/пароль из полученных таблиц'''
    global MAC_PASS
    #print(f'до функции: {file}')
    for file in FILES_TABLES:
        #print(f'В функции и цикле {file}')
        table_full = openpyxl.load_workbook(file)    # открываем файл
        table_sheet = table_full.active              # открываем активный лист
        for row in table_sheet.iter_rows(values_only=True):
            MAC_PASS[row[0]] = row[1]
def filter_mac_address(dict_source):
    '''Убираем None значения'''
    filter_dist = {}
    for key, value in dict_source.items():   # Итерируемся по элементам исходного словаря
        if value is not None:                # Если значение не равно None, добавляем элемент в новый словарь
            filter_dist[key] = value
    return filter_dist

def format_mac_address(dict_filter):
    '''Убираем дефисы и ставим строчные буквы в MAC словаре'''
    format_dict = {}
    for key, value in dict_filter.items():        # Итерируемся по элементам исходного словаря
        new_key = key.lower().replace('-', '')    # Приводим ключ к нижнему регистру и убираем дефис
        format_dict[new_key] = value              # Добавляем элементы в новый словарь
    return format_dict


def onu_get_mac(hello_text):
    '''Возвращает пMAC адрес устройства из преветствия по telnet'''
    hello_decode = hello_text.decode("utf-8")
    #print(hello_decode)
    hello_utf = hello_decode.lower().split(': ')[2][0:17].replace(':', '')
    return hello_utf


def send_show_command(ip, vendor, dict_format):
    '''Фунция принимет IP устройство, Новое имя венора, словарь с мак/паролем.
        Подключается к устройству и передает все значения'''
    print(f'Подключению к устройству {ip} выполнено')
    global MAC_READY
    #with pexpect.spawn(f'telnet {ip}', timeout=10, encoding='utf-8') as telnet_device:
    #     telnet_device.expect()
    #    print(telnet_device.before)
    #telnet_device = telnetlib.Telnet(ip)
    #print(telnet_device.read_until())

    try:
        with Telnet(ip, port=23) as tn:
            telnet_hello_onu = tn.read_until(b'n: ')
            print(f'telnet соединение установлено')
            #print(telnet_hello_onu)
            onu_mac = onu_get_mac(telnet_hello_onu) # получение из преветствия MAC адрес устройства
            print(f'MAC устройства {onu_mac}')
            time.sleep(0.5)

            # Проверка наличии MAC в словаре сканированных таблиц
            if onu_mac in dict_format.keys():

                # Начала пререпрошивки ONU
                if onu_mac not in MAC_READY: # Проверка перепрошивалось ли ранее
                    MAC_READY.append(onu_mac)   # Добавление мака в список перепрошитых устройств
                    onu_curret_password = dict_format[onu_mac] + '\n'
                    #print(onu_curret_password)
                    onu_curret_password_carenca = bytes(onu_curret_password, encoding = 'utf-8') # Байтовое Получение занчение пароля по MAC адресу из таблицы
                    #print(onu_curret_password_carenca)
                    tn.write(b'manu\n') # отправляем логин для входа на устройство
                    output_login_ddd = tn.read_until(b'd: ') # Получаем вывод после входа.decode('utf-8')
                    #print(output_login_ddd)
                    tn.write(onu_curret_password_carenca)  # Передаем байтовую строку с паролем
                    time.sleep(1) # ждем ответа на вход ONU
                    #output_login = tn.read_very_lazy().decode('utf-8') # Получаем вывод после входа
                    output_entru = tn.read_until(b'# ')
                    output_entru_decoder = output_entru.decode("utf-8")

                    if output_entru_decoder[-3:] == '/# ' : # Если пароль подошел или не подошел, проверкой последнего символа
                        tn.write(b'gccli sys vendor HWTC\n')    # установка назавния вендора
                        time.sleep(0.2)
                        print(tn.read_until(b'# ').decode("utf-8"))
                        tn.write(b'gccli sys vendorid HWTC\n')    # установка id вендора
                        time.sleep(0.2)
                        print(tn.read_until(b'# ').decode("utf-8"))
                        tn.write(b'gccli sys save\n')   # сохранение изменений
                        time.sleep(0.5)
                        print(tn.read_until(b'# ').decode("utf-8"))
                        tn.write(b'gccli sys show\n')   # проверка иходных параметров
                        time.sleep(0.2)
                        print(tn.read_until(b'# ').decode("utf-8"))
                        #output_onu = tn.read_very_lazy().decode("utf-8")
                        #output_onu = tn.read_until(b'# ')
                        #print(output_onu)
                        print(f'ONU с MAC {onu_mac} успешно перепрошит на вендора HWTC')
                        tn.close()
                        print("Ожидание нажатия клавиши для продолжения...")
                        keyboard.read_key()
                        return
                    else:
                        print(f'Пароль к MAC {onu_mac} не ПОДОШЕЛ проверьте ТАБЛИЦУ')
                        tn.close()
                        print("Ожидание нажатия клавиши...")
                        keyboard.read_key()
                        #time.sleep(5)
                        return

                    return
                else:
                    print(f'Устройство с MAC {onu_mac} уже перепрошивалось')
                    #print("Ожидание нажатия клавиши...")
                    #keyboard.read_key()
                    time.sleep(2)
                    return
            else:
                print(f'Устройство с MAC {onu_mac} не найдено в ТАБЛИЦАХ')
                print("Ожидание нажатия клавиши...")
                keyboard.read_key()
                #time.sleep(2)
                return
    except TimeoutError:
        print(f'FAIL telnet НЕ ДОСТУПЕН')
        print("Ожидание нажатия клавиши для ПОВТОРНОГО ЗАПУСКА...")
        keyboard.read_key()
        return


# Запускающая часть
if __name__ == "__main__":
    find_files_table() # шаг 1.
    if FILES_TABLES != []:
        print(f'Поиск таблиц окончен.')
        create_mac_pass() # шаг 2.
        if len(MAC_PASS) > 1:
            MAC_PASS_FILTER = filter_mac_address(MAC_PASS)
            MAC_PASS_FORMAT = format_mac_address(MAC_PASS_FILTER)
            print(f'Просканированны пары MAC:Password')
            print(f'Внесено {len(MAC_PASS_FORMAT)} записей пар mac/password.')
            print("Ожидание нажатия клавиши, для запуска подключения к устройству")
            keyboard.read_key()
            #print(MAC_PASS_FORMAT)  # ДЕБАГ
            # Шаг 2 завершен.

            # Step 3. Создать список перепрошитых онушек по маку, создан MAC_READY
            # Step 4. Пинг железки по адресу 192.168.101.1
            print(f'Ищу устройство: {DEVICE_IP}')
            # получить ответ от DEVICE_IP например пингом и запускается начало соединения
            while True:
                # Проблема выводит сообщения о пингах
                if os.system('ping -n 1 ' + DEVICE_IP) == 0:    # Step 5. Удачный пинг, подключаемся по telnet
                    send_show_command(DEVICE_IP, NAME_VENDOR, MAC_PASS_FORMAT)
                else:
                    print(f'{DEVICE_IP} Недоступен, повторяю попытку')
                    time.sleep(2)
                    #click.pause()  #почему-то не работает



        else:
            print(f'Словарь MAC адресов пуст, программа завершена')
            print("Ожидание нажатия клавиши...")
            keyboard.read_key()
    else:
        print(f'Таблицы не найдены, программа завершена')
        print("Ожидание нажатия клавиши...")
        keyboard.read_key()
