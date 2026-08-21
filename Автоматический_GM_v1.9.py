import sys
import os
import time
import shutil
import re
import win32clipboard
import pyautogui
import ctypes
from ctypes import wintypes
from datetime import datetime


# ----------------------------------------------------------------------
# Настройки
# ----------------------------------------------------------------------
# Настройки pyautogui
pyautogui.FAILSAFE = True  # перемещение мыши в левый верхний угол останавливает скрипт
pyautogui.PAUSE = 0.5     # пауза между командами pyautogui

# Пути к папкам (можно изменить при необходимости)
SOURCE_FOLDER = r"C:\1\AppliedBiosystems\GeneMapperID-X\Client"
DEST_BASE = r"C:\GM_RESULTS"
TEMP_FOLDER = r"C:\GM_TEMP"


# ----------------------------------------------------------------------
# Глобальные переменные для режимов работы
# ----------------------------------------------------------------------
MODE = "full"          # "full", "selective", "exclude", "multi"
SELECTIVE_LIST = []
EXCLUDE_LIST = []

# Для режима "multi"
MULTI_EXCLUDE_LIST = []   # список полных названий объектов для исключения
MULTI_BASE_NUMBER = None
MULTI_SAVED_CONTROLS = set()
MULTI_TXT_SAVED = False

# Для защиты от дубликатов (режимы 1-3)
processed_numbers = set()


# ----------------------------------------------------------------------
# Функции работы с клавиатурой и текстом
# ----------------------------------------------------------------------
def set_clipboard_text(text):
    """Помещает текст в буфер обмена."""
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return True
    except Exception as e:
        print(f"[Ошибка] Запись в буфер обмена: {e}")
        return False


def get_clipboard_text():
    """Возвращает текст из буфера обмена."""
    try:
        win32clipboard.OpenClipboard()
        data = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
        win32clipboard.CloseClipboard()
        return data.strip()
    except Exception as e:
        print(f"[Ошибка] Чтение буфера обмена: {e}")
        return ""


def get_numlock_state():
    """Возвращает True, если NumLock включен, иначе False."""
    try:
        # Используем GetKeyState с виртуальным кодом VK_NUMLOCK (0x90)
        # Старший бит определяет состояние: 1 = включен
        hllDll = ctypes.WinDLL("User32.dll")
        state = hllDll.GetKeyState(0x90)
        return (state & 0x0001) != 0  # Младший бит показывает переключение (toggle)
    except Exception as e:
        print(f"[Ошибка] Не удалось определить состояние NumLock: {e}")
        return False  # По умолчанию считаем выключенным


def set_numlock_state(on):
    """Включает NumLock, если on=True; выключает, если on=False.
    Использует keybd_event для эмуляции нажатия клавиши NumLock.
    """
    # Виртуальный код клавиши NumLock
    VK_NUMLOCK = 0x90
    # Флаги: 0 = нажатие, 2 = KEYEVENTF_KEYUP (отпускание)
    KEYEVENTF_KEYUP = 0x0002
    
    try:
        # Получаем текущее состояние
        current = get_numlock_state()
        if current == on:
            print(f"[NumLock] Уже в требуемом состоянии: {'ВКЛЮЧЕН' if on else 'ВЫКЛЮЧЁН'}")
            return True
        
        # Эмулируем нажатие и отпускание NumLock
        ctypes.windll.user32.keybd_event(VK_NUMLOCK, 0, 0, 0)          # Нажатие
        ctypes.windll.user32.keybd_event(VK_NUMLOCK, 0, KEYEVENTF_KEYUP, 0)  # Отпускание
        time.sleep(0.2)  # Небольшая задержка для применения
        print(f"[NumLock] Состояние изменено на: {'ВКЛЮЧЕН' if on else 'ВЫКЛЮЧЁН'}")
        return True
    except Exception as e:
        print(f"[Ошибка] Не удалось изменить состояние NumLock: {e}")
        return False


def type_text(text):
    """Ввод текста через буфер обмена (Ctrl+V)."""
    if not text:
        return
    if set_clipboard_text(text):
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.2)
    else:
        # запасной вариант – посимвольный ввод
        pyautogui.write(text, interval=0.05)
        time.sleep(0.2)


def press_keys(keys, delay=0.2):
    """Последовательное нажатие клавиш."""
    for key in keys:
        pyautogui.press(key)
        time.sleep(delay)


def press_hotkey(keys, delay=0.3):
    """Нажатие комбинации клавиш (например, ['ctrl','c'])."""
    pyautogui.hotkey(*keys)
    time.sleep(delay)


def wait(sec):
    """Задержка с выводом сообщения (опционально)."""
    time.sleep(sec)


# ----------------------------------------------------------------------
# Блок функций для работы с образцами
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# Функции для 1 режима (одно заключение - один объект)
# ----------------------------------------------------------------------
def process_one_sample():
    """Обработка одного образца для режима 1 (тотальное сохранение)."""
    global processed_numbers
    current_year = datetime.now().strftime("%y")

    # --- 1. Получение номера заключения ---
    print("[Навигация] Перемещение к ячейке с номером заключения...")
    press_keys(['right', 'right'])
    pyautogui.press('enter')
    wait(0.5)

    print("Выделение номера заключения...")
    for _ in range(4):
        press_hotkey(['ctrl', 'shift', 'left'])
        wait(0.1)

    press_hotkey(['ctrl', 'c'])
    wait(0.5)
    pyautogui.press('enter')

    case_number_raw = get_clipboard_text()
    if not case_number_raw:
        print("[Ошибка] Не удалось скопировать номер заключения. Прерывание.")
        return False

    case_number = clean_case_number_suffix(case_number_raw)
    print(f"[Итерация] Текущий объект: '{case_number_raw}' -> Очищенный: '{case_number}'")

    # --- 2. Немедленная остановка, если текущий объект – маркер ---
    if case_number in ("AL", "Al", "K+", "K-"):
        print(f"[Завершение] Обнаружен маркер: '{case_number}'. Программа завершается.")
        return False

    # --- 3. Проверка дубликата ---
    base_num = extract_case_number(case_number)
    if base_num in processed_numbers:
        print(f"[Дубликат] Номер '{base_num}' уже обработан в этом сеансе.")
        print("[Прерывание] Программа будет остановлена. Проверьте данные и запустите заново.")
        input("Нажмите Enter для завершения программы...")
        return False
    else:
        processed_numbers.add(base_num)
        print(f"[Обработка] Номер заключения '{base_num}' добавлен в обработанные")

    # --- 4. Сохранение текущего объекта (всегда в режиме 1) ---
    print(f"[Сохранение] Начинаем обработку объекта '{case_number}'...")
    perform_save_operations(case_number, current_year)

    # --- 5. Анализ следующего объекта ---
    next_info = check_next_object()
    next_case = next_info['next_case']
    has_next = next_info['has_next']

    # Если следующий объект – маркер или отсутствует, завершаем работу
    if not has_next:
        print("[Завершение] Следующий объект отсутствует или является маркером. Программа завершается.")
        # delete_current_object_simple()  # раскомментируйте, если хотите удалять последний объект
        return False

    # --- 6. Следующий объект существует и не маркер – создаём под него новый проект ---
    create_new_project_mode_1(next_case, current_year)

    # --- 7. Удаляем уже обработанный текущий объект ---
    print(f"[Очистка] Удаление обработанного объекта '{case_number}'...")
    delete_current_object_simple()

    # Продолжаем цикл
    return True


def perform_save_operations(case_number, current_year):
    """Выполняет операции сохранения фореграмм и файлов."""
    # Возврат в начальную точку
    print("[Навигация] Возврат в начальную точку...")
    pyautogui.press(['left'])
    press_hotkey(['shift', 'left'])
    wait(1)
    
    # Фореграмма (объект)
    print("[Операции сохранения файлов] Открытие и сохранение фореграммы объекта...")
    press_hotkey(['ctrl', 'l'])
    wait(3)
    pyautogui.press('alt')
    for _ in range(3):
        pyautogui.press('down')
        wait(0.1)
    pyautogui.press('enter')
    wait(2)
    pyautogui.press('enter')
    wait(5)
    pyautogui.press('esc')
    wait(1)

    # Фореграмма К+
    print("[Операции сохранения файлов] Переход на фореграмму К+...")
    press_hotkey(['ctrl', 'shift', 'pagedown'])
    wait(0.5)
    press_hotkey(['ctrl', 'shift', 'pagedown'])
    wait(0.5)
    press_hotkey(['shift', 'up'])

    print("[Операции сохранения файлов] Сохранение фореграммы К+...")
    press_hotkey(['ctrl', 'l'])
    wait(3)
    pyautogui.press('alt')
    for _ in range(3):
        pyautogui.press('down')
        wait(0.1)
    pyautogui.press('enter')
    wait(2)
    pyautogui.press('enter')
    wait(5)
    pyautogui.press('esc')
    wait(1)

    # Фореграмма К-
    print("[Операции сохранения файлов] Переход на фореграмму К-...")
    press_hotkey(['shift', 'down'])

    # Установка высоты оси Y
    print("[Операции сохранения файлов] Установка высоты фореграммы по оси Y...")
    press_hotkey(['ctrl', 'l'])
    wait(3)
    pyautogui.press('alt')
    for _ in range(3):
        pyautogui.press('right')
        wait(0.1)
    pyautogui.press('down')
    wait(0.1)
    pyautogui.press('right')
    wait(0.1)
    for _ in range(3):
        pyautogui.press('down')
        wait(0.1)
    pyautogui.press('enter')
    wait(1)
    for _ in range(2):
        pyautogui.press('tab')
        wait(0.1)
    press_hotkey(['ctrl', 'shift', 'left'])
    wait(0.1)
    pyautogui.press('backspace')
    wait(0.1)
    type_text("400")
    pyautogui.press('tab')
    wait(0.1)
    pyautogui.press('enter')
    wait(1)

    print("[Операции сохранения файлов] Сохранение фореграммы К-...")
    pyautogui.press('alt')
    for _ in range(3):
        pyautogui.press('down')
        wait(0.1)
    pyautogui.press('enter')
    wait(2)
    pyautogui.press('enter')
    wait(5)
    pyautogui.press('esc')
    wait(1)

    # Возврат в начало таблицы и добавление года
    print("[Операции сохранения файлов] Возврат в начало таблицы...")
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.2)
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.2)

    # Добавляем год
    press_keys(['right', 'right'])
    pyautogui.press('enter')
    wait(0.2)
    print("[Операции сохранения файлов] Добавление года...")
    type_text(f"-{current_year}")
    pyautogui.press('enter')
    wait(0.3)

    # Возврат в начальную точку
    pyautogui.press(['left'])
    press_hotkey(['shift', 'left'])

    # Сохранение Genotypes Table.txt
    print("[Операции сохранения файлов] Сохранение Genotypes Table.txt...")
    press_hotkey(['ctrl', 'shift', '3'])
    wait(1)
    press_hotkey(['ctrl', 'e'])
    wait(2)
    for _ in range(2):
        pyautogui.press('tab')
        wait(0.1)
    pyautogui.press('enter')
    wait(3)

    # Перенос файлов
    move_files(case_number, current_year)

    # Сохранение текущего проекта
    print("Сохранение текущего проекта...")
    press_hotkey(['ctrl', 'shift', '1'])
    wait(2)
    press_hotkey(['ctrl', 's'])
    wait(5)


def check_next_object():
    """Проверяет, есть ли следующий объект в списке и возвращает информацию о нем."""
    print("[Проверка] Определение следующего объекта...")
    
    try:
        # Переходим на следующую строку
        print("[Проверка] Чтение содержимого ячейки с номером...")
        pyautogui.press('down')
        wait(0.5)
        
        # Переходим к ячейке с номером
        press_keys(['right', 'right'])
        wait(0.1)
        pyautogui.press('enter')
        wait(0.5)
        
        # Выделяем и копируем
        for _ in range(4):
            press_hotkey(['ctrl', 'shift', 'left'])
            wait(0.1)
        press_hotkey(['ctrl', 'c'])
        wait(0.5)
        next_case_raw = get_clipboard_text()
        pyautogui.press('esc')
        wait(0.5)
        
        # Возвращаемся на предыдущую строку
        pyautogui.press('up')
        wait(0.5)
        
        # Очищаем суффикс
        next_case = clean_case_number_suffix(next_case_raw) if next_case_raw else ""
        
        # Проверяем, является ли следующий объект маркером завершения
        if not next_case or next_case in ("AL", "Al", "K+", "K-"):
            print(f"[Проверка] Следующий объект: '{next_case}' - это маркер завершения или пустота")
            return {'has_next': False, 'next_case': next_case, 'next_case_raw': next_case_raw}
        else:
            print(f"[Проверка] Следующий объект найден: '{next_case}'")
            return {'has_next': True, 'next_case': next_case, 'next_case_raw': next_case_raw}
            
    except Exception as e:
        print(f"[Ошибка] При проверке следующего объекта: {e}")
        return {'has_next': False, 'next_case': '', 'next_case_raw': ''}


def create_new_project_mode_1(next_case, current_year):
    """Создает новый проект для режима 1 (full)."""
    # Возврат в начало строки перед созданием проекта
    print("[Навигация] Возврат в начало строки...")
    pyautogui.press('left')
    wait(0.1)
    press_hotkey(['shift', 'left'])
    wait(0.1)

    # Обеспечение последующей работы Alt
    print("[Навигация] Обеспечение работы Alt...")
    press_hotkey(['ctrl', 'l'])
    wait(3)
    pyautogui.press('esc')
    wait(1)

    print(f"Создание нового проекта...")
    pyautogui.press('alt')
    wait(0.5)
    for _ in range(4):
        pyautogui.press('down')
        wait(0.1)
    pyautogui.press('enter')
    wait(3)
    print("Извлечение номера заключения для следующего объекта...")
    next_base_number = extract_case_number(next_case)
    if not next_base_number:
        print(f"[Ошибка] Не удалось извлечь номер заключения из следующего объекта '{next_case}'. Использую полное имя.")
        next_base_number = next_case
    else:
        print(f"[ИмяПроекта] Из полного имени следующего объекта '{next_case}' извлечён номер '{next_base_number}'")
    new_project_name = f"{next_base_number}-{current_year}"
    print(f"Создание нового проекта с именем {new_project_name} (основано на номере заключения)...")
    type_text(new_project_name)
    wait(1)
    for _ in range(2):
        pyautogui.press('tab')
        wait(0.2)
    pyautogui.press('enter')
    wait(20)


def delete_current_object_simple():
    """Функция удаления текущего объекта."""
    print("[Удаление] Позиционирование для удаления...")
    
    # Убеждаемся, что мы в правильной позиции и не в режиме редактирования
    pyautogui.press('esc')  # Выходим из режима редактирования, если были в нем
    wait(0.5)
        
    # Обеспечиваем работу Alt
    print("[Удаление] Обеспечение работы Alt...")
    
    # Активация Display Plots
    pyautogui.press('right')
    wait(0.1)
    press_hotkey(['shift', 'left'])
    wait(0.1)
    
    press_hotkey(['ctrl', 'l'])
    wait(3)
    pyautogui.press('esc')
    wait(1)
    
    # Удаляем объект через меню
    print("[Удаление] Вызов меню удаления...")
    pyautogui.press('alt')
    wait(0.1)
    pyautogui.press('right')  # меню Edit
    wait(0.1)
    for _ in range(5):
        pyautogui.press('down')
        wait(0.1)
    pyautogui.press('enter')  # Delete Sample
    wait(3)
    
    print("[Удаление] Подтверждение удаления...")
    pyautogui.press('enter')
    wait(5)
    
    # Сброс выделения New Project
    pyautogui.click()
    wait(0.5)
    
    print("[Удаление] Объект удален, курсор автоматически перемещен на следующий")


# ----------------------------------------------------------------------
# Функции для 2-3 режимов (одно заключение - один объект)
# ----------------------------------------------------------------------
def process_modes_2_3():
    """Обработка режимов 2-3 с предварительным анализом проекта."""
    global processed_numbers
    current_year = datetime.now().strftime("%y")
    
    print(f"\n[Режим {MODE}] Запуск обработки с предварительным анализом...")
    
    # Этап 1: Анализ проекта и подготовка
    # Возвращаемся к началу проекта
    print("[Навигация] Возврат к началу проекта...")
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.5)
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.5)
    press_keys(['right', 'right'])  # Переходим к ячейке с номером
    wait(0.1)

    project_objects = scan_project_objects()
    if not project_objects:
        print("[Ошибка] Не удалось просканировать объекты проекта.")
        return False
        
    first_target_object = find_first_target_object(project_objects)
    if not first_target_object:
        print(f"[Режим {MODE}] В проекте нет объектов для обработки.")
        return False
    
    pyautogui.press('esc')
    press_keys(['left', 'left'])  # Возвращаемся к началу строки
    wait(0.1)

    # Подготовка проекта (создание правильно названного, удаление ненужных)
    prepare_project_for_modes_2_3(project_objects, first_target_object, current_year)
    
    # После подготовки проект изменился: некоторые объекты удалены, активен первый целевой.
    # Повторно сканируем проект для получения актуального списка объектов.
    print("[Анализ] Повторное сканирование проекта после подготовки...")
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.5)
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.5)
    press_keys(['right', 'right'])
    wait(0.1)
    
    updated_objects = scan_project_objects()
    if not updated_objects:
        print("[Ошибка] Не удалось повторно просканировать объекты после подготовки.")
        return False
    
    pyautogui.press('esc')
    press_keys(['left', 'left'])
    wait(0.1)
    
    # Этап 2: Обработка подготовленного списка объектов
    print(f"[Режим {MODE}] Переход к обработке подготовленного проекта...")
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.5)
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.5)

    process_objects_from_list(updated_objects)
    
    return True


def scan_project_objects():
    """Сканирует все объекты в текущем проекте и возвращает список их названий."""
    print("[Анализ] Сканирование объектов проекта...")
    objects_list = []
    row_index = 0
    
    while True:
        # Читаем текущий объект
        pyautogui.press('enter')
        wait(0.5)
        
        # Выделяем и копируем
        for _ in range(4):
            press_hotkey(['ctrl', 'shift', 'left'])
            wait(0.1)
        press_hotkey(['ctrl', 'c'])
        wait(0.5)
        object_name_raw = get_clipboard_text()
        pyautogui.press('esc')
        wait(0.5)
        
        object_name = clean_case_number_suffix(object_name_raw) if object_name_raw else ""
        
        print(f"[Анализ] Строка {row_index + 1}: '{object_name_raw}' -> '{object_name}'")
        
        # Проверяем маркеры завершения
        if not object_name or object_name in ("AL", "Al", "K+", "K-"):
            print(f"[Анализ] Обнаружен маркер завершения или пустота: '{object_name}'. Сканирование завершено.")
            break
            
        objects_list.append({
            'raw_name': object_name_raw,
            'clean_name': object_name,
            'row_index': row_index
        })
        
        # Переход к следующей строке
        pyautogui.press('down')
        wait(0.5)
        row_index += 1
    
    print(f"[Анализ] Найдено объектов: {len(objects_list)}")
    for obj in objects_list:
        print(f"  - {obj['clean_name']}")
    
    return objects_list


def find_first_target_object(objects_list):
    """Находит первый объект, который нужно обработать согласно режиму."""
    print(f"[Анализ] Поиск первого целевого объекта для режима '{MODE}'...")
    
    for obj in objects_list:
        object_name = obj['clean_name']
        should_process = False
        
        if MODE == "selective":
            should_process = object_name in SELECTIVE_LIST
        elif MODE == "exclude":
            should_process = object_name not in EXCLUDE_LIST
            
        if should_process:
            print(f"[Анализ] Первый целевой объект: '{object_name}' (строка {obj['row_index'] + 1})")
            return obj
    
    print(f"[Анализ] Целевых объектов не найдено")
    return None


def prepare_project_for_modes_2_3(objects_list, first_target_object, current_year):
    """Подготавливает проект для обработки: создает правильно названный проект и удаляет ненужные объекты."""
    first_target_name = first_target_object['clean_name']
    first_target_row = first_target_object['row_index']
    
    print(f"[Подготовка] Первый целевой объект: '{first_target_name}' на строке {first_target_row + 1}")
    
    # Если первый объект в проекте НЕ является первым целевым, нужно создать новый проект
    if first_target_row > 0:
        print(f"[Подготовка] Первый объект проекта НЕ является целевым. Создание нового проекта...")
        
        # Возвращаемся к началу проекта для создания нового
        return_to_project_start()
        
        # Создаем новый проект с именем первого целевого объекта
        create_new_project_with_name(first_target_name, current_year)
        
        # Удаляем все объекты до первого целевого
        delete_objects_before_target(first_target_row)
        
    else:
        print(f"[Подготовка] Первый объект проекта уже является целевым. Создание нового проекта не требуется.")
        return_to_project_start()


def return_to_project_start():
    """Возвращается к началу проекта."""
    print("[Навигация] Возврат к началу проекта...")
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.5)
    press_hotkey(['ctrl', 'shift', 'pageup'])
    wait(0.5)


def create_new_project_with_name(object_name, current_year):
    """Создает новый проект с именем, основанным на номере заключения объекта."""
    base_number = extract_case_number(object_name)
    if not base_number:
        print(f"[Проект] Не удалось извлечь номер заключения из '{object_name}'. Использую полное имя.")
        base_number = object_name
    
    new_project_name = f"{base_number}-{current_year}"
    print(f"[Проект] Создание нового проекта '{new_project_name}' для объекта '{object_name}'...")
    
    # Активация Display Plots
    pyautogui.press('right')
    wait(0.1)
    press_hotkey(['shift', 'left'])
    wait(0.5)
    
    # Обеспечение работы Alt
    press_hotkey(['ctrl', 'l'])
    wait(3)
    pyautogui.press('esc')
    wait(1)
    
    # Создание нового проекта через "Сохранить как"
    pyautogui.press('alt')
    wait(0.5)
    for _ in range(4):
        pyautogui.press('down')
        wait(0.1)
    pyautogui.press('enter')
    wait(3)
    
    # Ввод имени проекта
    type_text(new_project_name)
    wait(1)
    for _ in range(2):
        pyautogui.press('tab')
        wait(0.2)
    pyautogui.press('enter')
    wait(15)
    
    print(f"[Проект] Новый проект '{new_project_name}' создан")


def create_project_for_case(case_number, current_year):
    """
    Создаёт новый проект с именем, основанным на номере заключения объекта case_number.
    Не выполняет лишней навигации, предполагается, что фокус уже находится в главном окне GeneMapper.
    """
    base_number = extract_case_number(case_number)
    if not base_number:
        print(f"[Проект] Ошибка: не удалось извлечь номер заключения из '{case_number}'. Использую полное имя.")
        base_number = case_number

    new_project_name = f"{base_number}-{current_year}"
    print(f"[Проект] Создание нового проекта '{new_project_name}' для объекта '{case_number}'")

    # Обеспечение работы Alt (как в других функциях)
    try:
        press_hotkey(['ctrl', 'l'])
        time.sleep(3)
        pyautogui.press('esc')
        time.sleep(1)

        # Открыть меню "Сохранить как"
        pyautogui.press('alt')
        time.sleep(0.5)
        for _ in range(4):
            pyautogui.press('down')
            time.sleep(0.1)
        pyautogui.press('enter')
        time.sleep(3)

        # Ввести имя нового проекта
        type_text(new_project_name)
        time.sleep(1)
        for _ in range(2):
            pyautogui.press('tab')
            time.sleep(0.2)
        pyautogui.press('enter')
        time.sleep(15)   # Ожидание завершения сохранения

        print(f"[Проект] Новый проект '{new_project_name}' успешно создан.")
    except Exception as e:
        print(f"[Проект] КРИТИЧЕСКАЯ ОШИБКА при создании проекта '{new_project_name}': {e}")
        raise   # Прерываем выполнение, так как дальнейшая работа невозможна

    # Дополнительно нажимаем Esc для снятия возможных выделений
    pyautogui.press('esc')
    time.sleep(0.5)


def delete_objects_before_target(target_row_index):
    """Удаляет все объекты до целевой строки."""
    print(f"[Очистка] Удаление объектов до строки {target_row_index + 1}...")
    
    # Возвращаемся к началу проекта
    return_to_project_start()
    
    # Удаляем объекты с начала до целевого
    for i in range(target_row_index):
        print(f"[Очистка] Удаление объекта на строке {i + 1}...")
        
        # Активация Display Plots
        pyautogui.press('right')
        wait(0.1)
        press_hotkey(['shift', 'left'])
        wait(0.5)
        
        # Обеспечение работы Alt
        press_hotkey(['ctrl', 'l'])
        wait(3)
        pyautogui.press('esc')
        wait(1)
        
        # Удаление объекта
        pyautogui.press('alt')
        wait(0.1)
        pyautogui.press('right')
        wait(0.1)
        for _ in range(5):
            pyautogui.press('down')
            wait(0.1)
        pyautogui.press('enter')
        wait(3)
        pyautogui.press('enter')
        wait(5)
        
        # Сброс фокуса
        pyautogui.click()
        wait(0.5)
        
        print(f"[Очистка] Объект на строке {i + 1} удален (курсор автоматически перешел на следующий)")
    
    print(f"[Очистка] Удалено {target_row_index} объектов. Целевой объект теперь первый в проекте.")


def process_objects_from_list(objects_list):
    """Обрабатывает объекты по заранее подготовленному списку.
    Для каждого объекта:
      - проверяет необходимость сохранения согласно режиму (selective/exclude)
      - если нужно сохранить, перед сохранением создаёт новый проект, если номер заключения изменился
      - выполняет стандартные операции сохранения
      - удаляет текущий объект (или пропущенный) через GUI
    """
    global processed_numbers, MODE, SELECTIVE_LIST, EXCLUDE_LIST
    current_year = datetime.now().strftime("%y")

    idx = 0
    total = len(objects_list)
    print(f"[Обработка] Начинаем обработку {total} объектов из списка.")

    # НОВОЕ: отслеживание последнего сохранённого номера заключения
    last_saved_base = None

    while idx < total:
        obj = objects_list[idx]
        case_number = obj['clean_name']
        row_index = obj['row_index']

        print(f"\n[Обработка] Объект {idx+1}/{total}: '{case_number}' (строка {row_index+1})")

        # Проверка маркеров завершения
        if case_number in ("AL", "Al", "K+", "K-"):
            print(f"[Завершение] Обнаружен маркер: '{case_number}'. Прерывание обработки.")
            break

        # Проверка дубликата
        base_num = extract_case_number(case_number)
        if base_num in processed_numbers:
            print(f"[Дубликат] Номер '{base_num}' уже обработан в этом сеансе. Удаляем текущий объект и продолжаем.")
            delete_current_object_simple()
            idx += 1
            continue
        else:
            processed_numbers.add(base_num)

        # Определяем, нужно ли сохранять текущий объект
        should_save = True
        if MODE == "selective":
            should_save = case_number in SELECTIVE_LIST
        elif MODE == "exclude":
            should_save = case_number not in EXCLUDE_LIST

        print(f"[Режим {MODE}] Объект '{case_number}': решение сохранить = {should_save}")

        if should_save:
            # НОВОЕ: проверяем, нужно ли создать новый проект
            if last_saved_base is None:
                # Первый сохраняемый объект – проект уже должен быть создан на этапе подготовки
                print(f"[Проект] Первый целевой объект '{case_number}'. Проект уже подготовлен (номер '{base_num}').")
                last_saved_base = base_num
            elif base_num != last_saved_base:
                # Номер изменился – создаём новый проект
                print(f"[Проект] Обнаружена смена номера заключения: '{last_saved_base}' -> '{base_num}'. Создаём новый проект...")
                create_project_for_case(case_number, current_year)
                last_saved_base = base_num
            else:
                print(f"[Проект] Номер заключения '{base_num}' совпадает с предыдущим. Проект не меняется.")

            # Выполняем стандартные операции сохранения
            press_keys(['right', 'right'])
            perform_save_operations(case_number, current_year)
            print(f"[Сохранение] Операции сохранения для '{case_number}' завершены.")
        else:
            print(f"[Пропуск] Объект '{case_number}' пропущен согласно режиму.")

        # Удаляем текущий объект (обработанный или пропущенный)
        pyautogui.press('esc')
        wait(0.2)
        delete_current_object_simple()
        print(f"[Удаление] Объект '{case_number}' удалён из проекта.")

        idx += 1
        time.sleep(1)

    print("[Обработка] Все объекты из списка обработаны.")


def process_one_sample_standard():
    """Стандартная обработка одного образца (как в режиме 1)."""
    global processed_numbers
    current_year = datetime.now().strftime("%y")

    # Получение номера заключения
    print("Перемещение к ячейке с номером заключения...")
    press_keys(['right', 'right'])
    pyautogui.press('enter')
    wait(0.5)

    print("Выделение номера заключения...")
    for _ in range(4):
        press_hotkey(['ctrl', 'shift', 'left'])
        wait(0.1)

    press_hotkey(['ctrl', 'c'])
    wait(0.5)
    pyautogui.press('enter')

    case_number_raw = get_clipboard_text()
    if not case_number_raw:
        print("[Ошибка] Не удалось скопировать номер заключения. Прерывание.")
        return False
    
    case_number = clean_case_number_suffix(case_number_raw)
    print(f"[Обработка] Текущий объект: '{case_number_raw}' -> Очищенный: '{case_number}'")

    # Проверка маркеров завершения
    if case_number in ("AL", "Al", "K+", "K-"):
        print(f"[Завершение] Обнаружен маркер: '{case_number}'. Программа завершается.")
        return False

    # Проверка дубликата
    base_num = extract_case_number(case_number)
    if base_num in processed_numbers:
        print(f"[Дубликат] Номер '{base_num}' уже обработан в этом сеансе.")
        input("Нажмите Enter, чтобы пропустить этот объект и продолжить...")
        delete_current_object_simple()
        return True
    else:
        processed_numbers.add(base_num)

    # Проверяем, нужно ли обрабатывать этот объект
    should_save = True
    if MODE == "selective":
        should_save = case_number in SELECTIVE_LIST
    elif MODE == "exclude":
        should_save = case_number not in EXCLUDE_LIST

    print(f"[Режим {MODE}] Объект '{case_number}': решение сохранить = {should_save}")

    if should_save:
        # Выполняем операции сохранения
        perform_save_operations(case_number, current_year)
        
        # Удаляем обработанный объект
        print(f"[Навигация] Возврат к началу строки...")
        press_keys(['left', 'left'])
        
        delete_current_object_simple()
        print(f"[Сохранение] Объект '{case_number}' обработан и удален")
        
    else:
        # Удаляем ненужный объект
        print(f"[Навигация] Возврат к началу строки...")
        press_keys(['left', 'left'])
        
        delete_current_object_simple()
        print(f"[Пропуск] Объект '{case_number}' удален")

    return True


def create_new_project_for_next_object(next_case, current_year):
    """Создает новый проект для следующего объекта в режимах 2-3."""
    print(f"[Режим {MODE}] Создание нового проекта для следующего объекта '{next_case}'...")
    
    # Активация Display Plots
    pyautogui.press('right')
    wait(0.1)
    press_hotkey(['shift', 'left'])
    wait(0.5)
    
    # Обеспечение работы Alt
    print(f"[Режим {MODE}] Обеспечение работы Alt...")
    press_hotkey(['ctrl', 'l'])
    wait(3)
    pyautogui.press('esc')
    wait(1)

    # Извлекаем номер заключения для имени проекта
    base_number = extract_case_number(next_case)
    if not base_number:
        print(f"[Проект] Не удалось извлечь номер заключения из '{next_case}'. Использую полное имя.")
        base_number = next_case
    else:
        print(f"[Проект] Из полного имени '{next_case}' извлечен номер заключения '{base_number}'")
    
    new_project_name = f"{base_number}-{current_year}"
    print(f"[Режим {MODE}] Создание проекта с именем '{new_project_name}'...")

    # Создание нового проекта через меню
    pyautogui.press('alt')
    wait(0.5)
    for _ in range(4):
        pyautogui.press('down')
        wait(0.1)
    pyautogui.press('enter')
    wait(3)

    # Ввод имени проекта
    type_text(new_project_name)
    wait(1)
    for _ in range(2):
        pyautogui.press('tab')
        wait(0.2)
    pyautogui.press('enter')
    wait(15)
    
    print(f"[Режим {MODE}] Новый проект '{new_project_name}' создан")

    # Убеждаемся в правильной позиции для следующей итерации
    press_keys(['right', 'right'])
    wait(0.5)


# ----------------------------------------------------------------------
# Функции для 4 режима (одно заключение - несколько объектов)
# ----------------------------------------------------------------------
def process_mode4():
    """Обработка режима с несколькими объектами одного номера."""
    global MULTI_BASE_NUMBER, MULTI_SAVED_CONTROLS, MULTI_TXT_SAVED
    print("\n[Режим multi] Запуск обработки множественных объектов одного номера.")
    MULTI_SAVED_CONTROLS.clear()
    MULTI_TXT_SAVED = False
    MULTI_BASE_NUMBER = None

    row_index = 0
    while True:
        print(f"\n[Режим multi] Обработка строки {row_index+1}...")
        # Чтение текущей ячейки
        case_number = get_current_cell_text()
        if not case_number:
            print("[Режим multi] Не удалось прочитать ячейку. Завершение.")
            break
        print(f"[Режим multi] Текущий образец: '{case_number}'")

        # Проверка маркеров окончания
        if case_number in ("AL", "Al"):
            print("[Режим multi] Обнаружен AL/Al. Завершение.")
            break

        # Обработка контроля K+ / K-
        if case_number in ("K+", "K-"):
            if case_number in MULTI_SAVED_CONTROLS:
                print(f"[Режим multi] Контроль '{case_number}' уже сохранён, пропускаем.")
            else:
                print(f"[Режим multi] Сохраняем контроль '{case_number}' (первый раз).")
                save_control_foregram(case_number)
                MULTI_SAVED_CONTROLS.add(case_number)
            # Переход к следующей строке
            pyautogui.press('down')
            time.sleep(0.5)
            row_index += 1
            continue

        # Обработка объекта (не контроль и не AL/Al)
        # Извлекаем номер заключения
        current_base = extract_case_number(case_number)
        if MULTI_BASE_NUMBER is None:
            MULTI_BASE_NUMBER = current_base
            print(f"[Режим multi] Установлен номер заключения: {MULTI_BASE_NUMBER}")
            # Сохраняем TXT один раз после определения номера
            save_txt_and_move_mode4(MULTI_BASE_NUMBER)
        else:
            if current_base != MULTI_BASE_NUMBER:
                print(f"[Режим multi] Номер заключения '{current_base}' не совпадает с первым '{MULTI_BASE_NUMBER}'. Завершение.")
                break

        # Проверка, нужно ли исключить этот объект
        if MULTI_EXCLUDE_LIST and case_number in MULTI_EXCLUDE_LIST:
            print(f"[Режим multi] Объект '{case_number}' в списке исключения. Пропускаем.")
        else:
            print(f"[Режим multi] Сохраняем фореграмму объекта '{case_number}'.")
            save_sample_foregram()

        # Переход к следующей строке
        pyautogui.press('down')
        time.sleep(0.5)
        row_index += 1

    print("[Режим multi] Обработка завершена.")


def save_sample_foregram():
    """Сохраняет фореграмму текущего образца (объекта или контроля без настройки Y)."""
    # Возврат в начало строки и активация Display Plots
    pyautogui.press('left')
    press_hotkey(['shift', 'left'])
    time.sleep(0.2)

    # Открыть диалог печати/сохранения
    press_hotkey(['ctrl', 'l'])
    time.sleep(3)
    pyautogui.press('alt')
    for _ in range(3):
        pyautogui.press('down')
        time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.press('esc')
    time.sleep(1)

def save_kminus_foregram():
    """Сохраняет фореграмму K- с предварительной настройкой оси Y."""
    # Выделение образца
    pyautogui.press('left')
    press_hotkey(['shift', 'left'])
    time.sleep(0.2)

    # Настройка оси Y
    press_hotkey(['ctrl', 'l'])
    time.sleep(3)
    pyautogui.press('alt')
    for _ in range(3):
        pyautogui.press('right')
        time.sleep(0.1)
    pyautogui.press('down')
    time.sleep(0.1)
    pyautogui.press('right')
    time.sleep(0.1)
    for _ in range(3):
        pyautogui.press('down')
        time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(1)

    for _ in range(2):
        pyautogui.press('tab')
        time.sleep(0.1)
    for _ in range(6):
        pyautogui.press('backspace')
        time.sleep(0.1)
    type_text("400")
    pyautogui.press('tab')
    time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(1)

    # Сохранение фореграммы
    pyautogui.press('alt')
    for _ in range(3):
        pyautogui.press('down')
        time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(2)
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.press('esc')
    time.sleep(1)

def save_control_foregram(control_type):
    """Сохраняет фореграмму контроля (control_type = 'K+' или 'K-')."""
    if control_type == 'K+':
        save_sample_foregram()
    elif control_type == 'K-':
        save_kminus_foregram()
    else:
        print(f"[Ошибка] Неизвестный тип контроля: {control_type}")


def move_txt_file_mode4(base_number):
    """Перемещает файл Genotypes Table.txt из SOURCE_FOLDER или TEMP_FOLDER в папку DEST_BASE/base_number."""
    dest_path = os.path.join(DEST_BASE, base_number)
    os.makedirs(dest_path, exist_ok=True)

    pattern = re.compile(rf"^{re.escape(base_number)} Genotypes Table.*\.txt$", re.IGNORECASE)

    # Поиск в SOURCE_FOLDER
    for folder in [SOURCE_FOLDER, TEMP_FOLDER]:
        if not os.path.exists(folder):
            continue
        try:
            for filename in os.listdir(folder):
                if pattern.match(filename):
                    src = os.path.join(folder, filename)
                    dst = os.path.join(dest_path, filename)
                    shutil.move(src, dst)
                    print(f"[Режим multi] Перемещён TXT: {filename} -> {dest_path}")
                    return True
        except Exception as e:
            print(f"[Ошибка] При поиске в {folder}: {e}")
    print(f"[Режим multi] Файл Genotypes Table для {base_number} не найден.")
    return False

def save_txt_and_move_mode4(base_number):
    """Сохраняет Genotypes Table.txt (один раз) и перемещает в папку с номером."""
    global MULTI_TXT_SAVED
    if MULTI_TXT_SAVED:
        return
    print("[Режим multi] Сохранение Genotypes Table.txt (один раз)...")
    press_hotkey(['ctrl', 'shift', '3'])
    time.sleep(1)
    press_hotkey(['ctrl', 'e'])
    time.sleep(2)
    for _ in range(2):
        pyautogui.press('tab')
        time.sleep(0.1)
    pyautogui.press('enter')
    time.sleep(3)
    press_hotkey(['ctrl', 'shift', '1'])
    if move_txt_file_mode4(base_number):
        MULTI_TXT_SAVED = True
        print("[Режим multi] TXT сохранён и перемещён.")
    else:
        print("[Режим multi] Не удалось переместить TXT-файл.")


# ----------------------------------------------------------------------
# Функции для работы с GeneMapper
# ----------------------------------------------------------------------
def select_mode():
    """Выбор режима работы программы."""
    global MODE, SELECTIVE_LIST, EXCLUDE_LIST, MULTI_EXCLUDE_LIST
    print("\nВыберите режим работы:")
    print("1. Тотальное сохранение (ВСЕ профили)")
    print("2. Избирательное сохранение (сохранять только УКАЗАННЫЕ профили)")
    print("3. Сохранение с исключением (сохранять все, КРОМЕ УКАЗАННЫХ)")
    print("4. Множественные объекты одного номера (сохранить все объекты и K+/K- один раз, без создания новых проектов)")
    while True:
        choice = input("Введите номер режима (1-4): ").strip()
        if choice == "1":
            MODE = "full"
            print("Выбран режим: Тотальное сохранение")
            break
        elif choice == "2":
            MODE = "selective"
            items = input("Введите названия профилей для сохранения через запятую (полностью, как в GeneMapper - например, 123 2B-26): ").strip()
            SELECTIVE_LIST = [item.strip() for item in items.split(',') if item.strip()] if items else []
            print(f"Выбран режим: Избирательное сохранение. Список для сохранения: {SELECTIVE_LIST}")
            break
        elif choice == "3":
            MODE = "exclude"
            items = input("Введите названия профилей для исключения через запятую (полностью, как в GeneMapper - например, 123 2B-26)): ").strip()
            EXCLUDE_LIST = [item.strip() for item in items.split(',') if item.strip()] if items else []
            print(f"Выбран режим: Сохранение с исключением. Список исключений: {EXCLUDE_LIST}")
            break
        elif choice == "9":
            MODE = "multi"
            print("Выбран режим: множественные объекты одного заключения")
            sub = input("Сохранить все объекты? (д/н): ").strip().lower()
            if sub == 'д':
                MULTI_EXCLUDE_LIST = []
            else:
                items = input("Введите полные названия объектов для исключения через запятую (например, '1234 3A, 1234 3A-2'): ").strip()
                MULTI_EXCLUDE_LIST = [item.strip() for item in items.split(',') if item.strip()] if items else []
            print(f"Список исключения для режима multi: {MULTI_EXCLUDE_LIST}")
            break
        else:
            print("Неверный ввод. Пожалуйста, введите 1, 2, 3 или 4.")


def extract_case_number(full_name):
    """Извлекает номер заключения из полного имени объекта.
    Возвращает строку до первого пробела, если пробел есть, иначе исходную строку.
    """
    if not full_name:
        return ""
    parts = full_name.split(' ', 1)
    return parts[0]


def clean_case_number_suffix(full_name):
    """Удаляет суффикс вида '-ГГ' (год) из названия объекта, если он есть.
    Например: '2092-26' -> '2092', '2093 1A-26' -> '2093 1A'
    """
    if not full_name:
        return full_name
    
    # Ищем паттерн: дефис + 2 цифры в конце строки
    pattern = r'-\d{2}$'
    cleaned = re.sub(pattern, '', full_name)
    
    if cleaned != full_name:
        print(f"[Очистка суффикса] '{full_name}' -> '{cleaned}'")
    
    return cleaned


def get_current_cell_text():
    """Возвращает текст из текущей ячейки (выделенной) в GeneMapper.
       После вызова курсор остаётся в ячейке, режим редактирования закрыт."""
    # Перемещение к ячейке с номером
    pyautogui.press('right')
    pyautogui.press('right')
    time.sleep(0.2)
    pyautogui.press('enter')
    time.sleep(0.5)
    # Выделение текста
    for _ in range(4):
        press_hotkey(['ctrl', 'shift', 'left'])
        time.sleep(0.5)
    press_hotkey(['ctrl', 'c'])
    time.sleep(0.5)
    text = get_clipboard_text()
    pyautogui.press('esc')
    time.sleep(0.2)
    return text


def move_files(case_number, year_suffix, base_number=None):
    """Переносит файлы в папку проекта и проверяет очистку временной папки.
    case_number - полное имя объекта (используется для логирования и совместимости)
    year_suffix - две последние цифры года
    base_number - номер заключения (часть до пробела), если не указан, вычисляется из case_number
    """
    if base_number is None:
        base_number = extract_case_number(case_number)
        print(f"[Перемещение файлов] base_number не указан, вычислен из case_number '{case_number}' -> '{base_number}'")

    # Имя папки: номер-год
    folder_name = f"{base_number}-{year_suffix}"
    dest_path = os.path.join(DEST_BASE, folder_name)

    # Создаём папку, если её нет
    try:
        os.makedirs(dest_path, exist_ok=True)
        print(f"Папка назначения: {dest_path} (основана на номере заключения '{base_number}')")
    except Exception as e:
        print(f"[Ошибка] Создание папки {dest_path}: {e}")
        return

    # Поиск и перемещение файла Genotypes Table
    # Шаблон: начинается с номера заключения, затем любые символы, затем дефис+год, пробел, "Genotypes Table", .txt
    pattern = re.compile(rf"^{re.escape(base_number)}.*-{re.escape(year_suffix)} Genotypes Table.*\.txt$", re.IGNORECASE)
    found = False

    # ---- 1. Поиск в папке Client (SOURCE_FOLDER) ----
    try:
        for filename in os.listdir(SOURCE_FOLDER):
            if pattern.match(filename):
                src = os.path.join(SOURCE_FOLDER, filename)
                dst = os.path.join(dest_path, filename)
                shutil.move(src, dst)
                print(f"[OK] Перемещён файл Genotypes Table из Client: {filename}")
                found = True
                break
    except Exception as e:
        print(f"[Ошибка] При работе с папкой {SOURCE_FOLDER}: {e}")

    # ---- 2. Если не нашли в Client, ищем во временной папке ----
    if not found:
        print(f"[Инфо] Файл Genotypes Table для {base_number}-{year_suffix} НЕ НАЙДЕН в {SOURCE_FOLDER}. Проверяем временную папку {TEMP_FOLDER}...")
        if os.path.exists(TEMP_FOLDER):
            try:
                for filename in os.listdir(TEMP_FOLDER):
                    if pattern.match(filename):
                        print(f"[Инфо] Файл Genotypes Table НАЙДЕН во временной папке: {filename}. Он будет перемещён при копировании всего содержимого {TEMP_FOLDER}.")
                        found = True
                        break
                if not found:
                    print(f"[Предупреждение] Файл Genotypes Table для номера {base_number}-{year_suffix} не найден НИ В {SOURCE_FOLDER}, НИ В {TEMP_FOLDER}.")
            except Exception as e:
                print(f"[Ошибка] При поиске во временной папке {TEMP_FOLDER}: {e}")
        else:
            print(f"[Предупреждение] Временная папка {TEMP_FOLDER} не существует, файл Genotypes Table не найден.")

    # ---- 3. Логирование содержимого временной папки до перемещения ----
    temp_contents_before = []
    if os.path.exists(TEMP_FOLDER):
        try:
            temp_contents_before = os.listdir(TEMP_FOLDER)
            print(f"[Временная папка] Перед перемещением содержит {len(temp_contents_before)} элементов: {temp_contents_before}")
        except Exception as e:
            print(f"[Ошибка] Не удалось прочитать содержимое временной папки {TEMP_FOLDER}: {e}")

    # ---- 4. Перемещение всего содержимого временной папки (включая возможный файл Genotypes Table) ----
    if os.path.exists(TEMP_FOLDER):
        try:
            for item in os.listdir(TEMP_FOLDER):
                src = os.path.join(TEMP_FOLDER, item)
                dst = os.path.join(dest_path, item)
                shutil.move(src, dst)
                print(f"Перемещён {item} из временной папки")
        except Exception as e:
            print(f"[Ошибка] При перемещении из временной папки: {e}")
    else:
        print(f"[Предупреждение] Временная папка {TEMP_FOLDER} не найдена")

    # ---- 5. Проверка, что временная папка пуста после перемещения ----
    if os.path.exists(TEMP_FOLDER):
        try:
            remaining = os.listdir(TEMP_FOLDER)
            if not remaining:
                print(f"[Временная папка] Папка {TEMP_FOLDER} полностью очищена (пуста).")
            else:
                print(f"[ВНИМАНИЕ] После перемещения в папке {TEMP_FOLDER} остались следующие элементы ({len(remaining)}): {remaining}")
        except Exception as e:
            print(f"[Ошибка] Не удалось проверить содержимое временной папки {TEMP_FOLDER}: {e}")


# ----------------------------------------------------------------------
# Главный цикл
# ----------------------------------------------------------------------
def main():
    original_numlock_state = None
    restored = False
    try:
        print("=" * 60)
        print("Автоматизация GeneMapper")
        print("Для остановки переместите мышь в левый верхний угол.")
        print("=" * 60)
        
        # Управление NumLock
        original_numlock_state = get_numlock_state()
        print(f"[NumLock] Исходное состояние: {'ВКЛЮЧЕН' if original_numlock_state else 'ВЫКЛЮЧЕН'}")
        
        if original_numlock_state:
            print("[NumLock] Обнаружен включенный NumLock. Программа требует выключенного NumLock. Выключаю...")
            if not set_numlock_state(False):
                print("[NumLock] НЕ УДАЛОСЬ выключить NumLock. Возможны проблемы с навигацией.")
            else:
                print("[NumLock] NumLock успешно выключен. Продолжаем работу.")
        else:
            print("[NumLock] NumLock уже выключен. Продолжаем без изменений.")

        # Выбор режима
        select_mode()

        if MODE == "multi":
            # ... (режим 4 остается тем же)
            process_mode4()
            return

        # Для режимов 1-3 
        print("Выполнение программы начнётся через:")
        for i in range(5, 0, -1):
            print(f"\n{i}...", end=" ", flush=True)
            time.sleep(1)
        print("\nСтарт!")

        if MODE == "full":
            # Режим 1 - стандартный цикл
            iteration = 0
            while True:
                iteration += 1
                print(f"\n--- Итерация {iteration} ---")
                if not process_one_sample():  # старая функция для режима 1
                    print("[Завершение] Обработка завершена.")
                    break
                print(f"--- Итерация {iteration} завершена ---")
                time.sleep(1)
        else:
            # Режимы 2-3 - логика с предварительным анализом
            process_modes_2_3()
            
    except KeyboardInterrupt:
        print("\n[Остановка] Программа остановлена пользователем.")
    except Exception as e:
        print(f"\n[Критическая ошибка] {e}")
        import traceback
        traceback.print_exc()
    finally:
        # Восстанавливаем исходное состояние NumLock
        if original_numlock_state is not None and not restored:
            if get_numlock_state() != original_numlock_state:
                print(f"[NumLock] Восстанавливаем исходное состояние: {'ВКЛЮЧЕН' if original_numlock_state else 'ВЫКЛЮЧЕН'}")
                set_numlock_state(original_numlock_state)
            else:
                print(f"[NumLock] Состояние NumLock уже соответствует исходному, восстановление не требуется.")
            restored = True


if __name__ == "__main__":
    main()