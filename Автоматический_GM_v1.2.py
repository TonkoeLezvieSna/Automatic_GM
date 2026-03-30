import sys
import os
import time
import shutil
import re
import win32clipboard
import pyautogui
from datetime import datetime


# ----------------------------------------------------------------------
# Настройки
# ----------------------------------------------------------------------
# Настройки pyautogui
pyautogui.FAILSAFE = True  # перемещение мыши в левый верхний угол останавливает скрипт
pyautogui.PAUSE = 0.6     # пауза между командами pyautogui

# Пути к папкам (можно изменить при необходимости)
SOURCE_FOLDER = r"C:\1\AppliedBiosystems\GeneMapperID-X\Client"
DEST_BASE = r"C:\GM_RESULTS"
TEMP_FOLDER = r"C:\GM_TEMP"


# ----------------------------------------------------------------------
# Глобальные переменные для режимов работы
# ----------------------------------------------------------------------
MODE = "full"          # "full", "selective", "exclude"
SELECTIVE_LIST = []    # список для избирательного сохранения
EXCLUDE_LIST = []      # список для исключения


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
# Основные действия для одного образца
# ----------------------------------------------------------------------
def process_one_sample():
    """Выполняет одну итерацию цикла: обработка одного номера заключения."""
    current_year = datetime.now().strftime("%y")  # две последние цифры года

    # ------------------------------------------------------------------
    # 1. Получение номера заключения
    # ------------------------------------------------------------------
    print("Перемещение к ячейке с номером заключения...")
    press_keys(['right', 'right'])
    pyautogui.press('enter')
    wait(0.5)

    print("Выделение номера заключения...")
    press_hotkey(['ctrl', 'shift', 'left'])
    wait(0.5)

    press_hotkey(['ctrl', 'c'])
    wait(0.5)

    case_number = get_clipboard_text()
    if not case_number:
        print("[Ошибка] Не удалось скопировать номер заключения. Прерывание.")
        return False
    print(f"Номер заключения: '{case_number}'")

    # Проверка маркеров завершения
    if case_number in ("AL", "Al", "K+", "K-"):
        print(f"Обнаружен маркер завершения: '{case_number}'. Программа завершается.")
        return False

    # ------------------------------------------------------------------
    # 2. Решение о сохранении текущего образца
    # ------------------------------------------------------------------
    should_save = True
    if MODE == "selective":
        should_save = case_number in SELECTIVE_LIST
    elif MODE == "exclude":
        should_save = case_number not in EXCLUDE_LIST
    print(f"[Режим {MODE}] Образец '{case_number}': решение сохранить = {should_save}")

    # ------------------------------------------------------------------
    # 3. Если сохраняем, выполняем все операции сохранения
    # ------------------------------------------------------------------
    if should_save:
        # Добавляем год и фиксируем
        print("Добавление года...")
        press_hotkey(['ctrl', 'right'])
        type_text(f"-{current_year}")
        pyautogui.press('enter')
        wait(0.5)

        # Возврат в начальную точку
        pyautogui.press(['left'])
        press_hotkey(['shift', 'left'])

        # Фореграмма (объект)
        print("Открытие и сохранение фореграммы объекта...")
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
        print("Переход на фореграмму К+...")
        pyautogui.press('pagedown')
        wait(0.5)
        pyautogui.press('pagedown')
        wait(0.5)
        press_hotkey(['shift', 'up'])

        print("Сохранение фореграммы К+...")
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
        print("Переход на фореграмму К-...")
        press_hotkey(['shift', 'down'])

        # Установка высоты оси Y
        print("Установка высоты фореграммы по оси Y...")
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
        for _ in range(6):
            pyautogui.press('backspace')
            wait(0.1)

        type_text("400")
        pyautogui.press('tab')
        wait(0.1)
        pyautogui.press('enter')
        wait(1)

        print("Сохранение фореграммы К-...")
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

        # Возврат в начало таблицы и сохранение Genotypes Table.txt
        print("Возврат в начало таблицы...")
        pyautogui.press('pageup')
        wait(0.2)
        pyautogui.press('pageup')
        wait(0.2)
        press_hotkey(['shift', 'down'])
        press_hotkey(['shift', 'up'])

        print("Сохранение Genotypes Table.txt...")
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

    else:
        # Выход из режима редактирования (если мы в нём) и пропуск образца
        pyautogui.press('esc')
        print(f"[Пропуск] Образец '{case_number}' пропущен по условию режима.")
        wait(0.5)
        for _ in range(2):
            pyautogui.press('left')
            wait(0.1)
        pyautogui.press('down')

    # ------------------------------------------------------------------
    # 4. Сохранение текущего проекта (всегда)
    # ------------------------------------------------------------------
    print("Сохранение текущего проекта...")
    press_hotkey(['ctrl', 'shift', '1'])
    wait(3)
    press_hotkey(['ctrl', 's'])
    wait(1)

    # ==================================================================
    # 5. Определение следующего объекта
    # ==================================================================
    print("Определение следующего объекта...")
    pyautogui.press('down')
    wait(0.5)
    for _ in range(2):
        pyautogui.press('right')
        wait(0.1)

    # Копируем номер следующего объекта
    pyautogui.press('enter')
    wait(0.5)
    press_hotkey(['ctrl', 'shift', 'left'])
    wait(0.5)
    press_hotkey(['ctrl', 'c'])
    wait(0.5)
    next_case = get_clipboard_text()
    pyautogui.press('esc')
    wait(0.5)

    if not next_case or next_case in ("AL", "K+", "K-") or next_case == case_number:
        print(f"Следующий объект: '{next_case}' -> это маркер, пустота или совпадает с текущим. Текущий объект последний.")
        is_last = True
    else:
        print(f"Следующий объект: '{next_case}'")
        is_last = False

    # Возврат на текущий объект (который мы только что обработали)
    print("Возврат на текущий объект...")
    for _ in range(2):
        pyautogui.press('left')
        wait(0.1)
    press_hotkey(['shift', 'up'])
    wait(0.1)

    # Обеспечение последующей работы Alt
    print("Обеспечение работы Alt...")
    press_hotkey(['ctrl', 'l'])
    wait(3)
    pyautogui.press('esc')
    wait(1)

    # ------------------------------------------------------------------
    # 6. Создание нового проекта (только если есть следующий номерной объект)
    # ------------------------------------------------------------------
    if not is_last:
        print(f"Создание нового проекта с именем {next_case}-{current_year}...")
        pyautogui.press('alt')
        wait(0.5)
        for _ in range(4):
            pyautogui.press('down')
            wait(0.1)
        pyautogui.press('enter')
        wait(3)

        # Используем next_case для имени нового проекта
        type_text(f"{next_case}-{current_year}")
        wait(1)
        for _ in range(2):
            pyautogui.press('tab')
            wait(0.2)
        pyautogui.press('enter')
        wait(10)

        # Удаление профиля
        print("Обеспечение работы Alt...")
        press_hotkey(['shift', 'down'])
        wait(0.1)
        press_hotkey(['shift', 'up'])
        wait(0.1)
        press_hotkey(['ctrl', 'l'])
        wait(3)
        pyautogui.press('esc')
        wait(1)

        print("Удаление профиля...")
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

        # Сброс выделения New Project
        pyautogui.click()

        # Обеспечение активности обрабатываемой ячейки (для следующего цикла)
        press_hotkey(['shift', 'down'])
        wait(0.1)
        press_hotkey(['shift', 'up'])
        wait(0.1)

    else:
        print("Последний номерной объект, новый проект не создаётся.")

    # Возвращаем True, если нужно продолжать (есть следующий объект)
    return not is_last


# ----------------------------------------------------------------------
# Дополнительные функции
# ----------------------------------------------------------------------
def select_mode():
    """Выбор режима работы программы."""
    global MODE, SELECTIVE_LIST, EXCLUDE_LIST
    print("\nВыберите режим работы:")
    print("1. Тотальное сохранение (ВСЕ профили, кроме маркеров AL, Al, K+, K-)")
    print("2. Избирательное сохранение (сохранять только УКАЗАННЫЕ профили)")
    print("3. Сохранение с исключением (сохранять все, КРОМЕ УКАЗАННЫХ и кроме маркеров AL, Al, K+, K-)")
    while True:
        choice = input("Введите номер режима (1-3): ").strip()
        if choice == "1":
            MODE = "full"
            print("Выбран режим: Тотальное сохранение")
            break
        elif choice == "2":
            MODE = "selective"
            items = input("Введите названия профилей для сохранения через запятую (полностью, как в GeneMapper): ").strip()
            if items:
                SELECTIVE_LIST = [item.strip() for item in items.split(',') if item.strip()]
            else:
                SELECTIVE_LIST = []
            print(f"Выбран режим: Избирательное сохранение. Список для сохранения: {SELECTIVE_LIST}")
            break
        elif choice == "3":
            MODE = "exclude"
            items = input("Введите названия профилей для исключения через запятую (полностью, как в GeneMapper): ").strip()
            if items:
                EXCLUDE_LIST = [item.strip() for item in items.split(',') if item.strip()]
            else:
                EXCLUDE_LIST = []
            print(f"Выбран режим: Сохранение с исключением. Список исключений: {EXCLUDE_LIST}")
            break
        else:
            print("Неверный ввод. Пожалуйста, введите 1, 2 или 3.")


def extract_case_number(full_name):
    """Извлекает номер заключения из полного имени объекта.
    Возвращает строку до первого пробела, если пробел есть, иначе исходную строку.
    """
    if not full_name:
        return ""
    parts = full_name.split(' ', 1)
    return parts[0]


def move_files(case_number, year_suffix, base_number=None):
    """Переносит файлы в папку проекта и проверяет очистку временной папки.
    case_number - полное имя объекта (используется для логирования и совместимости)
    year_suffix - две последние цифры года
    base_number - номер заключения (часть до пробела), если не указан, вычисляется из case_number
    """
    if base_number is None:
        base_number = extract_case_number(case_number)
        print(f"[move_files] base_number не указан, вычислен из case_number '{case_number}' -> '{base_number}'")

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
    try:
        for filename in os.listdir(SOURCE_FOLDER):
            if pattern.match(filename):
                src = os.path.join(SOURCE_FOLDER, filename)
                dst = os.path.join(dest_path, filename)
                shutil.move(src, dst)
                print(f"Перемещён файл: {filename}")
                found = True
                break
        if not found:
            print(f"[Предупреждение] Файл Genotypes Table для номера {base_number}-{year_suffix} (полное имя объекта: {case_number}) не найден в {SOURCE_FOLDER}")
    except Exception as e:
        print(f"[Ошибка] При работе с папкой {SOURCE_FOLDER}: {e}")

    # Логирование содержимого временной папки до перемещения
    temp_contents_before = []
    if os.path.exists(TEMP_FOLDER):
        try:
            temp_contents_before = os.listdir(TEMP_FOLDER)
            print(f"[Временная папка] Перед перемещением содержит {len(temp_contents_before)} элементов: {temp_contents_before}")
        except Exception as e:
            print(f"[Ошибка] Не удалось прочитать содержимое временной папки {TEMP_FOLDER}: {e}")

    # Перемещение содержимого временной папки
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

    # Проверка, что временная папка пуста после перемещения
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
    try:
        print("=" * 60)
        print("Автоматизация GeneMapper")
        print("Для остановки переместите мышь в левый верхний угол.")
        print("=" * 60)

        # Выбор режима
        select_mode()

        print("Выполнение программы начнётся через:")
        for i in range(5, 0, -1):
            print(f"\n{i}...", end=" ", flush=True)
            time.sleep(1)
        print("\nСтарт!")

        iteration = 0
        while True:
            iteration += 1
            print(f"\n--- Итерация {iteration} ---")
            if not process_one_sample():
                print("Обработка прервана (обнаружен маркер завершения или ошибка).")
                break
            print(f"--- Итерация {iteration} завершена ---")
            time.sleep(2)
    except KeyboardInterrupt:
        print("\nПрограмма остановлена пользователем.")
    except Exception as e:
        print(f"\n[Критическая ошибка] {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()