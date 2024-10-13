import re
from fuzzywuzzy import fuzz
import os
import functools
import datetime
import time
import traceback
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException

# Export
# - complete_match(name, text)
# - find_name_in_string(name, text)


def debug_decorator(func, record_errors=True, display_errors_in_full=True, display_errors_short=True, display_info=True):

    log_dir = 'errors'  # Папка для хранения логов ошибок
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)  # Создаем папку, если она не существует
    excel_file = 'errors/error_log.xlsx'  # Имя Excel файла для хранения логов ошибок

    @functools.wraps(func)
    def wrapper_debug(*args, **kwargs):
        while True:
            try:
                # Пытаемся открыть существующий файл или создать новый
                if os.path.exists(excel_file):
                    workbook = load_workbook(excel_file)
                    sheet = workbook.active
                else:
                    workbook = Workbook()
                    sheet = workbook.active
                    # Создаём заголовки для нового файла
                    sheet.append(
                        ["Timestamp", "Function Name", "Arguments", "Keyword Arguments", "Error Type", "Error Message",
                         "Traceback"])
                break  # Выходим из цикла, если файл доступен для записи

            except (PermissionError, InvalidFileException):
                print(f"Excel file '{excel_file}' is locked or in use. Retrying in 3 seconds...")
                time.sleep(3)

        try:
            result = func(*args, **kwargs)
            if display_info:
                print(f"--- DEBUGGING FUNCTION '{func.__name__}' ---")
                print(f"Arguments: {args}")
                print(f"Keyword Arguments: {kwargs}")
                print(f"Result: {result}")
                print(f"---------------------------\n")
            return result

        except Exception as e:
            error_traceback = traceback.format_exc()
            # Записываем ошибку в Excel файл, если `record_errors=True`
            if record_errors:
                error_time = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                # Собираем информацию об ошибке
                row = [
                    error_time,  # Время ошибки
                    func.__name__,  # Имя функции
                    str(args),  # Аргументы
                    str(kwargs),  # Ключевые аргументы
                    type(e).__name__,  # Тип ошибки
                    str(e),  # Сообщение об ошибке
                    error_traceback  # Полная трассировка
                ]

                # Добавляем строку в таблицу
                sheet.append(row)

                while True:
                    try:
                        # Пытаемся сохранить файл
                        workbook.save(excel_file)
                        break  # Выходим из цикла, если сохранение прошло успешно
                    except PermissionError:
                        print(f"Cannot save Excel file '{excel_file}' as it is locked. Retrying in 3 seconds...")
                        time.sleep(3)  # Ждем 3 секунды перед повторной попыткой

                print(f"Error in function '{func.__name__}' logged to {excel_file}")

            # Показываем ошибку в консоли, если `display_errors=True`
            if display_errors_short:
                print(f"Error in function '{func.__name__}'")

            elif display_errors_in_full:
                print(f"--- ERROR in function '{func.__name__}' ---")
                print(f"Arguments: {args}")
                print(f"Keyword Arguments: {kwargs}")
                print(f"Error Type: {type(e).__name__}")
                print(f"Error Message: {e}")
                print(f"Traceback: {error_traceback}")
                print(f"---------------------------\n")

    return wrapper_debug



def normalize_string(s):
    """Helper function to normalize strings: lowercases, removes extra spaces and common separators."""
    # Convert to lowercase
    s = s.lower()
    # Remove common separators and multiple spaces
    s = re.sub(r'[\s\-_,.]+', ' ', s).strip()
    return s

@debug_decorator
def find_name_in_string(name, text, threshold=80):
    """
    Function to search for a name in a given text, taking into account variations in capitalization,
    separators, and slight differences.

    Parameters:
    - name (str): The name to search for.
    - text (str): The string containing the text where the name might be present.
    - threshold (int): The minimum fuzzy match score (0-100) to consider a name as matched.

    Returns:
    - bool: True if a sufficiently close match is found, False otherwise.
    - int: The fuzzy match score (0-100) if a match is found, 0 otherwise.
    """
    # Normalize both the name and the text
    normalized_name = normalize_string(name)
    normalized_text = normalize_string(text)

    # Split the text into potential words/phrases
    words = normalized_text.split()

    # Try to match the normalized name against all possible substrings in the text
    for i in range(len(words)):
        for j in range(i + 1, len(words) + 1):
            substring = ' '.join(words[i:j])
            match_score = fuzz.ratio(normalized_name, substring)
            if match_score >= threshold:
                return True, match_score

    return False, 0

@debug_decorator
def complete_match(name, text):
    """
    Debug function to test the find_name_in_string function with a complete match.
    """
    normalized_name = normalize_string(name)
    normalized_text = normalize_string(text)
    if normalized_name in normalized_text:
        return True, "Identical"


if __name__ == '__main__':

    test_cases = [
        ("John Doe", "AG123123 01.04.20024 John Doe. str Marsweg 62a"),
        ("Jöht Doe", "AG123123 01.04.20024 John -  Doe. str Marsweg 62a"),
        ("Jihn Doe", "AG123123 01.04.20024 John Doe. str Marsweg 62a"),
        ("john doe", "AG123123 01.04.20024 John Doe. str Marsweg 62a"),

        ("00str Marsweg 62a", "AG123123 01.04.20024 John Doe. str Marsweg 62a"),
        ("00John Doe", "The client's name is John Doe.")
    ]
    for name, text in test_cases:
        complete_match(name, text)
    for name, text in test_cases:
        find_name_in_string(name, text)
