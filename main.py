import csv
import time
import pandas as pd
import openpyxl
from openpyxl.chart import LineChart, Reference
import random
import string


# Функция для поиска LCS
def lcs_length_and_sequence(A, B):
    m, n = len(A), len(B)
    dp = [[0] * (n + 1) for _ in range(m + 1)]

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if A[i - 1] == B[j - 1]:
                dp[i][j] = dp[i - 1][j - 1] + 1
            else:
                dp[i][j] = max(dp[i - 1][j], dp[i][j - 1])

    return dp[m][n], ""  # Мы фиксируем только длину, без восстановления строки


# Функция для генерации случайных строк
def generate_random_strings(num_lines, min_length=4, max_length=8):
    return [
        ''.join(random.choices(string.ascii_uppercase, k=random.randint(min_length, max_length)))
        for _ in range(num_lines)
    ]


# Функция для создания файлов с возрастающим числом строк
def generate_sequential_files(base_path, start=1000, step=1000, end=100000):
    all_strings = []
    for num_lines in range(start, end + step, step):
        new_strings = generate_random_strings(step)
        all_strings.extend(new_strings)
        file_name = f"{base_path}/file_{num_lines}.csv"
        with open(file_name, 'w', newline='') as f:
            writer = csv.writer(f)
            for string in all_strings:
                writer.writerow([string])
        print(f"Файл {file_name} создан с {num_lines} строками.")


# Функция для проведения эксперимента
def run_experiment(base_path, comparison_file, start=1000, step=1000, end=100000):
    times = []  # Для записи времени выполнения
    num_lines_list = list(range(start, end + step, step))  # Количество строк

    # Чтение строки для сравнения
    with open(comparison_file, 'r') as f:
        reader = csv.reader(f)
        comparison_string = next(reader)[0]
        print(f"Строка для сравнения: {comparison_string}")

    # Выполнение эксперимента
    for num_lines in num_lines_list:
        file_name = f"{base_path}/file_{num_lines}.csv"

        # Считываем строки из текущего файла
        with open(file_name, 'r') as f:
            reader = csv.reader(f)
            strings = [row[0] for row in reader if row]

        # Замер времени выполнения
        start_time = time.time()
        for string in strings:
            lcs_length_and_sequence(string, comparison_string)
        end_time = time.time()

        elapsed_time = end_time - start_time
        times.append((num_lines, elapsed_time))
        print(f"Файл {file_name}: время выполнения {elapsed_time:.6f} секунд")

    return times


# Функция для создания Excel с диаграммой
def create_excel_with_chart(results, output_file):
    """
    Создаёт Excel-файл с данными и диаграммой зависимости времени выполнения от количества строк.
    
    :param results: Список кортежей (количество строк, время выполнения).
    :param output_file: Путь для сохранения Excel-файла.
    """
    # Создание DataFrame из результатов
    df = pd.DataFrame(results, columns=["Количество строк", "Время выполнения (с)"])
    df.to_excel(output_file, sheet_name="Данные", index=False)

    # Открываем файл для добавления диаграммы
    wb = openpyxl.load_workbook(output_file)
    ws = wb["Данные"]

    # Создаём диаграмму
    chart = LineChart()  # Создаём объект диаграммы
    chart.title = "Зависимость времени выполнения от количества строк"
    chart.style = 13
    chart.x_axis.title = "Количество строк"
    chart.y_axis.title = "Время выполнения (с)"

    # Настройка интервалов для подписей на осях
    chart.x_axis.majorUnit = 10000  # Интервал подписей на оси X
    chart.y_axis.majorUnit = 0.05  # Интервал подписей на оси Y

    chart.x_axis.tickLblPos = "low"  # Размещение подписей оси X внизу

    # Указание данных для диаграммы
    data = Reference(ws, min_col=2, min_row=2, max_row=len(results) + 1)
    categories = Reference(ws, min_col=1, min_row=2, max_row=len(results) + 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)

    chart.width = 30  # Ширина диаграммы (в дюймах, 1 дюйм ≈ 96 пикселей)
    chart.height = 15  # Высота диаграммы


    # Добавление диаграммы в Excel
    ws.add_chart(chart, "E5")  # Позиция диаграммы
    wb.save(output_file)
    print(f"Excel-файл с диаграммой создан: {output_file}")


# Основной запуск
base_path = r"C:\Users\YSTS\LCS"  # Путь для файлов
comparison_file = r"C:\Users\YSTS\LCS\file2.csv"  # Эталонная строка

# Генерация файлов
generate_sequential_files(base_path, start=1000, step=1000, end=100000)

# Проведение эксперимента
results = run_experiment(base_path, comparison_file, start=1000, step=1000, end=100000)

# Создание Excel-файла с результатами и диаграммой
output_excel_file = r"C:\Users\YSTS\LCS\lcs_experiment_with_chart.xlsx"
create_excel_with_chart(results, output_excel_file)
