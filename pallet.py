import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
from datetime import datetime

def load_file(label, book_var):
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filepath:
        label.config(text=f"Выбранный файл: {filepath}")
        try:
            excel_file = pd.ExcelFile(filepath)
            book_var.set(excel_file.sheet_names[0] if excel_file.sheet_names else '')
            book_combobox['values'] = excel_file.sheet_names
            book_combobox.current(0)
            return filepath
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось открыть файл: {str(e)}")
    return None

def select_first_file():
    global first_file_path
    first_file_path = load_file(first_file_label, first_book_var)
    update_rcenter_list()

def select_second_file():
    global second_file_path
    second_file_path = load_file(second_file_label, second_book_var)

def update_rcenter_list():
    if not first_file_path:
        messagebox.showwarning("Ошибка", "Выберите первую таблицу для загрузки распределительных центров.")
        return
    
    try:
        sheet_name = first_book_var.get()
        df = pd.read_excel(first_file_path, sheet_name=sheet_name)
        if 'Распределительный Центр' in df.columns:
            rcenters = sorted(df['Распределительный Центр'].unique().tolist())
            rcenter_combo['values'] = rcenters
            if rcenters:
                rcenter_combo.current(0)
                update_dates_list()  # Обновляем список дат для первого центра
        else:
            messagebox.showerror("Ошибка", "В первой таблице не найдена колонка 'Распределительный Центр'.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при загрузке данных: {str(e)}")

def update_dates_list():
    rcenter = rcenter_combo.get()
    if not rcenter:
        return

    try:
        sheet_name = first_book_var.get()
        df = pd.read_excel(first_file_path, sheet_name=sheet_name)
        filtered_df = df[df['Распределительный Центр'] == rcenter]

        if 'Дата' in filtered_df.columns:
            dates = sorted(filtered_df['Дата'].dropna().unique().tolist())
            # Обновляем DateEntry с доступными датами
            if dates:
                # Установим минимальную и максимальную дату
                min_date = min(dates)
                max_date = max(dates)
                date_entry.config(
                    mindate=min_date,
                    maxdate=max_date,
                    date_pattern='y-mm-dd'
                )
                # Установим текущую дату как первую доступную
                if dates:
                    date_entry.set_date(min_date)
        else:
            messagebox.showerror("Ошибка", "В первой таблице не найдена колонка 'Дата'.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при загрузке дат: {str(e)}")

def process_data():
    try:
        if not first_file_path or not second_file_path:
            messagebox.showwarning("Ошибка", "Выберите оба файла.")
            return

        sheet_name_src = first_book_var.get()
        sheet_name_tgt = second_book_var.get()

        df_source = pd.read_excel(first_file_path, sheet_name=sheet_name_src)
        df_target = pd.read_excel(second_file_path, sheet_name=sheet_name_tgt)

        # Получаем значения для фильтрации
        date = date_entry.get_date().strftime('%Y-%m-%d')
        rcenter = rcenter_combo.get()

        # Преобразуем колонку с датами в datetime формат для корректной фильтрации
        df_source['Дата'] = pd.to_datetime(df_source['Дата'], format='%Y-%m-%d', errors='coerce')

        # Применяем фильтр по дате и распределительному центру
        filtered_data = df_source[(df_source['Распределительный Центр'] == rcenter) & (df_source['Дата'].dt.strftime('%Y-%m-%d') == date)]

        # Вставляем данные в целевую таблицу (уже существующую)
        df_target = pd.concat([df_target, filtered_data], ignore_index=True)

        # Сохраняем обновленную таблицу в тот же файл
        df_target.to_excel(second_file_path, sheet_name=sheet_name_tgt, index=False)
        messagebox.showinfo("Успех", "Данные успешно добавлены в существующую таблицу!")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")

# Создаем GUI
root = tk.Tk()
root.title("Перенос данных между таблицами")

# Поля для выбора файлов
first_file_label = tk.Label(root, text="Файл с исходными данными не выбран")
first_file_label.pack(pady=5)
first_file_button = tk.Button(root, text="Выбрать первую таблицу", command=select_first_file)
first_file_button.pack(pady=5)

second_file_label = tk.Label(root, text="Файл для вставки данных не выбран")
second_file_label.pack(pady=5)
second_file_button = tk.Button(root, text="Выбрать вторую таблицу", command=select_second_file)
second_file_button.pack(pady=5)

# Поле для выбора книги (таблицы) в первом и втором файле
tk.Label(root, text="Выберите книгу из первого файла:").pack(pady=5)
first_book_var = tk.StringVar()
book_combobox = ttk.Combobox(root, textvariable=first_book_var)
book_combobox.pack(pady=5)

tk.Label(root, text="Выберите книгу из второго файла:").pack(pady=5)
second_book_var = tk.StringVar()
book_combobox2 = ttk.Combobox(root, textvariable=second_book_var)
book_combobox2.pack(pady=5)

# Поле для выбора распределительного центра
tk.Label(root, text="Выберите распределительный центр:").pack(pady=5)
rcenter_combo = ttk.Combobox(root)
rcenter_combo.pack(pady=5)
rcenter_combo.bind("<<ComboboxSelected>>", lambda _: update_dates_list())

# Поле для выбора даты через календарь
tk.Label(root, text="Выберите дату:").pack(pady=5)
date_entry = DateEntry(root, selectmode='day', date_pattern='y-mm-dd')
date_entry.pack(pady=5)

# Кнопка для обработки данных
process_button = tk.Button(root, text="Загрузить данные и обработать", command=process_data)
process_button.pack(pady=10)

# Переменные для хранения путей к файлам и имен таблиц
first_file_path = None
second_file_path = None

# Запуск приложения
root.mainloop()
