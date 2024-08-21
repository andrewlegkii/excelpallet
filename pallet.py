import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
import os

def load_file(label):
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if filepath:
        label.config(text=f"Выбранный файл: {filepath}")
        return filepath
    return None

def select_first_file():
    global first_file_path
    first_file_path = load_file(first_file_label)
    update_rcenter_list()

def select_second_file():
    global second_file_path
    second_file_path = load_file(second_file_label)

def update_rcenter_list():
    if not first_file_path:
        messagebox.showwarning("Ошибка", "Выберите первую таблицу для загрузки распределительных центров.")
        return
    
    try:
        df = pd.read_excel(first_file_path)
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
        df = pd.read_excel(first_file_path)
        filtered_df = df[df['Распределительный Центр'] == rcenter]

        if 'Дата' in filtered_df.columns:
            dates = sorted(filtered_df['Дата'].dropna().unique().tolist())
            # Теперь даты загружаются в календарь, но выбор происходит вручную через DateEntry
        else:
            messagebox.showerror("Ошибка", "В первой таблице не найдена колонка 'Дата'.")
    except Exception as e:
        messagebox.showerror("Ошибка", f"Произошла ошибка при загрузке дат: {str(e)}")

def process_data():
    try:
        if not first_file_path or not second_file_path:
            messagebox.showwarning("Ошибка", "Выберите оба файла.")
            return

        df_source = pd.read_excel(first_file_path)
        df_target = pd.read_excel(second_file_path)

        # Получаем значения для фильтрации
        date = date_entry.get_date().strftime('%Y-%m-%d')
        rcenter = rcenter_combo.get()

        # Преобразуем колонку с датами в datetime формат для корректной фильтрации
        df_source['Дата'] = pd.to_datetime(df_source['Дата'], format='%Y-%m-%d', errors='coerce')

        # Применяем фильтр по дате и распределительному центру
        filtered_data = df_source[(df_source['Распределительный Центр'] == rcenter) & (df_source['Дата'] == date)]

        # Вставляем данные в целевую таблицу (уже существующую)
        df_target = pd.concat([df_target, filtered_data], ignore_index=True)

        # Сохраняем обновленную таблицу в тот же файл
        df_target.to_excel(second_file_path, index=False)
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

# Переменные для хранения путей к файлам
first_file_path = None
second_file_path = None

# Запуск приложения
root.mainloop()
