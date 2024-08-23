```markdown
# Управление Поступлением Паллет

## Описание проекта

Этот проект представляет собой графическое приложение на Python, которое позволяет переносить данные о паллетах между таблицами Excel. Программа позволяет выбирать распределительные центры и даты, затем автоматически переносит данные о количестве паллет из одной таблицы в другую, обновляя существующие записи, если они уже есть, и добавляя новые, если данных нет.

## Основные функции

1. **Выбор файлов Excel**:
   - Откройте исходный файл, из которого будут извлекаться данные.
   - Выберите целевой файл, в который будут добавляться или обновляться данные.

2. **Выбор книги и листа**:
   - Выберите книгу и лист в каждом файле Excel, с которым хотите работать.

3. **Выбор распределительного центра**:
   - Из списка доступных распределительных центров выберите тот, для которого хотите обработать данные.

4. **Выбор даты**:
   - Выберите дату для обработки данных, доступные даты будут зависеть от выбранного распределительного центра.

5. **Обновление и добавление данных**:
   - Если для выбранного распределительного центра и даты уже есть запись в целевой таблице с пустым значением количества паллет, оно будет обновлено. Если запись уже заполнена, программа уведомит вас об этом.
   - Если записи для выбранной даты и распределительного центра еще нет, она будет добавлена в целевую таблицу.

## Установка и запуск

1. **Установите необходимые зависимости**:

   Сначала установите Python, если он еще не установлен, а затем установите необходимые библиотеки с помощью `pip`:

   ```bash
   pip install pandas tkcalendar babel pyinstaller
   ```

2. **Создайте исполняемый файл**:

   Используйте PyInstaller для создания исполняемого файла:

   ```bash
   pyinstaller --onefile --windowed --collect-all babel main.py
   ```

   Это создаст файл `main.exe` в папке `dist`, который можно запустить на Windows.

3. **Запуск приложения**:

   Запустите созданный исполняемый файл `main.exe`, и вы увидите графический интерфейс для выбора файлов и обработки данных.

## Примечания

- Убедитесь, что ваши таблицы Excel содержат необходимые колонки: `Распределительный Центр`, `Дата`, и `Количество паллет`.
- Проверьте формат даты в таблицах, чтобы он соответствовал формату `YYYY-MM-DD`.