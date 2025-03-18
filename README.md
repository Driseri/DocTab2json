# DOCX Парсер Сервис

## Обзор
Этот сервис извлекает данные из таблиц в `.docx` документах, где:
- **Жирный текст** представляет **названия атрибутов (ключи)**.
- **Обычный текст** представляет **значения (список, связанный с каждым ключом)**.
- Извлеченные данные сохраняются в структурированном формате JSON.

## Возможности
- Парсит таблицы из файлов Microsoft Word (`.docx`).
- Определяет **жирный текст** и использует его как **ключ JSON**.
- Собирает **обычный текст** и связывает его с ближайшим **жирным ключом**.
- Выводит JSON либо **в виде словаря**, либо сохраняет **в файл**.
- Обрабатывает многостраничные таблицы без проблем.

## Установка
Этот сервис требует Python 3.7 или новее.

### Установите зависимости
```bash
pip install python-docx
```

## Использование
### Импорт парсера в ваш Python-проект
```python
from table2json import docx_table_to_json
```

### Разбор `.docx` файла и получение JSON-выходных данных
```python
docx_file = "example.docx"
parsed_data = docx_table_to_json(docx_file)
print(parsed_data)
```

### Сохранение вывода в JSON-файл
```python
docx_table_to_json("example.docx", "output.json")
```

### Использование в командной строке
```bash
python table2json.py example.docx output.json
```
Этот код извлечет данные из `example.docx` и сохранит их в `output.json`.

## Пример вывода
Если в `.docx` файле находится следующая таблица:

| Колонка 1 | Колонка 2 |
|----------|----------|
| -        | **Заголовок 1** |
| -        | Значение 1 |
| -        | Значение 2 |
| -        | **Заголовок 2** |
| -        | Значение 3 |

Выходной JSON будет таким:
```json
{
    "Заголовок 1": ["Значение 1", "Значение 2"],
    "Заголовок 2": ["Значение 3"]
}
```

## Обработка ошибок
- Если `.docx` файл не найден, будет вызвана ошибка.
- Если в документе нет таблиц, будет возвращен пустой словарь.
- Если в таблице нет жирного текста, значения не будут привязаны ни к одному ключу.