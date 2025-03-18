from docx import Document
import json
from typing import Dict, List, Any, Optional
import sys


def is_bold(paragraph) -> bool:
    """Определяет, содержит ли параграф жирный текст."""
    return any(run.bold for run in paragraph.runs if run.bold is not None)


def insert_into_hierarchy(data: Dict[str, Any], keys: List[str], value: Any):
    """Рекурсивно вставляет значения в иерархический JSON."""
    current = data
    for key in keys[:-1]:
        if key not in current:
            current[key] = {}
        current = current[key]
    current[keys[-1]] = value


def docx_table_to_json(docx_path: str, output_path: Optional[str] = None) -> Dict[str, Any]:
    """
    :param docx_path: Путь к файлу .docx
    :param output_path: Если указан, сохраняет результат в JSON-файл
    :return: Иерархический словарь, где структура определяется по первой колонке,
             а значения берутся из второй колонки.
    """
    try:
        doc = Document(docx_path)
    except Exception as e:
        raise ValueError(f"Ошибка при открытии документа: {e}")

    data = {}  # Итоговый JSON-словарь
    hierarchy_stack = []  # Стек для отслеживания иерархии
    current_level = data  # Текущий уровень вложенности

    for table in doc.tables:
        for row in table.rows:
            if len(row.cells) > 1:
                index_cell = row.cells[0].text.strip()
                text_cell = row.cells[1].text.strip()

                # Пропуск строк с №
                if '№' in index_cell:
                    continue
                
                # Пропуск строк, если во второй колонке менее 3 символов
                if len(text_cell) < 3:
                    continue

                if index_cell and index_cell.replace('.', '').isdigit():  # Если в первой колонке число
                    level = index_cell.count('.')  # Определяем уровень вложенности
                    hierarchy_stack = hierarchy_stack[:level]  # Обрезаем стек до нужного уровня
                    if hierarchy_stack:
                        current_level = hierarchy_stack[-1]  # Получаем родительский уровень
                    else:
                        current_level = data  # Если уровень верхний, работаем с корневым JSON
                    
                    current_level[text_cell] = {}  # Добавляем новый узел
                    hierarchy_stack.append(current_level[text_cell])  # Сохраняем в стек
                elif hierarchy_stack:  # Если в первой колонке пусто
                    hierarchy_stack[-1][text_cell] = {}  # Добавляем в текущий узел

    if output_path:
        try:
            with open(output_path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=4, ensure_ascii=False)
            print(f"Данные успешно сохранены в {output_path}")
        except Exception as e:
            raise ValueError(f"Ошибка при сохранении JSON: {e}")

    return data


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Использование: python table2json.py <путь_к_файлу.docx> <опционально: путь_к_JSON>")
    else:
        docx_file = sys.argv[1]
        json_output = sys.argv[2] if len(sys.argv) > 2 else None
        docx_table_to_json(docx_file, json_output)
