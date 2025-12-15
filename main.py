"""
Основной файл для запуска сравнения DOCX документов.

Точка входа в программу. Поддерживает два режима работы:
1. Командная строка - передача путей через аргументы
2. Интерактивный режим - ввод путей при запуске

Использование:
    python main.py файл1.docx файл2.docx результат.xlsx
    или
    python main.py  # интерактивный режим
"""

import sys
import os
from pathlib import Path
from compare import Compare
from excel_export import ExcelExporter
from llm_adapter import LLMAdapter
from validators import validate_file_path, validate_output_path
from logger_config import logger
from exceptions import CompareDocxError


def main():
    """
    Основная функция для запуска сравнения DOCX документов.
    
    Выполняет:
    1. Получение путей к файлам (командная строка или интерактивно)
    2. Проверку существования файлов
    3. Загрузку и парсинг документов
    4. Сравнение документов
    5. Экспорт результатов в Excel
    6. Вывод статистики
    
    Returns:
        int: Код возврата (0 - успех, 1 - ошибка)
    """
    logger.info("=" * 60)
    logger.info("Сравнение DOCX документов")
    logger.info("=" * 60)
    print("=" * 60)
    print("Сравнение DOCX документов")
    print("=" * 60)
    
    # Получение путей к файлам
    if len(sys.argv) >= 3:
        # Режим командной строки
        file1_path = sys.argv[1]  # Первый документ (базовый)
        file2_path = sys.argv[2]  # Второй документ (измененный)
        output_path = sys.argv[3] if len(sys.argv) >= 4 else "comparison_result.xlsx"  # Выходной файл
    else:
        # Интерактивный режим - запрос путей у пользователя
        print("\nВведите пути к файлам для сравнения:")
        file1_path = input("Путь к первому DOCX файлу: ").strip().strip('"')
        file2_path = input("Путь ко второму DOCX файлу: ").strip().strip('"')
        output_path = input("Путь к выходному Excel файлу (по умолчанию: comparison_result.xlsx): ").strip().strip('"')
        
        if not output_path:
            output_path = "comparison_result.xlsx"
    
    # Валидация путей к файлам
    try:
        file1_path, file1_path_obj = validate_file_path(file1_path)
        file2_path, file2_path_obj = validate_file_path(file2_path)
        output_path_obj = validate_output_path(output_path)
        output_path = str(output_path_obj)
        logger.info(f"Валидация файлов успешна")
    except CompareDocxError as e:
        logger.error(f"Ошибка валидации: {e}")
        print(f"\nОшибка: {e}")
        return 1
    
    try:
        # Шаг 1: Загрузка документов
        print("\nЗагрузка документов...")
        print(f"Файл 1: {os.path.basename(file1_path)}")
        print(f"Файл 2: {os.path.basename(file2_path)}")
        
        # Шаг 1.5: Инициализация LLM адаптера (опционально)
        # Конфигурация читается из .env файла или переменных окружения
        # См. .env.example для примера настройки
        llm_adapter = None
        try:
            llm_adapter = LLMAdapter()  # Читает конфигурацию из .env или переменных окружения
            if llm_adapter.is_enabled():
                model_info = llm_adapter.get_model_info()
                print(f"\nLLM адаптер инициализирован.")
                print(f"  Модель: {model_info['model']}")
                print(f"  Температура: {model_info['temperature']}")
                print(f"  Макс. токенов: {model_info['max_tokens']}")
                print("Будет выполнен дополнительный анализ изменений.")
            else:
                print("\nLLM адаптер недоступен. Сравнение будет выполнено без LLM анализа.")
                print("Для включения LLM анализа создайте файл .env с настройками (см. .env.example)")
                llm_adapter = None
        except Exception as e:
            print(f"\nПредупреждение: не удалось инициализировать LLM адаптер: {e}")
            print("Сравнение будет выполнено без LLM анализа.")
            llm_adapter = None
        
        # Шаг 2: Сравнение документов
        # При создании объекта Compare автоматически выполняется:
        # - Парсинг обоих документов
        # - Сравнение абзацев
        # - Сравнение таблиц
        # - Сравнение изображений
        # - Дополнительный анализ через LLM (если адаптер доступен)
        print("\nВыполнение сравнения...")
        comparator = Compare(file1_path, file2_path, llm_adapter=llm_adapter)
        
        # Шаг 3: Получение результатов
        results = comparator.get_comparison_results()  # Результаты сравнения абзацев
        statistics = comparator.get_statistics()  # Общая статистика
        table_changes = comparator.get_table_changes()  # Изменения в таблицах
        image_changes = comparator.get_image_changes()  # Изменения в изображениях
        
        # Шаг 4: Вывод статистики
        print(f"\nОбработано абзацев: {statistics['total']}")
        print(f"Идентичных: {statistics['identical']} ({statistics['identical_percent']:.1f}%)")
        print(f"Измененных: {statistics['modified']} ({statistics['modified_percent']:.1f}%)")
        print(f"Добавленных: {statistics['added']} ({statistics['added_percent']:.1f}%)")
        print(f"Удаленных: {statistics['deleted']} ({statistics['deleted_percent']:.1f}%)")
        
        if table_changes:
            print(f"\nИзменений в таблицах: {len(table_changes)}")
        if image_changes:
            print(f"Изменений в изображениях: {len(image_changes)}")
        
        if llm_adapter and llm_adapter.is_enabled():
            llm_analyzed = statistics.get("llm_analyzed", 0)
            if llm_analyzed > 0:
                print(f"Проанализировано через LLM: {llm_analyzed} элементов")
        
        # Шаг 5: Экспорт в Excel
        print(f"\nЭкспорт результатов в Excel...")
        exporter = ExcelExporter(output_path)
        exporter.export_comparison(
            results,
            statistics,
            os.path.basename(file1_path),
            os.path.basename(file2_path),
            table_changes,
            image_changes
        )
        
        # Шаг 6: Завершение
        print(f"\nРезультаты сохранены в файл: {os.path.abspath(output_path)}")
        print("=" * 60)
        print("Сравнение завершено успешно!")
        
    except Exception as e:
        # Обработка ошибок с детальным выводом
        print(f"\nОшибка при выполнении сравнения: {str(e)}")
        import traceback
        traceback.print_exc()  # Полный стек вызовов для отладки
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())

