"""
Основной файл для запуска сравнения DOCX документов.
"""

import sys
import os
from pathlib import Path
from compare import Compare
from excel_export import ExcelExporter


def main():
    """Основная функция для запуска сравнения."""
    print("=" * 60)
    print("Сравнение DOCX документов")
    print("=" * 60)
    
    # Получение путей к файлам
    if len(sys.argv) >= 3:
        file1_path = sys.argv[1]
        file2_path = sys.argv[2]
        output_path = sys.argv[3] if len(sys.argv) >= 4 else "comparison_result.xlsx"
    else:
        # Интерактивный режим
        print("\nВведите пути к файлам для сравнения:")
        file1_path = input("Путь к первому DOCX файлу: ").strip().strip('"')
        file2_path = input("Путь ко второму DOCX файлу: ").strip().strip('"')
        output_path = input("Путь к выходному Excel файлу (по умолчанию: comparison_result.xlsx): ").strip().strip('"')
        
        if not output_path:
            output_path = "comparison_result.xlsx"
    
    # Проверка существования файлов
    if not os.path.exists(file1_path):
        print(f"Ошибка: Файл '{file1_path}' не найден!")
        return
    
    if not os.path.exists(file2_path):
        print(f"Ошибка: Файл '{file2_path}' не найден!")
        return
    
    try:
        print("\nЗагрузка документов...")
        print(f"Файл 1: {os.path.basename(file1_path)}")
        print(f"Файл 2: {os.path.basename(file2_path)}")
        
        # Сравнение документов
        print("\nВыполнение сравнения...")
        comparator = Compare(file1_path, file2_path)
        
        # Получение результатов
        results = comparator.get_comparison_results()
        statistics = comparator.get_statistics()
        table_changes = comparator.get_table_changes()
        image_changes = comparator.get_image_changes()
        
        print(f"\nОбработано абзацев: {statistics['total']}")
        print(f"Идентичных: {statistics['identical']} ({statistics['identical_percent']:.1f}%)")
        print(f"Измененных: {statistics['modified']} ({statistics['modified_percent']:.1f}%)")
        print(f"Добавленных: {statistics['added']} ({statistics['added_percent']:.1f}%)")
        print(f"Удаленных: {statistics['deleted']} ({statistics['deleted_percent']:.1f}%)")
        
        if table_changes:
            print(f"\nИзменений в таблицах: {len(table_changes)}")
        if image_changes:
            print(f"Изменений в изображениях: {len(image_changes)}")
        
        # Экспорт в Excel
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
        
        print(f"\nРезультаты сохранены в файл: {os.path.abspath(output_path)}")
        print("=" * 60)
        print("Сравнение завершено успешно!")
        
    except Exception as e:
        print(f"\nОшибка при выполнении сравнения: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())

