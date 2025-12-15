"""
Модуль командной строки для сравнения DOCX документов.

Предоставляет расширенный CLI с опциями для:
- Выбора формата экспорта (Excel, JSON, CSV, HTML)
- Фильтрации результатов
- Настройки уровня логирования
- Отключения LLM анализа
- И других опций
"""

import argparse
import sys
import os
from pathlib import Path
from compare import Compare
from excel_export import ExcelExporter
from json_export import JSONExporter
from csv_export import CSVExporter
from html_export import HTMLExporter
from llm_adapter import LLMAdapter
from validators import validate_file_path, validate_output_path
from logger_config import logger, setup_logger
from exceptions import CompareDocxError
import logging


def create_parser():
    """Создание парсера аргументов командной строки."""
    parser = argparse.ArgumentParser(
        description='Сравнение двух DOCX документов с детальным анализом изменений',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:

  # Базовое сравнение с экспортом в Excel:
  python cli.py file1.docx file2.docx -o result.xlsx

  # Экспорт в JSON:
  python cli.py file1.docx file2.docx --format json -o result.json

  # Экспорт в несколько форматов:
  python cli.py file1.docx file2.docx --format excel json html

  # Фильтрация только изменений:
  python cli.py file1.docx file2.docx --filter-status modified added

  # Отключение LLM анализа:
  python cli.py file1.docx file2.docx --no-llm

  # Уровень логирования DEBUG:
  python cli.py file1.docx file2.docx --log-level DEBUG
        """
    )
    
    # Обязательные аргументы
    parser.add_argument(
        'file1',
        type=str,
        help='Путь к первому DOCX файлу (базовый документ)'
    )
    
    parser.add_argument(
        'file2',
        type=str,
        help='Путь ко второму DOCX файлу (измененный документ)'
    )
    
    # Опциональные аргументы
    parser.add_argument(
        '-o', '--output',
        type=str,
        default='comparison_result.xlsx',
        help='Путь к выходному файлу (по умолчанию: comparison_result.xlsx)'
    )
    
    parser.add_argument(
        '--format',
        nargs='+',
        choices=['excel', 'json', 'csv', 'html'],
        default=['excel'],
        help='Формат(ы) экспорта результатов (можно указать несколько)'
    )
    
    parser.add_argument(
        '--output-dir',
        type=str,
        help='Директория для сохранения результатов (для CSV создается несколько файлов)'
    )
    
    parser.add_argument(
        '--filter-status',
        nargs='+',
        choices=['identical', 'modified', 'added', 'deleted'],
        help='Фильтр по статусу изменений'
    )
    
    parser.add_argument(
        '--filter-min-similarity',
        type=float,
        help='Минимальная схожесть для фильтрации (0.0-1.0)'
    )
    
    parser.add_argument(
        '--filter-change-types',
        nargs='+',
        help='Фильтр по типам изменений'
    )
    
    parser.add_argument(
        '--no-llm',
        action='store_true',
        help='Отключить LLM анализ изменений'
    )
    
    parser.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='Уровень логирования (по умолчанию: INFO)'
    )
    
    parser.add_argument(
        '--log-file',
        type=str,
        help='Путь к файлу для записи логов'
    )
    
    parser.add_argument(
        '--json-pretty',
        action='store_true',
        default=True,
        help='Форматировать JSON с отступами (по умолчанию: включено)'
    )
    
    parser.add_argument(
        '--json-compact',
        action='store_true',
        help='Компактный JSON без отступов (переопределяет --json-pretty)'
    )
    
    return parser


def main():
    """Основная функция CLI."""
    parser = create_parser()
    args = parser.parse_args()
    
    # Настройка логирования
    log_level = getattr(logging, args.log_level.upper())
    setup_logger(level=log_level, log_file=args.log_file)
    
    logger.info("=" * 60)
    logger.info("Сравнение DOCX документов")
    logger.info("=" * 60)
    
    # Валидация путей к файлам
    try:
        file1_path, file1_path_obj = validate_file_path(args.file1)
        file2_path, file2_path_obj = validate_file_path(args.file2)
        logger.info(f"Валидация файлов успешна")
    except CompareDocxError as e:
        logger.error(f"Ошибка валидации: {e}")
        print(f"\nОшибка: {e}")
        return 1
    
    # Инициализация LLM адаптера (если не отключен)
    llm_adapter = None
    if not args.no_llm:
        try:
            llm_adapter = LLMAdapter()
            if llm_adapter.is_enabled():
                model_info = llm_adapter.get_model_info()
                logger.info(f"LLM адаптер инициализирован: {model_info['model']}")
            else:
                logger.info("LLM адаптер недоступен")
                llm_adapter = None
        except Exception as e:
            logger.warning(f"Не удалось инициализировать LLM адаптер: {e}")
            llm_adapter = None
    
    # Сравнение документов
    try:
        print("\nВыполнение сравнения...")
        comparator = Compare(file1_path, file2_path, llm_adapter=llm_adapter)
        
        # Получение результатов
        results = comparator.get_comparison_results()
        statistics = comparator.get_statistics()
        table_changes = comparator.get_table_changes()
        image_changes = comparator.get_image_changes()
        
        # Применение фильтров
        filters = {}
        if args.filter_status:
            filters["status"] = args.filter_status
        if args.filter_min_similarity is not None:
            filters["min_similarity"] = args.filter_min_similarity
        if args.filter_change_types:
            filters["change_types"] = args.filter_change_types
        
        if filters:
            from json_export import JSONExporter
            filtered_results = JSONExporter("", pretty=False)._apply_filters(results, filters)
            logger.info(f"Применены фильтры: {len(filtered_results)} из {len(results)} результатов")
        else:
            filtered_results = results
        
        # Вывод статистики
        print(f"\n{'='*60}")
        print("Статистика сравнения:")
        print(f"{'='*60}")
        print(f"Обработано абзацев: {statistics['total']}")
        print(f"Идентичных: {statistics['identical']} ({statistics['identical_percent']:.1f}%)")
        print(f"Измененных: {statistics['modified']} ({statistics['modified_percent']:.1f}%)")
        print(f"Добавленных: {statistics['added']} ({statistics['added_percent']:.1f}%)")
        print(f"Удаленных: {statistics['deleted']} ({statistics['deleted_percent']:.1f}%)")
        
        # Статистика по типам изменений
        if 'change_types' in statistics and statistics['change_types']:
            print(f"\nТипы изменений:")
            for change_type, count in sorted(statistics['change_types'].items(), key=lambda x: x[1], reverse=True):
                print(f"  {change_type}: {count}")
        
        if table_changes:
            print(f"\nИзменений в таблицах: {len(table_changes)}")
        if image_changes:
            print(f"Изменений в изображениях: {len(image_changes)}")
        
        if llm_adapter and llm_adapter.is_enabled():
            llm_analyzed = statistics.get("llm_analyzed", 0)
            if llm_analyzed > 0:
                print(f"Проанализировано через LLM: {llm_analyzed} элементов")
        
        # Экспорт в выбранные форматы
        file1_name = os.path.basename(file1_path)
        file2_name = os.path.basename(file2_path)
        
        output_dir = args.output_dir or Path(args.output).parent
        output_base = Path(args.output).stem
        
        for fmt in args.format:
            print(f"\nЭкспорт в {fmt.upper()}...")
            
            if fmt == 'excel':
                output_path = args.output if 'excel' in args.format and len(args.format) == 1 else str(Path(output_dir) / f"{output_base}.xlsx")
                exporter = ExcelExporter(output_path)
                exporter.export_comparison(
                    filtered_results if filters else results,
                    statistics,
                    file1_name,
                    file2_name,
                    table_changes,
                    image_changes
                )
                print(f"Результаты сохранены: {os.path.abspath(output_path)}")
            
            elif fmt == 'json':
                output_path = args.output if fmt == 'json' and len(args.format) == 1 else str(Path(output_dir) / f"{output_base}.json")
                pretty = args.json_pretty and not args.json_compact
                exporter = JSONExporter(output_path, pretty=pretty)
                exporter.export_comparison(
                    filtered_results if filters else results,
                    statistics,
                    file1_name,
                    file2_name,
                    table_changes,
                    image_changes,
                    filters if filters else None
                )
                print(f"Результаты сохранены: {os.path.abspath(output_path)}")
            
            elif fmt == 'csv':
                exporter = CSVExporter(str(output_dir))
                exporter.export_comparison(
                    filtered_results if filters else results,
                    statistics,
                    file1_name,
                    file2_name,
                    table_changes,
                    image_changes
                )
                print(f"Результаты сохранены в директории: {os.path.abspath(output_dir)}")
            
            elif fmt == 'html':
                output_path = args.output if fmt == 'html' and len(args.format) == 1 else str(Path(output_dir) / f"{output_base}.html")
                exporter = HTMLExporter(output_path)
                exporter.export_comparison(
                    filtered_results if filters else results,
                    statistics,
                    file1_name,
                    file2_name,
                    table_changes,
                    image_changes
                )
                print(f"Результаты сохранены: {os.path.abspath(output_path)}")
        
        print(f"\n{'='*60}")
        print("Сравнение завершено успешно!")
        print(f"{'='*60}")
        
    except Exception as e:
        logger.error(f"Ошибка при выполнении сравнения: {e}")
        import traceback
        traceback.print_exc()
        return 1
    
    return 0


if __name__ == "__main__":
    sys.exit(main())

