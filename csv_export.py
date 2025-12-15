"""
Модуль для экспорта результатов сравнения в CSV формат.

Класс CSVExporter предоставляет функционал для:
- Экспорта результатов сравнения в CSV файлы
- Создания отдельных CSV файлов для разных типов данных
- Правильной обработки кодировки UTF-8 с BOM для Excel
"""

import csv
from typing import List, Dict, Optional
from pathlib import Path
from datetime import datetime
from logger_config import logger
from exceptions import ExportError


class CSVExporter:
    """
    Класс для экспорта результатов сравнения в CSV.
    
    Создает отдельные CSV файлы для:
    - Сравнения абзацев
    - Статистики
    - Изменений в таблицах
    - Изменений в изображениях
    """
    
    def __init__(self, output_dir: str, delimiter: str = ',', encoding: str = 'utf-8-sig'):
        """
        Инициализация экспортера.
        
        Args:
            output_dir: Директория для сохранения CSV файлов
            delimiter: Разделитель полей (по умолчанию ',')
            encoding: Кодировка файлов (по умолчанию 'utf-8-sig' для Excel)
        """
        self.output_dir = Path(output_dir)
        self.delimiter = delimiter
        self.encoding = encoding
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def export_comparison(self, comparison_results: List[Dict],
                         statistics: Dict, file1_name: str, file2_name: str,
                         table_changes: List[Dict] = None,
                         image_changes: List[Dict] = None):
        """
        Экспорт результатов сравнения в CSV файлы.
        
        Args:
            comparison_results: Список результатов сравнения
            statistics: Статистика сравнения
            file1_name: Имя первого файла
            file2_name: Имя второго файла
            table_changes: Список изменений таблиц
            image_changes: Список изменений изображений
        """
        try:
            base_name = Path(file1_name).stem + "_vs_" + Path(file2_name).stem
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Экспорт сравнения абзацев
            comparison_file = self.output_dir / f"{base_name}_comparison_{timestamp}.csv"
            self._export_comparison_results(comparison_file, comparison_results, file1_name, file2_name)
            
            # Экспорт только изменений
            changes_only = [r for r in comparison_results if r.get("status") != "identical"]
            if changes_only:
                changes_file = self.output_dir / f"{base_name}_changes_only_{timestamp}.csv"
                self._export_comparison_results(changes_file, changes_only, file1_name, file2_name)
            
            # Экспорт статистики
            stats_file = self.output_dir / f"{base_name}_statistics_{timestamp}.csv"
            self._export_statistics(stats_file, statistics, file1_name, file2_name)
            
            # Экспорт изменений таблиц
            if table_changes:
                tables_file = self.output_dir / f"{base_name}_tables_{timestamp}.csv"
                self._export_table_changes(tables_file, table_changes)
            
            # Экспорт изменений изображений
            if image_changes:
                images_file = self.output_dir / f"{base_name}_images_{timestamp}.csv"
                self._export_image_changes(images_file, image_changes)
            
            logger.info(f"Результаты экспортированы в CSV: {self.output_dir}")
            
        except Exception as e:
            logger.error(f"Ошибка при экспорте в CSV: {e}")
            raise ExportError(str(self.output_dir), str(e))
    
    def _export_comparison_results(self, file_path: Path, results: List[Dict],
                                   file1_name: str, file2_name: str):
        """Экспорт результатов сравнения абзацев."""
        headers = [
            "№", "Статус", "Тип исправления", "Подтип изменений",
            f"Полный путь ({file1_name})", f"Страница ({file1_name})",
            f"Абзац № ({file1_name})", f"Текст ({file1_name})",
            f"Полный путь ({file2_name})", f"Страница ({file2_name})",
            f"Абзац № ({file2_name})", f"Текст ({file2_name})",
            "Схожесть", "Различия", "Описание изменений", "Ответ LLM"
        ]
        
        with open(file_path, 'w', newline='', encoding=self.encoding) as f:
            writer = csv.writer(f, delimiter=self.delimiter)
            writer.writerow(headers)
            
            for result in results:
                row = [
                    result.get("index_1") or "",
                    result.get("status", ""),
                    result.get("change_type", ""),
                    result.get("change_subtype", ""),
                    result.get("full_path_1") or "",
                    result.get("page_1") or "",
                    result.get("index_1") or "",
                    result.get("text_1") or "",
                    result.get("full_path_2") or "",
                    result.get("page_2") or "",
                    result.get("index_2") or "",
                    result.get("text_2") or "",
                    result.get("similarity", 0),
                    "; ".join(result.get("differences", [])),
                    result.get("change_description", ""),
                    result.get("llm_response", "")
                ]
                writer.writerow(row)
    
    def _export_statistics(self, file_path: Path, statistics: Dict,
                           file1_name: str, file2_name: str):
        """Экспорт статистики."""
        with open(file_path, 'w', newline='', encoding=self.encoding) as f:
            writer = csv.writer(f, delimiter=self.delimiter)
            writer.writerow(["Показатель", "Значение"])
            writer.writerow(["Файл 1", file1_name])
            writer.writerow(["Файл 2", file2_name])
            writer.writerow(["Дата сравнения", datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
            writer.writerow([])
            
            stats_data = [
                ("Всего абзацев", statistics.get("total", 0)),
                ("Идентичных", statistics.get("identical", 0)),
                ("Измененных", statistics.get("modified", 0)),
                ("Добавленных", statistics.get("added", 0)),
                ("Удаленных", statistics.get("deleted", 0)),
                ("Процент идентичных", f"{statistics.get('identical_percent', 0):.1f}%"),
                ("Процент измененных", f"{statistics.get('modified_percent', 0):.1f}%"),
            ]
            
            for key, value in stats_data:
                writer.writerow([key, value])
    
    def _export_table_changes(self, file_path: Path, table_changes: List[Dict]):
        """Экспорт изменений таблиц."""
        headers = ["№", "Статус", "Название таблицы 1", "Название таблицы 2",
                   "Описание", "Описание изменений"]
        
        with open(file_path, 'w', newline='', encoding=self.encoding) as f:
            writer = csv.writer(f, delimiter=self.delimiter)
            writer.writerow(headers)
            
            for idx, change in enumerate(table_changes, 1):
                row = [
                    idx,
                    change.get("status", ""),
                    change.get("table_1_name") or change.get("table_1_index") or "",
                    change.get("table_2_name") or change.get("table_2_index") or "",
                    change.get("description", ""),
                    change.get("change_description", "")
                ]
                writer.writerow(row)
    
    def _export_image_changes(self, file_path: Path, image_changes: List[Dict]):
        """Экспорт изменений изображений."""
        headers = ["№", "Статус", "Название изображения 1", "Название изображения 2",
                   "Описание", "Описание изменений"]
        
        with open(file_path, 'w', newline='', encoding=self.encoding) as f:
            writer = csv.writer(f, delimiter=self.delimiter)
            writer.writerow(headers)
            
            for idx, change in enumerate(image_changes, 1):
                row = [
                    idx,
                    change.get("status", ""),
                    change.get("image_1_name") or change.get("image_1_index") or "",
                    change.get("image_2_name") or change.get("image_2_index") or "",
                    change.get("description", ""),
                    change.get("change_description", "")
                ]
                writer.writerow(row)

