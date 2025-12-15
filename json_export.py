"""
Модуль для экспорта результатов сравнения в JSON формат.

Класс JSONExporter предоставляет функционал для:
- Экспорта результатов сравнения в структурированный JSON
- Поддержки форматированного (pretty) и компактного JSON
- Фильтрации результатов перед экспортом
"""

import json
from typing import List, Dict, Optional
from pathlib import Path
from datetime import datetime
from logger_config import logger
from exceptions import ExportError


class JSONExporter:
    """
    Класс для экспорта результатов сравнения в JSON.
    
    Создает структурированный JSON файл с:
    - Детальным сравнением абзацев
    - Статистикой сравнения
    - Изменениями в таблицах и изображениях
    - Метаданными сравнения
    """
    
    def __init__(self, output_path: str, pretty: bool = True):
        """
        Инициализация экспортера.
        
        Args:
            output_path: Путь к выходному JSON файлу
            pretty: Форматировать JSON с отступами (по умолчанию True)
        """
        self.output_path = Path(output_path)
        self.pretty = pretty
        
        # Убеждаемся, что расширение .json
        if self.output_path.suffix.lower() != '.json':
            self.output_path = self.output_path.with_suffix('.json')
    
    def export_comparison(self, comparison_results: List[Dict],
                         statistics: Dict, file1_name: str, file2_name: str,
                         table_changes: List[Dict] = None,
                         image_changes: List[Dict] = None,
                         filters: Optional[Dict] = None):
        """
        Экспорт результатов сравнения в JSON.
        
        Args:
            comparison_results: Список результатов сравнения
            statistics: Статистика сравнения
            file1_name: Имя первого файла
            file2_name: Имя второго файла
            table_changes: Список изменений таблиц
            image_changes: Список изменений изображений
            filters: Опциональные фильтры для результатов
                    {
                        "status": ["modified", "added"],  # Фильтр по статусу
                        "min_similarity": 0.5,  # Минимальная схожесть
                        "has_llm_response": True,  # Только с LLM ответами
                        "change_types": ["Добавление текста"]  # Фильтр по типам изменений
                    }
        """
        try:
            # Применение фильтров
            filtered_results = self._apply_filters(comparison_results, filters)
            
            # Формирование структуры данных
            export_data = {
                "metadata": {
                    "file1": file1_name,
                    "file2": file2_name,
                    "export_date": datetime.now().isoformat(),
                    "total_results": len(filtered_results),
                    "total_original": len(comparison_results)
                },
                "statistics": statistics,
                "comparison_results": filtered_results,
                "table_changes": table_changes or [],
                "image_changes": image_changes or []
            }
            
            # Сериализация в JSON
            if self.pretty:
                json_str = json.dumps(export_data, ensure_ascii=False, indent=2)
            else:
                json_str = json.dumps(export_data, ensure_ascii=False, separators=(',', ':'))
            
            # Сохранение файла
            self.output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.output_path, 'w', encoding='utf-8') as f:
                f.write(json_str)
            
            logger.info(f"Результаты экспортированы в JSON: {self.output_path}")
            
        except Exception as e:
            logger.error(f"Ошибка при экспорте в JSON: {e}")
            raise ExportError(str(self.output_path), str(e))
    
    def _apply_filters(self, results: List[Dict], filters: Optional[Dict]) -> List[Dict]:
        """
        Применение фильтров к результатам.
        
        Args:
            results: Список результатов
            filters: Словарь с фильтрами
            
        Returns:
            Отфильтрованный список результатов
        """
        if not filters:
            return results
        
        filtered = results
        
        # Фильтр по статусу
        if "status" in filters:
            statuses = filters["status"]
            if isinstance(statuses, str):
                statuses = [statuses]
            filtered = [r for r in filtered if r.get("status") in statuses]
        
        # Фильтр по минимальной схожести
        if "min_similarity" in filters:
            min_sim = filters["min_similarity"]
            filtered = [r for r in filtered if r.get("similarity", 0) >= min_sim]
        
        # Фильтр по наличию LLM ответа
        if "has_llm_response" in filters:
            has_llm = filters["has_llm_response"]
            if has_llm:
                filtered = [r for r in filtered if r.get("llm_response")]
            else:
                filtered = [r for r in filtered if not r.get("llm_response")]
        
        # Фильтр по типам изменений
        if "change_types" in filters:
            change_types = filters["change_types"]
            if isinstance(change_types, str):
                change_types = [change_types]
            filtered = [r for r in filtered if r.get("change_type") in change_types]
        
        return filtered

