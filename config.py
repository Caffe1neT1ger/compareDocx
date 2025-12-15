"""
Модуль конфигурации проекта.

Содержит все настройки и константы, используемые в проекте.
Позволяет легко изменять параметры без модификации основного кода.
"""

from dataclasses import dataclass
from typing import Optional
import os


@dataclass
class ComparisonConfig:
    """Конфигурация для сравнения документов."""
    
    # Пороги схожести текстов
    similarity_threshold_identical: float = 1.0  # Полное совпадение
    similarity_threshold_high: float = 0.95  # Очень высокая схожесть
    similarity_threshold_medium: float = 0.8  # Средняя схожесть
    similarity_threshold_low: float = 0.6  # Низкая схожесть (минимум для сопоставления)
    
    # Параметры нормализации текста
    normalize_case: bool = False  # Приводить ли к нижнему регистру при сравнении
    
    # Параметры отпечатков текста
    fingerprint_first_words: int = 5  # Количество первых слов для отпечатка
    fingerprint_last_words: int = 5  # Количество последних слов для отпечатка


@dataclass
class DocumentConfig:
    """Конфигурация для работы с документами."""
    
    # Оценка страниц
    chars_per_page: int = 2000  # Примерное количество символов на страницу
    
    # Ограничения безопасности
    max_file_size_mb: int = 50  # Максимальный размер файла в МБ
    max_paragraphs: int = 10000  # Максимальное количество абзацев в документе
    max_tables: int = 1000  # Максимальное количество таблиц
    max_images: int = 500  # Максимальное количество изображений
    
    # Поиск названий таблиц/изображений
    search_backward_paragraphs: int = 10  # Количество абзацев назад для поиска названия


@dataclass
class LLMConfig:
    """Конфигурация для LLM адаптера."""
    
    # Параметры запросов
    timeout_seconds: int = 30  # Таймаут запроса к LLM
    max_retries: int = 3  # Максимальное количество попыток при ошибке
    retry_delay_seconds: float = 1.0  # Задержка между попытками
    
    # Батчинг (группировка запросов)
    enable_batching: bool = True  # Включить группировку запросов
    batch_size: int = 5  # Размер батча для группировки запросов
    max_concurrent_requests: int = 3  # Максимальное количество одновременных запросов


@dataclass
class ExcelExportConfig:
    """Конфигурация для экспорта в Excel."""
    
    # Ограничения для читаемости
    max_differences_display: int = 5  # Максимальное количество различий для отображения
    max_cell_value_length: int = 50  # Максимальная длина значения ячейки в описании
    max_text_length_in_cell: int = 32767  # Максимальная длина текста в ячейке Excel
    
    # Форматирование
    auto_adjust_column_width: bool = True  # Автоматическая подстройка ширины столбцов
    min_column_width: int = 10  # Минимальная ширина столбца
    max_column_width: int = 100  # Максимальная ширина столбца


class Config:
    """
    Главный класс конфигурации проекта.
    
    Объединяет все конфигурации и предоставляет единую точку доступа.
    Можно расширять через переменные окружения или файл конфигурации.
    """
    
    def __init__(self):
        """Инициализация конфигурации с возможностью переопределения через переменные окружения."""
        self.comparison = ComparisonConfig()
        self.document = DocumentConfig()
        self.llm = LLMConfig()
        self.excel = ExcelExportConfig()
        
        # Загрузка настроек из переменных окружения (если есть)
        self._load_from_env()
    
    def _load_from_env(self):
        """Загрузка настроек из переменных окружений."""
        # Схожесть
        if os.getenv("COMPARISON_SIMILARITY_THRESHOLD_MEDIUM"):
            self.comparison.similarity_threshold_medium = float(
                os.getenv("COMPARISON_SIMILARITY_THRESHOLD_MEDIUM")
            )
        
        # Документы
        if os.getenv("DOCUMENT_CHARS_PER_PAGE"):
            self.document.chars_per_page = int(os.getenv("DOCUMENT_CHARS_PER_PAGE"))
        
        if os.getenv("DOCUMENT_MAX_FILE_SIZE_MB"):
            self.document.max_file_size_mb = int(os.getenv("DOCUMENT_MAX_FILE_SIZE_MB"))
        
        # LLM
        if os.getenv("LLM_TIMEOUT_SECONDS"):
            self.llm.timeout_seconds = int(os.getenv("LLM_TIMEOUT_SECONDS"))
        
        if os.getenv("LLM_MAX_CONCURRENT_REQUESTS"):
            self.llm.max_concurrent_requests = int(os.getenv("LLM_MAX_CONCURRENT_REQUESTS"))


# Глобальный экземпляр конфигурации
config = Config()

