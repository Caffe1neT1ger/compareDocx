"""
Модуль для настройки логирования в проекте.

Предоставляет единую систему логирования для всех модулей проекта.
"""

import logging
import sys
from pathlib import Path
from typing import Optional


def setup_logger(
    name: str = "compareDocx",
    level: int = logging.INFO,
    log_file: Optional[str] = None,
    format_string: Optional[str] = None
) -> logging.Logger:
    """
    Настройка логгера для проекта.
    
    Args:
        name: Имя логгера
        level: Уровень логирования (logging.DEBUG, INFO, WARNING, ERROR)
        log_file: Путь к файлу для записи логов (опционально)
        format_string: Кастомный формат строки логирования (опционально)
    
    Returns:
        Настроенный логгер
    """
    logger = logging.getLogger(name)
    logger.setLevel(level)
    
    # Избегаем дублирования обработчиков
    if logger.handlers:
        return logger
    
    # Формат по умолчанию
    if format_string is None:
        format_string = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    
    formatter = logging.Formatter(format_string)
    
    # Обработчик для консоли
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    # Обработчик для файла (если указан)
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)  # В файл пишем все
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    
    return logger


# Глобальный логгер для проекта
logger = setup_logger()

