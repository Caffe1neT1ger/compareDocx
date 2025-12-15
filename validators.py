"""
Модуль для валидации входных данных.

Предоставляет функции для проверки файлов, путей и других входных данных.
"""

import os
from pathlib import Path
from typing import Tuple
from exceptions import ValidationError, FileSizeError
from config import config


def validate_file_path(file_path: str) -> Tuple[str, Path]:
    """
    Валидация пути к файлу.
    
    Args:
        file_path: Путь к файлу
    
    Returns:
        Кортеж (абсолютный путь как строка, Path объект)
    
    Raises:
        ValidationError: Если путь невалиден
    """
    if not file_path:
        raise ValidationError("Путь к файлу не может быть пустым")
    
    # Нормализация пути
    normalized_path = os.path.normpath(file_path.strip().strip('"'))
    path_obj = Path(normalized_path)
    
    # Проверка существования
    if not path_obj.exists():
        raise ValidationError(f"Файл не найден: {normalized_path}")
    
    # Проверка, что это файл, а не директория
    if not path_obj.is_file():
        raise ValidationError(f"Указанный путь не является файлом: {normalized_path}")
    
    # Проверка расширения
    if path_obj.suffix.lower() != '.docx':
        raise ValidationError(
            f"Неподдерживаемый формат файла: {path_obj.suffix}. "
            f"Ожидается .docx"
        )
    
    return str(path_obj.absolute()), path_obj


def validate_file_size(file_path: Path) -> None:
    """
    Проверка размера файла.
    
    Args:
        file_path: Путь к файлу
    
    Raises:
        FileSizeError: Если файл слишком большой
    """
    file_size = file_path.stat().st_size
    file_size_mb = file_size / (1024 * 1024)
    max_size_mb = config.document.max_file_size_mb
    
    if file_size_mb > max_size_mb:
        raise FileSizeError(str(file_path), file_size_mb, max_size_mb)


def validate_output_path(output_path: str) -> Path:
    """
    Валидация пути для выходного файла.
    
    Args:
        output_path: Путь к выходному файлу
    
    Returns:
        Path объект
    
    Raises:
        ValidationError: Если путь невалиден
    """
    if not output_path:
        raise ValidationError("Путь к выходному файлу не может быть пустым")
    
    # Нормализация пути
    normalized_path = os.path.normpath(output_path.strip().strip('"'))
    path_obj = Path(normalized_path)
    
    # Проверка расширения
    if path_obj.suffix.lower() not in ['.xlsx', '.xls']:
        # Если расширение не указано, добавляем .xlsx
        if not path_obj.suffix:
            path_obj = path_obj.with_suffix('.xlsx')
        else:
            raise ValidationError(
                f"Неподдерживаемый формат выходного файла: {path_obj.suffix}. "
                f"Ожидается .xlsx"
            )
    
    # Проверка, что директория существует или может быть создана
    parent_dir = path_obj.parent
    if parent_dir.exists() and not parent_dir.is_dir():
        raise ValidationError(f"Родительский путь не является директорией: {parent_dir}")
    
    # Попытка создать директорию, если её нет
    try:
        parent_dir.mkdir(parents=True, exist_ok=True)
    except (OSError, PermissionError) as e:
        raise ValidationError(f"Не удалось создать директорию для выходного файла: {e}")
    
    return path_obj


def validate_document_structure(paragraphs_count: int, tables_count: int, images_count: int) -> None:
    """
    Проверка структуры документа на соответствие ограничениям.
    
    Args:
        paragraphs_count: Количество абзацев
        tables_count: Количество таблиц
        images_count: Количество изображений
    
    Raises:
        ValidationError: Если структура не соответствует ограничениям
    """
    max_paragraphs = config.document.max_paragraphs
    max_tables = config.document.max_tables
    max_images = config.document.max_images
    
    if paragraphs_count > max_paragraphs:
        raise ValidationError(
            f"Слишком много абзацев в документе: {paragraphs_count}. "
            f"Максимум: {max_paragraphs}"
        )
    
    if tables_count > max_tables:
        raise ValidationError(
            f"Слишком много таблиц в документе: {tables_count}. "
            f"Максимум: {max_tables}"
        )
    
    if images_count > max_images:
        raise ValidationError(
            f"Слишком много изображений в документе: {images_count}. "
            f"Максимум: {max_images}"
        )

