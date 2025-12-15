"""
Unit-тесты для модуля validators.
"""

import pytest
from pathlib import Path
import tempfile
import os
from validators import (
    validate_file_path,
    validate_file_size,
    validate_output_path,
    validate_document_structure
)
from exceptions import ValidationError, FileSizeError
from config import config


class TestValidateFilePath:
    """Тесты для validate_file_path."""
    
    def test_valid_file_path(self, tmp_path):
        """Тест валидации существующего файла."""
        # Создаем временный DOCX файл
        test_file = tmp_path / "test.docx"
        test_file.write_bytes(b"test content")
        
        result_path, path_obj = validate_file_path(str(test_file))
        assert isinstance(path_obj, Path)
        assert path_obj.exists()
    
    def test_nonexistent_file(self):
        """Тест валидации несуществующего файла."""
        with pytest.raises(ValidationError):
            validate_file_path("nonexistent.docx")
    
    def test_empty_path(self):
        """Тест валидации пустого пути."""
        with pytest.raises(ValidationError):
            validate_file_path("")
    
    def test_wrong_extension(self, tmp_path):
        """Тест валидации файла с неправильным расширением."""
        test_file = tmp_path / "test.txt"
        test_file.write_bytes(b"test")
        
        with pytest.raises(ValidationError):
            validate_file_path(str(test_file))


class TestValidateFileSize:
    """Тесты для validate_file_size."""
    
    def test_valid_file_size(self, tmp_path):
        """Тест валидации файла допустимого размера."""
        test_file = tmp_path / "test.docx"
        # Создаем файл небольшого размера
        test_file.write_bytes(b"x" * 1000)
        
        # Не должно быть исключения
        validate_file_size(test_file)
    
    def test_too_large_file(self, tmp_path, monkeypatch):
        """Тест валидации слишком большого файла."""
        # Временно уменьшаем максимальный размер для теста
        original_max = config.document.max_file_size_mb
        monkeypatch.setattr(config.document, "max_file_size_mb", 0.001)  # 1 KB
        
        test_file = tmp_path / "test.docx"
        # Создаем файл больше лимита
        test_file.write_bytes(b"x" * 2000)
        
        with pytest.raises(FileSizeError):
            validate_file_size(test_file)
        
        # Восстанавливаем оригинальное значение
        monkeypatch.setattr(config.document, "max_file_size_mb", original_max)


class TestValidateOutputPath:
    """Тесты для validate_output_path."""
    
    def test_valid_output_path(self, tmp_path):
        """Тест валидации валидного пути для выходного файла."""
        output_path = tmp_path / "output.xlsx"
        result = validate_output_path(str(output_path))
        assert isinstance(result, Path)
    
    def test_output_path_without_extension(self, tmp_path):
        """Тест добавления расширения .xlsx если не указано."""
        output_path = tmp_path / "output"
        result = validate_output_path(str(output_path))
        assert result.suffix == ".xlsx"
    
    def test_output_path_creates_directory(self, tmp_path):
        """Тест создания директории если её нет."""
        output_path = tmp_path / "new_dir" / "output.xlsx"
        result = validate_output_path(str(output_path))
        assert result.parent.exists()


class TestValidateDocumentStructure:
    """Тесты для validate_document_structure."""
    
    def test_valid_structure(self):
        """Тест валидации валидной структуры документа."""
        # Не должно быть исключения
        validate_document_structure(100, 10, 5)
    
    def test_too_many_paragraphs(self, monkeypatch):
        """Тест валидации документа с слишком большим количеством абзацев."""
        original_max = config.document.max_paragraphs
        monkeypatch.setattr(config.document, "max_paragraphs", 100)
        
        with pytest.raises(ValidationError):
            validate_document_structure(200, 10, 5)
        
        monkeypatch.setattr(config.document, "max_paragraphs", original_max)
    
    def test_too_many_tables(self, monkeypatch):
        """Тест валидации документа с слишком большим количеством таблиц."""
        original_max = config.document.max_tables
        monkeypatch.setattr(config.document, "max_tables", 10)
        
        with pytest.raises(ValidationError):
            validate_document_structure(100, 20, 5)
        
        monkeypatch.setattr(config.document, "max_tables", original_max)

