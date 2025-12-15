"""
Unit-тесты для модуля compare.
"""

import pytest
from compare import Compare
from docx_file import DocxFile
from config import config


class TestNormalizeText:
    """Тесты для метода _normalize_text."""
    
    def test_normalize_spaces(self):
        """Тест нормализации множественных пробелов."""
        comparator = Compare.__new__(Compare)
        text = "Текст   с    множественными     пробелами"
        result = comparator._normalize_text(text)
        assert "  " not in result  # Не должно быть двойных пробелов
    
    def test_normalize_line_breaks(self):
        """Тест нормализации переносов строк."""
        comparator = Compare.__new__(Compare)
        text = "Текст\nс\nпереносами\nстрок"
        result = comparator._normalize_text(text)
        assert "\n" not in result
    
    def test_empty_text(self):
        """Тест нормализации пустого текста."""
        comparator = Compare.__new__(Compare)
        result = comparator._normalize_text("")
        assert result == ""
    
    def test_whitespace_only(self):
        """Тест нормализации текста только с пробелами."""
        comparator = Compare.__new__(Compare)
        result = comparator._normalize_text("   \n\n   ")
        assert result == ""


class TestCalculateSimilarity:
    """Тесты для метода _calculate_similarity."""
    
    def test_identical_texts(self):
        """Тест схожести идентичных текстов."""
        comparator = Compare.__new__(Compare)
        text = "Одинаковый текст"
        similarity = comparator._calculate_similarity(text, text)
        assert similarity == 1.0
    
    def test_different_texts(self):
        """Тест схожести разных текстов."""
        comparator = Compare.__new__(Compare)
        text1 = "Первый текст"
        text2 = "Совсем другой текст"
        similarity = comparator._calculate_similarity(text1, text2)
        assert 0.0 <= similarity < 1.0
    
    def test_similar_texts(self):
        """Тест схожести похожих текстов."""
        comparator = Compare.__new__(Compare)
        text1 = "Текст с небольшими изменениями"
        text2 = "Текст с небольшими изменениями и дополнениями"
        similarity = comparator._calculate_similarity(text1, text2)
        assert similarity > 0.5


class TestDetermineChangeType:
    """Тесты для метода _determine_change_type."""
    
    def test_formatting_only(self):
        """Тест определения изменения только форматирования."""
        comparator = Compare.__new__(Compare)
        text1 = "Одинаковый текст"
        text2 = "Одинаковый текст"  # Тот же текст
        change_type = comparator._determine_change_type(text1, text2)
        assert change_type == "Изменение форматирования"
    
    def test_text_addition(self):
        """Тест определения добавления текста."""
        comparator = Compare.__new__(Compare)
        text1 = "Короткий текст"
        text2 = "Короткий текст с дополнительными словами"
        change_type = comparator._determine_change_type(text1, text2)
        assert "Добавление" in change_type or "добавление" in change_type.lower()


@pytest.fixture
def sample_documents(tmp_path):
    """Фикстура для создания тестовых документов."""
    # В реальных тестах здесь можно создать временные DOCX файлы
    # Для упрощения используем моки
    pass

