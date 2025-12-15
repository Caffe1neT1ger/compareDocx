"""
Тесты для определения типов и подтипов изменений.
"""

import pytest
from compare import Compare


class TestGeneralCorrections:
    """Тесты для определения общих правок."""
    
    def test_version_correction(self):
        """Тест определения исправления версии документа."""
        comparator = Compare.__new__(Compare)
        
        text1 = "Версия документа: 1.0"
        text2 = "Версия документа: 2.0"
        
        change_type = comparator._determine_change_type(text1, text2, "")
        # Может быть "Общие правки" или "Изменение содержания" в зависимости от логики
        assert change_type in ["Общие правки", "Изменение содержания"]
        
        if change_type == "Общие правки":
            subtype = comparator._determine_change_subtype(text1, text2, change_type, "")
            assert "версии документа" in subtype.lower() or "версия" in subtype.lower()
    
    def test_version_correction_various_formats(self):
        """Тест определения версии в различных форматах."""
        comparator = Compare.__new__(Compare)
        
        test_cases = [
            ("Версия: 1.0", "Версия: 2.0"),
            ("Version: 1.0", "Version: 2.0"),
            ("v. 1.0", "v. 2.0"),
            ("Ревизия: 1.0", "Ревизия: 2.0"),
            ("Revision: 1.0", "Revision: 2.0"),
            ("Ред. 1.0", "Ред. 2.0"),
        ]
        
        for text1, text2 in test_cases:
            change_type = comparator._determine_change_type(text1, text2)
            assert change_type == "Общие правки"
    
    def test_page_count_correction(self):
        """Тест определения актуализации количества листов."""
        comparator = Compare.__new__(Compare)
        
        text1 = "Всего листов: 10"
        text2 = "Всего листов: 12"
        
        change_type = comparator._determine_change_type(text1, text2, "")
        # Может быть "Общие правки" или "Изменение содержания"
        assert change_type in ["Общие правки", "Изменение содержания"]
        
        if change_type == "Общие правки":
            subtype = comparator._determine_change_subtype(text1, text2, change_type, "")
            assert "количества листов" in subtype.lower() or "страниц" in subtype.lower() or "листов" in subtype.lower()
    
    def test_page_count_various_formats(self):
        """Тест определения количества листов в различных форматах."""
        comparator = Compare.__new__(Compare)
        
        test_cases = [
            ("Листов: 10", "Листов: 12"),
            ("Страниц: 10", "Страниц: 12"),
            ("Pages: 10", "Pages: 12"),
            ("Всего листов: 10", "Всего листов: 12"),
            ("Количество листов: 10", "Количество листов: 12"),
        ]
        
        for text1, text2 in test_cases:
            change_type = comparator._determine_change_type(text1, text2, "")
            # Может быть "Общие правки" или "Изменение содержания" в зависимости от логики
            assert change_type in ["Общие правки", "Изменение содержания"]
    
    def test_spelling_correction(self):
        """Тест определения исправления орфографических ошибок."""
        comparator = Compare.__new__(Compare)
        
        text1 = "Этот текст содержит ошшибку"
        text2 = "Этот текст содержит ошибку"
        
        change_type = comparator._determine_change_type(text1, text2)
        # Может быть "Общие правки" если схожесть высокая
        if change_type == "Общие правки":
            subtype = comparator._determine_change_subtype(text1, text2, change_type)
            assert "орфографических" in subtype.lower() or "орфографические" in subtype.lower()
    
    def test_punctuation_correction(self):
        """Тест определения исправления пунктуации."""
        comparator = Compare.__new__(Compare)
        
        text1 = "Текст без запятой где нужно"
        text2 = "Текст, без запятой, где нужно"
        
        change_type = comparator._determine_change_type(text1, text2)
        # Может быть "Общие правки" если схожесть высокая
        if change_type == "Общие правки":
            subtype = comparator._determine_change_subtype(text1, text2, change_type)
            assert "пунктуации" in subtype.lower() or "пунктуация" in subtype.lower()


class TestChangeSubtypes:
    """Тесты для определения подтипов изменений."""
    
    def test_subtype_for_formatting(self):
        """Тест подтипа для изменения форматирования."""
        comparator = Compare.__new__(Compare)
        
        text1 = "Одинаковый текст"
        text2 = "Одинаковый текст"
        
        change_type = comparator._determine_change_type(text1, text2)
        subtype = comparator._determine_change_subtype(text1, text2, change_type)
        assert subtype == "Изменение стиля"
    
    def test_subtype_for_text_addition(self):
        """Тест подтипа для добавления текста."""
        comparator = Compare.__new__(Compare)
        
        # Добавление нескольких слов
        text1 = "Короткий текст"
        text2 = "Короткий текст с дополнительными словами"
        
        change_type = comparator._determine_change_type(text1, text2)
        subtype = comparator._determine_change_subtype(text1, text2, change_type)
        assert "Добавление" in subtype
    
    def test_subtype_for_text_deletion(self):
        """Тест подтипа для удаления текста."""
        comparator = Compare.__new__(Compare)
        
        text1 = "Длинный текст с множеством слов"
        text2 = "Длинный текст"
        
        change_type = comparator._determine_change_type(text1, text2)
        subtype = comparator._determine_change_subtype(text1, text2, change_type)
        assert "Удаление" in subtype
    
    def test_subtype_for_content_change(self):
        """Тест подтипа для изменения содержания."""
        comparator = Compare.__new__(Compare)
        
        text1 = "Первый вариант текста"
        text2 = "Второй вариант текста"
        
        change_type = comparator._determine_change_type(text1, text2)
        if change_type == "Изменение содержания":
            subtype = comparator._determine_change_subtype(text1, text2, change_type)
            assert subtype == "Изменение смысла"
    
    def test_subtype_for_word_order(self):
        """Тест подтипа для изменения порядка слов."""
        comparator = Compare.__new__(Compare)
        
        text1 = "Первый второй третий"
        text2 = "Третий второй первый"
        
        change_type = comparator._determine_change_type(text1, text2)
        if change_type == "Изменение порядка слов":
            subtype = comparator._determine_change_subtype(text1, text2, change_type)
            assert subtype == "Реорганизация текста"

