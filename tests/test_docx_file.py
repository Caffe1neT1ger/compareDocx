"""
Тесты для модуля docx_file.
"""

import pytest
from pathlib import Path
from docx_file import DocxFile
from exceptions import DocumentLoadError, DocumentParseError


class TestDocxFile:
    """Тесты для класса DocxFile."""
    
    def test_load_existing_file(self):
        """Тест загрузки существующего файла."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc_file = documents_dir / "test_document_1.docx"
        
        if not doc_file.exists():
            pytest.skip("Тестовый документ не найден")
        
        docx = DocxFile(str(doc_file))
        assert docx.file_path == str(doc_file)
        assert docx.document is not None
    
    def test_load_nonexistent_file(self):
        """Тест загрузки несуществующего файла."""
        with pytest.raises(DocumentLoadError):
            DocxFile("nonexistent_file.docx")
    
    def test_get_all_paragraphs(self):
        """Тест получения всех абзацев."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc_file = documents_dir / "test_document_1.docx"
        
        if not doc_file.exists():
            pytest.skip("Тестовый документ не найден")
        
        docx = DocxFile(str(doc_file))
        paragraphs = docx.get_all_paragraphs()
        
        assert isinstance(paragraphs, list)
        assert len(paragraphs) > 0
        
        # Проверяем структуру абзаца
        if paragraphs:
            para = paragraphs[0]
            assert "text" in para
            # Проверяем наличие хотя бы одного из полей индекса
            assert any(key in para for key in ["index", "index_1", "section_index", "chapter_index"])
    
    def test_get_tables(self):
        """Тест получения таблиц."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc_file = documents_dir / "test_document_1.docx"
        
        if not doc_file.exists():
            pytest.skip("Тестовый документ не найден")
        
        docx = DocxFile(str(doc_file))
        tables = docx.get_tables()
        
        assert isinstance(tables, list)
        # Может быть 0 таблиц, это нормально
    
    def test_get_images(self):
        """Тест получения изображений."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc_file = documents_dir / "test_document_1.docx"
        
        if not doc_file.exists():
            pytest.skip("Тестовый документ не найден")
        
        docx = DocxFile(str(doc_file))
        images = docx.get_images()
        
        assert isinstance(images, list)
        # Может быть 0 изображений, это нормально
    
    def test_page_estimation(self):
        """Тест оценки страниц."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc_file = documents_dir / "test_document_1.docx"
        
        if not doc_file.exists():
            pytest.skip("Тестовый документ не найден")
        
        docx = DocxFile(str(doc_file))
        paragraphs = docx.get_all_paragraphs()
        
        # Проверяем, что у абзацев есть поле page
        if paragraphs:
            para = paragraphs[0]
            assert "page" in para
            assert isinstance(para["page"], (int, type(None)))
    
    def test_full_path_building(self):
        """Тест построения полного пути."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc_file = documents_dir / "test_document_1.docx"
        
        if not doc_file.exists():
            pytest.skip("Тестовый документ не найден")
        
        docx = DocxFile(str(doc_file))
        paragraphs = docx.get_all_paragraphs()
        
        # Проверяем, что у абзацев есть поле full_path
        if paragraphs:
            para = paragraphs[0]
            assert "full_path" in para
            assert isinstance(para["full_path"], str)

