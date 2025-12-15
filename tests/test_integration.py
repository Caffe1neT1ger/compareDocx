"""
Интеграционные тесты для полного цикла работы программы.
"""

import pytest
import os
from pathlib import Path
from compare import Compare
from excel_export import ExcelExporter
from json_export import JSONExporter
from csv_export import CSVExporter
from html_export import HTMLExporter


@pytest.mark.integration
class TestFullComparisonCycle:
    """Интеграционные тесты полного цикла сравнения."""
    
    def test_comparison_with_real_documents(self):
        """Тест сравнения реальных документов из папки documents."""
        documents_dir = Path(__file__).parent.parent / "documents"
        
        # Проверяем наличие тестовых документов
        doc1 = documents_dir / "test_document_1.docx"
        doc2 = documents_dir / "test_document_2.docx"
        
        if not (doc1.exists() and doc2.exists()):
            pytest.skip("Тестовые документы не найдены")
        
        # Выполняем сравнение
        comparator = Compare(str(doc1), str(doc2))
        
        # Проверяем результаты
        results = comparator.get_comparison_results()
        assert len(results) > 0
        
        statistics = comparator.get_statistics()
        assert statistics["total"] > 0
        assert "identical" in statistics
        assert "modified" in statistics
    
    def test_export_to_excel(self, tmp_path):
        """Тест экспорта в Excel."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc1 = documents_dir / "test_document_1.docx"
        doc2 = documents_dir / "test_document_2.docx"
        
        if not (doc1.exists() and doc2.exists()):
            pytest.skip("Тестовые документы не найдены")
        
        comparator = Compare(str(doc1), str(doc2))
        results = comparator.get_comparison_results()
        statistics = comparator.get_statistics()
        
        output_file = tmp_path / "test_output.xlsx"
        exporter = ExcelExporter(str(output_file))
        exporter.export_comparison(
            results,
            statistics,
            "test1.docx",
            "test2.docx",
            comparator.get_table_changes(),
            comparator.get_image_changes()
        )
        
        assert output_file.exists()
    
    def test_export_to_json(self, tmp_path):
        """Тест экспорта в JSON."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc1 = documents_dir / "test_document_1.docx"
        doc2 = documents_dir / "test_document_2.docx"
        
        if not (doc1.exists() and doc2.exists()):
            pytest.skip("Тестовые документы не найдены")
        
        comparator = Compare(str(doc1), str(doc2))
        results = comparator.get_comparison_results()
        statistics = comparator.get_statistics()
        
        output_file = tmp_path / "test_output.json"
        exporter = JSONExporter(str(output_file), pretty=True)
        exporter.export_comparison(
            results,
            statistics,
            "test1.docx",
            "test2.docx",
            comparator.get_table_changes(),
            comparator.get_image_changes()
        )
        
        assert output_file.exists()
        # Проверяем, что файл содержит валидный JSON
        import json
        with open(output_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            assert "statistics" in data
            assert "comparison_results" in data
    
    def test_export_to_csv(self, tmp_path):
        """Тест экспорта в CSV."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc1 = documents_dir / "test_document_1.docx"
        doc2 = documents_dir / "test_document_2.docx"
        
        if not (doc1.exists() and doc2.exists()):
            pytest.skip("Тестовые документы не найдены")
        
        comparator = Compare(str(doc1), str(doc2))
        results = comparator.get_comparison_results()
        statistics = comparator.get_statistics()
        
        exporter = CSVExporter(str(tmp_path))
        exporter.export_comparison(
            results,
            statistics,
            "test1.docx",
            "test2.docx",
            comparator.get_table_changes(),
            comparator.get_image_changes()
        )
        
        # Проверяем, что созданы CSV файлы
        csv_files = list(tmp_path.glob("*.csv"))
        assert len(csv_files) > 0
    
    def test_export_to_html(self, tmp_path):
        """Тест экспорта в HTML."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc1 = documents_dir / "test_document_1.docx"
        doc2 = documents_dir / "test_document_2.docx"
        
        if not (doc1.exists() and doc2.exists()):
            pytest.skip("Тестовые документы не найдены")
        
        comparator = Compare(str(doc1), str(doc2))
        results = comparator.get_comparison_results()
        statistics = comparator.get_statistics()
        
        output_file = tmp_path / "test_output.html"
        exporter = HTMLExporter(str(output_file))
        exporter.export_comparison(
            results,
            statistics,
            "test1.docx",
            "test2.docx",
            comparator.get_table_changes(),
            comparator.get_image_changes()
        )
        
        assert output_file.exists()
        # Проверяем, что файл содержит HTML
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            assert "<html" in content.lower()
            assert "</html>" in content.lower()

