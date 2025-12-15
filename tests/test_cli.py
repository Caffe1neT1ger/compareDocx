"""
Тесты для CLI интерфейса.
"""

import pytest
import subprocess
import sys
from pathlib import Path


class TestCLI:
    """Тесты для командной строки."""
    
    @pytest.fixture
    def test_documents(self):
        """Фикстура с путями к тестовым документам."""
        documents_dir = Path(__file__).parent.parent / "documents"
        doc1 = documents_dir / "test_document_1.docx"
        doc2 = documents_dir / "test_document_2.docx"
        
        if not (doc1.exists() and doc2.exists()):
            pytest.skip("Тестовые документы не найдены")
        
        return str(doc1), str(doc2)
    
    def test_cli_help(self):
        """Тест вывода справки."""
        result = subprocess.run(
            [sys.executable, "cli.py", "--help"],
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent.parent
        )
        
        assert result.returncode == 0
        assert "usage" in result.stdout.lower() or "использование" in result.stdout.lower()
        assert "--help" in result.stdout
    
    def test_cli_basic_comparison(self, tmp_path, test_documents):
        """Тест базового сравнения через CLI."""
        doc1, doc2 = test_documents
        
        result = subprocess.run(
            [sys.executable, "cli.py", doc1, doc2, "--xlsx", "--no-llm"],
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent.parent
        )
        
        assert result.returncode == 0
        assert "успешно" in result.stdout.lower() or "success" in result.stdout.lower()
    
    def test_cli_multiple_formats(self, tmp_path, test_documents):
        """Тест экспорта в несколько форматов."""
        doc1, doc2 = test_documents
        
        result = subprocess.run(
            [sys.executable, "cli.py", doc1, doc2, "--xlsx", "--csv", "--json", "--html", "--no-llm"],
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent.parent
        )
        
        assert result.returncode == 0
    
    def test_cli_output_dir(self, tmp_path, test_documents):
        """Тест указания директории для результатов."""
        doc1, doc2 = test_documents
        output_dir = tmp_path / "custom_output"
        
        result = subprocess.run(
            [sys.executable, "cli.py", doc1, doc2, "--xlsx", "--no-llm", "--output-dir", str(output_dir)],
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent.parent
        )
        
        assert result.returncode == 0
        # Проверяем, что результаты созданы в указанной директории
        result_dirs = list(output_dir.glob("comparison_*"))
        assert len(result_dirs) > 0
    
    def test_cli_filter_status(self, tmp_path, test_documents):
        """Тест фильтрации по статусу."""
        doc1, doc2 = test_documents
        
        result = subprocess.run(
            [sys.executable, "cli.py", doc1, doc2, "--json", "--no-llm", "--filter-status", "modified", "added"],
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent.parent
        )
        
        assert result.returncode == 0
    
    def test_cli_filter_min_similarity(self, tmp_path, test_documents):
        """Тест фильтрации по минимальной схожести."""
        doc1, doc2 = test_documents
        
        result = subprocess.run(
            [sys.executable, "cli.py", doc1, doc2, "--json", "--no-llm", "--filter-min-similarity", "0.5"],
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent.parent
        )
        
        assert result.returncode == 0
    
    def test_cli_invalid_file(self):
        """Тест обработки несуществующего файла."""
        result = subprocess.run(
            [sys.executable, "cli.py", "nonexistent1.docx", "nonexistent2.docx", "--no-llm"],
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent.parent
        )
        
        assert result.returncode != 0
    
    def test_cli_log_level(self, tmp_path, test_documents):
        """Тест установки уровня логирования."""
        doc1, doc2 = test_documents
        
        result = subprocess.run(
            [sys.executable, "cli.py", doc1, doc2, "--xlsx", "--no-llm", "--log-level", "DEBUG"],
            capture_output=True,
            text=True,
            cwd=Path(__file__).parent.parent
        )
        
        assert result.returncode == 0

