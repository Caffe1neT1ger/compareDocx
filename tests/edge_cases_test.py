"""
Тестирование граничных случаев и обработки ошибок
"""

import subprocess
import sys
import io
from pathlib import Path

# Установка кодировки для Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def test_edge_case(cmd, description, should_fail=False):
    """Тест граничного случая"""
    print(f"\n{'='*60}")
    print(f"ГРАНИЧНЫЙ СЛУЧАЙ: {description}")
    print(f"{'='*60}")
    print(f"Команда: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace',
            timeout=60
        )
        
        if should_fail:
            if result.returncode != 0:
                print(f"[OK] Ожидаемая ошибка обработана корректно")
                return True
            else:
                print(f"[FAIL] Ожидалась ошибка, но команда выполнилась успешно")
                return False
        else:
            if result.returncode == 0:
                print(f"[OK] Успешно выполнено")
                return True
            else:
                print(f"[FAIL] Ошибка выполнения (код: {result.returncode})")
                if result.stderr:
                    print("STDERR:")
                    for line in result.stderr.split('\n')[:5]:
                        print(f"  {line}")
                return False
    except subprocess.TimeoutExpired:
        print(f"[FAIL] Таймаут выполнения")
        return False
    except Exception as e:
        print(f"[FAIL] Исключение: {e}")
        return False

def main():
    print("="*60)
    print("ТЕСТИРОВАНИЕ ГРАНИЧНЫХ СЛУЧАЕВ И ОБРАБОТКИ ОШИБОК")
    print("="*60)
    
    test_cases = [
        # Несуществующие файлы
        {
            'cmd': ['python', 'cli.py', 'nonexistent1.docx', 'nonexistent2.docx', '--xlsx'],
            'desc': 'Несуществующие файлы',
            'should_fail': True
        },
        
        # Неправильное расширение
        {
            'cmd': ['python', 'cli.py', 'README.md', 'README.md', '--xlsx'],
            'desc': 'Неправильное расширение файла (.md вместо .docx)',
            'should_fail': True
        },
        
        # Один файл не существует
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'nonexistent.docx', '--xlsx'],
            'desc': 'Один файл не существует',
            'should_fail': True
        },
        
        # Справка
        {
            'cmd': ['python', 'cli.py', '--help'],
            'desc': 'Вывод справки',
            'should_fail': False
        },
        
        # Неправильные параметры фильтрации
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', 
                   '--xlsx', '--filter-status', 'invalid_status'],
            'desc': 'Неправильный статус фильтрации',
            'should_fail': False  # Должно обработаться корректно
        },
        
        # Неправильное значение схожести
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', 
                   '--xlsx', '--filter-min-similarity', 'invalid'],
            'desc': 'Неправильное значение схожести (не число)',
            'should_fail': True
        },
        
        # Схожесть вне диапазона
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', 
                   '--xlsx', '--filter-min-similarity', '2.0'],
            'desc': 'Схожесть вне диапазона (>1.0)',
            'should_fail': False  # Должно обработаться корректно
        },
        
        # Путь к несуществующей директории вывода
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', 
                   '--xlsx', '--output-dir', 'nonexistent/path/to/results'],
            'desc': 'Несуществующая директория вывода (должна создаться)',
            'should_fail': False
        },
        
        # Одинаковые файлы
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_1.docx', '--xlsx'],
            'desc': 'Сравнение файла с самим собой',
            'should_fail': False
        },
        
        # Множественные флаги форматов
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', 
                   '--xlsx', '--xlsx', '--json', '--json'],
            'desc': 'Дублирование флагов форматов',
            'should_fail': False
        },
    ]
    
    results = []
    for i, test_case in enumerate(test_cases, 1):
        print(f"\n[{i}/{len(test_cases)}]")
        success = test_edge_case(
            test_case['cmd'],
            test_case['desc'],
            test_case.get('should_fail', False)
        )
        results.append({
            'test': test_case['desc'],
            'success': success
        })
    
    # Итоговый отчет
    print("\n" + "="*60)
    print("ИТОГОВЫЙ ОТЧЕТ")
    print("="*60)
    
    passed = sum(1 for r in results if r['success'])
    failed = len(results) - passed
    
    print(f"\nВсего тестов: {len(results)}")
    print(f"Успешно: {passed} ({passed/len(results)*100:.1f}%)")
    print(f"Провалено: {failed}")
    
    print("\nДетали тестов:")
    for i, result in enumerate(results, 1):
        status = "[OK]" if result['success'] else "[FAIL]"
        print(f"{status} [{i}] {result['test']}")
    
    return 0 if failed == 0 else 1

if __name__ == '__main__':
    sys.exit(main())

