"""
Масштабное тестирование проекта compareDocx
"""

import subprocess
import sys
import json
import io
from pathlib import Path
import time
from datetime import datetime

# Установка кодировки для Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')

def run_command(cmd, description):
    """Запуск команды и возврат результата"""
    print(f"\n{'='*60}")
    print(f"ТЕСТ: {description}")
    print(f"{'='*60}")
    print(f"Команда: {' '.join(cmd)}")
    
    start_time = time.time()
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        elapsed = time.time() - start_time
        
        if result.returncode == 0:
            print(f"[OK] УСПЕХ ({elapsed:.2f} сек)")
            if result.stdout:
                # Показываем только последние строки вывода
                lines = result.stdout.strip().split('\n')
                for line in lines[-10:]:
                    print(f"  {line}")
            return True, elapsed, result.stdout
        else:
            print(f"[FAIL] ОШИБКА (код: {result.returncode})")
            if result.stderr:
                print("STDERR:")
                for line in result.stderr.split('\n')[:10]:
                    print(f"  {line}")
            return False, elapsed, result.stderr
    except Exception as e:
        elapsed = time.time() - start_time
        print(f"[FAIL] ИСКЛЮЧЕНИЕ: {e}")
        return False, elapsed, str(e)

def check_result_files(result_dir, expected_formats):
    """Проверка наличия файлов результатов"""
    result_path = Path(result_dir)
    if not result_path.exists():
        return False, f"Директория не найдена: {result_dir}"
    
    files = list(result_path.glob("*"))
    found_formats = []
    for fmt in expected_formats:
        pattern = f"*.{fmt}"
        matching = list(result_path.glob(pattern))
        if matching:
            found_formats.append(fmt)
    
    missing = set(expected_formats) - set(found_formats)
    if missing:
        return False, f"Отсутствуют форматы: {missing}. Найдено: {found_formats}"
    
    return True, f"Найдены все форматы: {found_formats}"

def analyze_json_result(json_file):
    """Анализ JSON результата"""
    try:
        with open(json_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        stats = {
            'total': len(data.get('comparison_results', [])),
            'identical': sum(1 for r in data.get('comparison_results', []) if r.get('status') == 'identical'),
            'modified': sum(1 for r in data.get('comparison_results', []) if r.get('status') == 'modified'),
            'added': sum(1 for r in data.get('comparison_results', []) if r.get('status') == 'added'),
            'deleted': sum(1 for r in data.get('comparison_results', []) if r.get('status') == 'deleted'),
            'with_llm': sum(1 for r in data.get('comparison_results', []) if r.get('llm_response')),
            'tables': len(data.get('table_changes', [])),
            'images': len(data.get('image_changes', []))
        }
        return True, stats
    except Exception as e:
        return False, str(e)

def main():
    print("="*60)
    print("МАСШТАБНОЕ ТЕСТИРОВАНИЕ ПРОЕКТА compareDocx")
    print("="*60)
    print(f"Время начала: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    test_results = []
    total_time = 0
    
    # Базовые тесты
    test_cases = [
        # Базовое сравнение
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', '--xlsx'],
            'desc': 'Базовое сравнение (Excel)',
            'formats': ['xlsx']
        },
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', '--json'],
            'desc': 'Базовое сравнение (JSON)',
            'formats': ['json']
        },
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', '--csv'],
            'desc': 'Базовое сравнение (CSV)',
            'formats': ['csv']
        },
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', '--html'],
            'desc': 'Базовое сравнение (HTML)',
            'formats': ['html']
        },
        
        # Множественные форматы
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', '--xlsx', '--json', '--csv', '--html'],
            'desc': 'Экспорт во все форматы одновременно',
            'formats': ['xlsx', 'json', 'csv', 'html']
        },
        
        # Фильтрация
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', '--xlsx', '--filter-status', 'modified', 'added'],
            'desc': 'Фильтрация по статусу (modified, added)',
            'formats': ['xlsx']
        },
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', '--json', '--filter-min-similarity', '0.7'],
            'desc': 'Фильтрация по минимальной схожести (0.7)',
            'formats': ['json']
        },
        
        # Расширенные документы
        {
            'cmd': ['python', 'cli.py', 'documents/extended_test_document_1.docx', 'documents/extended_test_document_2.docx', '--xlsx', '--json'],
            'desc': 'Сравнение расширенных документов',
            'formats': ['xlsx', 'json']
        },
        
        # Дополнительные документы
        {
            'cmd': ['python', 'cli.py', 'documents/additional_test_document_1.docx', 'documents/additional_test_document_2.docx', '--xlsx', '--json'],
            'desc': 'Сравнение документов с кастомными стилями',
            'formats': ['xlsx', 'json']
        },
        
        # Без LLM
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', '--xlsx', '--no-llm'],
            'desc': 'Сравнение без LLM анализа',
            'formats': ['xlsx']
        },
        
        # Логирование
        {
            'cmd': ['python', 'cli.py', 'documents/test_document_1.docx', 'documents/test_document_2.docx', '--xlsx', '--log-level', 'DEBUG'],
            'desc': 'Сравнение с уровнем логирования DEBUG',
            'formats': ['xlsx']
        },
    ]
    
    # Запуск тестов
    for i, test_case in enumerate(test_cases, 1):
        print(f"\n[{i}/{len(test_cases)}]")
        success, elapsed, output = run_command(test_case['cmd'], test_case['desc'])
        total_time += elapsed
        
        result_info = {
            'test': test_case['desc'],
            'success': success,
            'time': elapsed,
            'output': output[:500] if output else ''
        }
        
        # Проверка файлов результатов (если успешно)
        if success:
            # Находим последнюю директорию результатов
            results_dir = Path('results')
            if results_dir.exists():
                result_dirs = sorted([d for d in results_dir.iterdir() if d.is_dir() and d.name.startswith('comparison_')], 
                                    key=lambda x: x.stat().st_mtime, reverse=True)
                if result_dirs:
                    latest_dir = result_dirs[0]
                    files_ok, files_msg = check_result_files(latest_dir, test_case['formats'])
                    result_info['files_check'] = files_ok
                    result_info['files_msg'] = files_msg
                    
                    # Анализ JSON если есть
                    json_files = list(latest_dir.glob("*.json"))
                    if json_files:
                        json_ok, json_stats = analyze_json_result(json_files[0])
                        result_info['json_stats'] = json_stats if json_ok else None
        
        test_results.append(result_info)
    
    # Итоговый отчет
    print("\n" + "="*60)
    print("ИТОГОВЫЙ ОТЧЕТ")
    print("="*60)
    
    passed = sum(1 for r in test_results if r['success'])
    failed = len(test_results) - passed
    
    print(f"\nВсего тестов: {len(test_results)}")
    print(f"Успешно: {passed} ({passed/len(test_results)*100:.1f}%)")
    print(f"Провалено: {failed}")
    print(f"Общее время: {total_time:.2f} сек")
    
    print("\nДетали тестов:")
    for i, result in enumerate(test_results, 1):
        status = "[OK]" if result['success'] else "[FAIL]"
        print(f"{status} [{i}] {result['test']} ({result['time']:.2f} сек)")
        if 'files_check' in result:
            files_status = "[OK]" if result['files_check'] else "[FAIL]"
            print(f"    Файлы: {files_status} {result.get('files_msg', '')}")
        if 'json_stats' in result and result['json_stats']:
            stats = result['json_stats']
            print(f"    Статистика: всего={stats['total']}, изменено={stats['modified']}, "
                  f"добавлено={stats['added']}, LLM={stats['with_llm']}")
    
    # Сохранение отчета
    report_file = Path('results') / f"test_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    report_file.parent.mkdir(exist_ok=True)
    with open(report_file, 'w', encoding='utf-8') as f:
        json.dump({
            'timestamp': datetime.now().isoformat(),
            'summary': {
                'total': len(test_results),
                'passed': passed,
                'failed': failed,
                'total_time': total_time
            },
            'results': test_results
        }, f, indent=2, ensure_ascii=False)
    
    print(f"\nОтчет сохранен: {report_file}")
    
    return 0 if failed == 0 else 1

if __name__ == '__main__':
    sys.exit(main())

