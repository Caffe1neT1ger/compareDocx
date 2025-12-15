"""
Скрипт для проверки результатов LLM анализа.
"""

import json
import glob
import sys
from pathlib import Path

# Установка кодировки для Windows
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

# Находим последний JSON файл с результатами, содержащий LLM ответы
results_path = Path("results")
if not results_path.exists():
    print("Папка results не найдена")
    exit(1)

result_dirs = sorted([d for d in results_path.iterdir() if d.is_dir() and d.name.startswith("comparison_")], reverse=True)

if not result_dirs:
    print("Не найдены результаты сравнения")
    exit(1)

# Ищем последний файл с LLM ответами
json_file = None
for result_dir in result_dirs:
    json_files = list(result_dir.glob("*.json"))
    if json_files:
        try:
            with open(json_files[0], 'r', encoding='utf-8') as f:
                data = json.load(f)
                # Проверяем, есть ли LLM ответы
                if any(r.get('llm_response') for r in data.get('comparison_results', [])):
                    json_file = json_files[0]
                    break
        except:
            continue

if not json_file:
    # Если не нашли файл с LLM, берем последний
    latest_dir = result_dirs[0]
    json_files = list(latest_dir.glob("*.json"))
    if json_files:
        json_file = json_files[0]
    else:
        print("Не найден JSON файл с результатами")
        exit(1)

print(f"Проверка файла: {json_file}\n")

# Загрузка данных
with open(json_file, 'r', encoding='utf-8') as f:
    data = json.load(f)

# Статистика
total_results = len(data['comparison_results'])
results_with_llm = [r for r in data['comparison_results'] if r.get('llm_response')]
results_without_llm = [r for r in data['comparison_results'] if not r.get('llm_response')]

print("=" * 60)
print("СТАТИСТИКА LLM АНАЛИЗА")
print("=" * 60)
print(f"Всего элементов: {total_results}")
print(f"С LLM ответами: {len(results_with_llm)} ({len(results_with_llm)/total_results*100:.1f}%)")
print(f"Без LLM ответов: {len(results_without_llm)} ({len(results_without_llm)/total_results*100:.1f}%)")

# Примеры ответов LLM
if results_with_llm:
    print("\n" + "=" * 60)
    print("ПРИМЕРЫ ОТВЕТОВ LLM:")
    print("=" * 60)
    
    for i, result in enumerate(results_with_llm[:10], 1):
        status = result.get('status', 'unknown')
        change_type = result.get('change_type', '')
        change_subtype = result.get('change_subtype', '')
        llm_response = result.get('llm_response', '')
        
        print(f"\n{i}. Статус: {status}, Тип: {change_type}")
        if change_subtype:
            print(f"   Подтип: {change_subtype}")
        print(f"   Ответ LLM: {llm_response[:200]}...")
        if len(llm_response) > 200:
            print(f"   (полный ответ: {len(llm_response)} символов)")

# Статистика по типам изменений с LLM
print("\n" + "=" * 60)
print("СТАТИСТИКА ПО ТИПАМ ИЗМЕНЕНИЙ (с LLM ответами):")
print("=" * 60)

change_types_with_llm = {}
for result in results_with_llm:
    change_type = result.get('change_type', 'Не определен')
    change_types_with_llm[change_type] = change_types_with_llm.get(change_type, 0) + 1

for change_type, count in sorted(change_types_with_llm.items(), key=lambda x: x[1], reverse=True):
    print(f"  {change_type}: {count}")

print("\n" + "=" * 60)
print("ПРОВЕРКА ЗАВЕРШЕНА")
print("=" * 60)

