"""
Модуль для экспорта результатов сравнения в HTML формат.

Класс HTMLExporter предоставляет функционал для:
- Создания интерактивного HTML отчета
- Цветовой индикации изменений
- Фильтрации и поиска через JavaScript
- Удобного просмотра в браузере
"""

from typing import List, Dict, Optional
from pathlib import Path
from datetime import datetime
from logger_config import logger
from exceptions import ExportError


class HTMLExporter:
    """
    Класс для экспорта результатов сравнения в HTML.
    
    Создает интерактивный HTML файл с:
    - Детальным сравнением абзацев
    - Статистикой сравнения
    - Изменениями в таблицах и изображениях
    - Интерактивной фильтрацией и поиском
    """
    
    def __init__(self, output_path: str):
        """
        Инициализация экспортера.
        
        Args:
            output_path: Путь к выходному HTML файлу
        """
        self.output_path = Path(output_path)
        
        # Убеждаемся, что расширение .html
        if self.output_path.suffix.lower() not in ['.html', '.htm']:
            self.output_path = self.output_path.with_suffix('.html')
    
    def export_comparison(self, comparison_results: List[Dict],
                         statistics: Dict, file1_name: str, file2_name: str,
                         table_changes: List[Dict] = None,
                         image_changes: List[Dict] = None,
                         summary_changes: str = ""):
        """
        Экспорт результатов сравнения в HTML.
        
        Args:
            comparison_results: Список результатов сравнения
            statistics: Статистика сравнения
            file1_name: Имя первого файла
            file2_name: Имя второго файла
            table_changes: Список изменений таблиц
            image_changes: Список изменений изображений
        """
        try:
            html_content = self._generate_html(
                comparison_results, statistics, file1_name, file2_name,
                table_changes, image_changes, summary_changes
            )
            
            # Сохранение файла
            self.output_path.parent.mkdir(parents=True, exist_ok=True)
            with open(self.output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            logger.info(f"Результаты экспортированы в HTML: {self.output_path}")
            
        except Exception as e:
            logger.error(f"Ошибка при экспорте в HTML: {e}")
            raise ExportError(str(self.output_path), str(e))
    
    def _generate_html(self, comparison_results: List[Dict],
                      statistics: Dict, file1_name: str, file2_name: str,
                      table_changes: List[Dict] = None,
                      image_changes: List[Dict] = None,
                      summary_changes: str = "") -> str:
        """Генерация HTML контента."""
        
        # Фильтруем только изменения
        changes_only = [r for r in comparison_results if r.get("status") != "identical"]
        
        html = f"""<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Сравнение документов: {file1_name} vs {file2_name}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background-color: #f5f5f5;
            padding: 20px;
        }}
        .container {{
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            padding: 30px;
            border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }}
        h1 {{
            color: #2c3e50;
            margin-bottom: 10px;
            border-bottom: 3px solid #3498db;
            padding-bottom: 10px;
        }}
        h2 {{
            color: #34495e;
            margin-top: 30px;
            margin-bottom: 15px;
            padding: 10px;
            background-color: #ecf0f1;
            border-left: 4px solid #3498db;
        }}
        .metadata {{
            background-color: #f8f9fa;
            padding: 15px;
            border-radius: 5px;
            margin-bottom: 20px;
        }}
        .metadata p {{
            margin: 5px 0;
        }}
        .stats-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 15px;
            margin: 20px 0;
        }}
        .stat-card {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 20px;
            border-radius: 8px;
            text-align: center;
        }}
        .stat-card.identical {{
            background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        }}
        .stat-card.modified {{
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        }}
        .stat-card.added {{
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
        }}
        .stat-card.deleted {{
            background: linear-gradient(135deg, #fa709a 0%, #fee140 100%);
        }}
        .stat-value {{
            font-size: 2em;
            font-weight: bold;
            margin: 10px 0;
        }}
        .stat-label {{
            font-size: 0.9em;
            opacity: 0.9;
        }}
        .filters {{
            background-color: #e8f4f8;
            padding: 15px;
            border-radius: 5px;
            margin: 20px 0;
        }}
        .filters input, .filters select {{
            padding: 8px;
            margin: 5px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
            background: white;
        }}
        th {{
            background-color: #34495e;
            color: white;
            padding: 12px;
            text-align: left;
            position: sticky;
            top: 0;
            z-index: 10;
        }}
        td {{
            padding: 10px;
            border-bottom: 1px solid #ddd;
        }}
        tr:hover {{
            background-color: #f5f5f5;
        }}
        .status-identical {{
            background-color: #d4edda;
        }}
        .status-modified {{
            background-color: #fff3cd;
        }}
        .status-added {{
            background-color: #d1ecf1;
        }}
        .status-deleted {{
            background-color: #f8d7da;
        }}
        .badge {{
            display: inline-block;
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 0.85em;
            font-weight: bold;
        }}
        .badge-identical {{
            background-color: #28a745;
            color: white;
        }}
        .badge-modified {{
            background-color: #ffc107;
            color: #000;
        }}
        .badge-added {{
            background-color: #17a2b8;
            color: white;
        }}
        .badge-deleted {{
            background-color: #dc3545;
            color: white;
        }}
        .text-diff {{
            font-family: 'Courier New', monospace;
            font-size: 0.9em;
            max-height: 200px;
            overflow-y: auto;
            background-color: #f8f9fa;
            padding: 10px;
            border-radius: 4px;
        }}
        .hidden {{
            display: none;
        }}
        .section-toggle {{
            cursor: pointer;
            user-select: none;
        }}
        .section-toggle:hover {{
            background-color: #d5e8f3;
        }}
        .section-content {{
            margin-left: 20px;
        }}
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 Сравнение документов</h1>
        
        <div class="metadata">
            <p><strong>Файл 1:</strong> {file1_name}</p>
            <p><strong>Файл 2:</strong> {file2_name}</p>
            <p><strong>Дата сравнения:</strong> {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
        </div>
        
        <h2>📈 Статистика</h2>
        <div class="stats-grid">
            <div class="stat-card identical">
                <div class="stat-label">Идентичных</div>
                <div class="stat-value">{statistics.get('identical', 0)}</div>
                <div class="stat-label">{statistics.get('identical_percent', 0):.1f}%</div>
            </div>
            <div class="stat-card modified">
                <div class="stat-label">Измененных</div>
                <div class="stat-value">{statistics.get('modified', 0)}</div>
                <div class="stat-label">{statistics.get('modified_percent', 0):.1f}%</div>
            </div>
            <div class="stat-card added">
                <div class="stat-label">Добавленных</div>
                <div class="stat-value">{statistics.get('added', 0)}</div>
                <div class="stat-label">{statistics.get('added_percent', 0):.1f}%</div>
            </div>
            <div class="stat-card deleted">
                <div class="stat-label">Удаленных</div>
                <div class="stat-value">{statistics.get('deleted', 0)}</div>
                <div class="stat-label">{statistics.get('deleted_percent', 0):.1f}%</div>
            </div>
        </div>
        
        {f'''
        <h2>📝 Краткое описание изменений</h2>
        <div class="summary-section" style="background-color: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; border: 1px solid #dee2e6;">
            <pre style="white-space: pre-wrap; font-family: inherit; margin: 0; font-size: 1em; line-height: 1.6;">{self._escape_html(summary_changes)}</pre>
        </div>
        ''' if summary_changes and summary_changes.strip() and summary_changes.strip() != "Общие правки." else ''}
        
        <h2 class="section-toggle" onclick="toggleSection('filters')">🔍 Фильтры и поиск</h2>
        <div id="filters" class="filters section-content">
            <input type="text" id="searchInput" placeholder="Поиск по тексту..." onkeyup="filterTable()">
            <select id="statusFilter" onchange="filterTable()">
                <option value="">Все статусы</option>
                <option value="identical">Идентичные</option>
                <option value="modified">Измененные</option>
                <option value="added">Добавленные</option>
                <option value="deleted">Удаленные</option>
            </select>
            <select id="changeTypeFilter" onchange="filterTable()">
                <option value="">Все типы изменений</option>
                <option value="Без изменений">Без изменений</option>
                <option value="Изменение форматирования">Изменение форматирования</option>
                <option value="Добавление текста">Добавление текста</option>
                <option value="Удаление текста">Удаление текста</option>
            </select>
        </div>
        
        <h2 class="section-toggle" onclick="toggleSection('comparison')">📝 Сравнение абзацев</h2>
        <div id="comparison" class="section-content">
            <table id="comparisonTable">
                <thead>
                    <tr>
                        <th>№</th>
                        <th>Статус</th>
                        <th>Тип исправления</th>
                        <th>Подтип изменений</th>
                        <th>Путь ({file1_name})</th>
                        <th>Страница</th>
                        <th>Текст 1</th>
                        <th>Текст 2</th>
                        <th>Схожесть</th>
                        <th>Описание изменений</th>
                        <th>Ответ LLM</th>
                    </tr>
                </thead>
                <tbody>
"""
        
        # Добавление строк таблицы
        for idx, result in enumerate(comparison_results, 1):
            status = result.get("status", "")
            status_class = f"status-{status}"
            badge_class = f"badge-{status}"
            
            status_ru = {
                "identical": "Идентичен",
                "modified": "Изменен",
                "added": "Добавлен",
                "deleted": "Удален"
            }.get(status, status)
            
            change_desc = result.get('change_description', '')
            llm_resp = result.get('llm_response', '')
            
            # Если нет изменений, ставим "Без изменений"
            if not change_desc and status == 'identical':
                change_desc = 'Без изменений'
            if not llm_resp:
                llm_resp = 'Без изменений'
            
            html += f"""
                    <tr class="{status_class}">
                        <td>{idx}</td>
                        <td><span class="badge {badge_class}">{status_ru}</span></td>
                        <td>{result.get('change_type', '')}</td>
                        <td>{result.get('change_subtype', '')}</td>
                        <td>{result.get('full_path_2') or result.get('full_path_1') or ''}</td>
                        <td>{result.get('page_2') or result.get('page_1') or ''}</td>
                        <td><div class="text-diff">{self._escape_html(result.get('text_1', ''))}</div></td>
                        <td><div class="text-diff">{self._escape_html(result.get('text_2', ''))}</div></td>
                        <td>{result.get('similarity', 0):.2%}</td>
                        <td>{self._escape_html(change_desc)}</td>
                        <td>{self._escape_html(llm_resp)}</td>
                    </tr>
"""
        
        html += """
                </tbody>
            </table>
        </div>
"""
        
        # Таблицы
        if table_changes:
            html += f"""
        <h2 class="section-toggle" onclick="toggleSection('tables')">📊 Изменения в таблицах</h2>
        <div id="tables" class="section-content">
            <table>
                <thead>
                    <tr>
                        <th>№</th>
                        <th>Статус</th>
                        <th>Название таблицы 1</th>
                        <th>Название таблицы 2</th>
                        <th>Описание</th>
                        <th>Описание изменений</th>
                    </tr>
                </thead>
                <tbody>
"""
            for idx, change in enumerate(table_changes, 1):
                html += f"""
                    <tr>
                        <td>{idx}</td>
                        <td>{change.get('status', '')}</td>
                        <td>{change.get('table_1_name') or change.get('table_1_index') or ''}</td>
                        <td>{change.get('table_2_name') or change.get('table_2_index') or ''}</td>
                        <td>{self._escape_html(change.get('description', ''))}</td>
                        <td>{self._escape_html(change.get('change_description', ''))}</td>
                    </tr>
"""
            html += """
                </tbody>
            </table>
        </div>
"""
        
        # Изображения
        if image_changes:
            html += f"""
        <h2 class="section-toggle" onclick="toggleSection('images')">🖼️ Изменения в изображениях</h2>
        <div id="images" class="section-content">
            <table>
                <thead>
                    <tr>
                        <th>№</th>
                        <th>Статус</th>
                        <th>Название изображения 1</th>
                        <th>Название изображения 2</th>
                        <th>Описание</th>
                        <th>Описание изменений</th>
                    </tr>
                </thead>
                <tbody>
"""
            for idx, change in enumerate(image_changes, 1):
                html += f"""
                    <tr>
                        <td>{idx}</td>
                        <td>{change.get('status', '')}</td>
                        <td>{change.get('image_1_name') or change.get('image_1_index') or ''}</td>
                        <td>{change.get('image_2_name') or change.get('image_2_index') or ''}</td>
                        <td>{self._escape_html(change.get('description', ''))}</td>
                        <td>{self._escape_html(change.get('change_description', ''))}</td>
                    </tr>
"""
            html += """
                </tbody>
            </table>
        </div>
"""
        
        # JavaScript для фильтрации и поиска
        html += """
        <script>
            function toggleSection(sectionId) {
                const section = document.getElementById(sectionId);
                if (section.style.display === 'none') {
                    section.style.display = 'block';
                } else {
                    section.style.display = 'none';
                }
            }
            
            function filterTable() {
                const searchInput = document.getElementById('searchInput').value.toLowerCase();
                const statusFilter = document.getElementById('statusFilter').value;
                const changeTypeFilter = document.getElementById('changeTypeFilter').value;
                const table = document.getElementById('comparisonTable');
                const rows = table.getElementsByTagName('tr');
                
                for (let i = 1; i < rows.length; i++) {
                    const row = rows[i];
                    const cells = row.getElementsByTagName('td');
                    
                    // Проверка поиска
                    let matchesSearch = true;
                    if (searchInput) {
                        matchesSearch = false;
                        for (let j = 0; j < cells.length; j++) {
                            if (cells[j].textContent.toLowerCase().includes(searchInput)) {
                                matchesSearch = true;
                                break;
                            }
                        }
                    }
                    
                    // Проверка статуса
                    const statusBadge = cells[1].querySelector('.badge');
                    const rowStatus = statusBadge ? statusBadge.textContent.toLowerCase() : '';
                    let matchesStatus = !statusFilter || 
                        (statusFilter === 'identical' && rowStatus.includes('идентичен')) ||
                        (statusFilter === 'modified' && rowStatus.includes('изменен')) ||
                        (statusFilter === 'added' && rowStatus.includes('добавлен')) ||
                        (statusFilter === 'deleted' && rowStatus.includes('удален'));
                    
                    // Проверка типа изменения
                    const changeType = cells[2].textContent.trim();
                    let matchesChangeType = !changeTypeFilter || changeType === changeTypeFilter;
                    
                    if (matchesSearch && matchesStatus && matchesChangeType) {
                        row.style.display = '';
                    } else {
                        row.style.display = 'none';
                    }
                }
            }
        </script>
    </div>
</body>
</html>
"""
        
        return html
    
    def _escape_html(self, text: str) -> str:
        """Экранирование HTML символов."""
        if not text:
            return ""
        return (str(text)
                .replace('&', '&amp;')
                .replace('<', '&lt;')
                .replace('>', '&gt;')
                .replace('"', '&quot;')
                .replace("'", '&#39;')
                .replace('\n', '<br>'))

