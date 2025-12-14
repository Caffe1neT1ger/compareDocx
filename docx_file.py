"""
Модуль для работы с DOCX файлами.

Класс DocxFile предоставляет функционал для:
- Парсинга структуры документа (разделы, главы, абзацы)
- Извлечения таблиц и изображений
- Построения иерархии документа
- Определения уровней заголовков (включая кастомные стили)
- Вычисления примерных номеров страниц
- Построения полных путей до элементов документа

Поддерживает как стандартные стили Word, так и кастомные стили.
"""

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from typing import List, Dict, Optional, Tuple
import re
import hashlib
import io


class DocxFile:
    """
    Класс для работы с DOCX файлами.
    
    Обеспечивает полный парсинг документа с извлечением:
    - Абзацев с метаданными (стиль, уровень, путь, страница)
    - Структуры документа (разделы, подразделы, главы)
    - Таблиц с содержимым
    - Изображений с метаданными
    """
    
    def __init__(self, file_path: str):
        """
        Инициализация класса и парсинг документа.
        
        Args:
            file_path: Путь к DOCX файлу
            
        Примечание:
            При инициализации автоматически выполняется полный парсинг документа.
        """
        self.file_path = file_path
        self.document = Document(file_path)  # Загрузка документа через python-docx
        
        # Структуры данных для хранения информации о документе
        self.paragraphs = []  # Список всех абзацев с метаданными
        self.sections = []  # Список разделов документа (уровень 1-2)
        self.chapters = []  # Список глав документа (уровень 3+)
        self.tables = []  # Список таблиц с содержимым
        self.images = []  # Список изображений с метаданными
        self.hierarchy_stack = []  # Стек для построения иерархии (путь до элемента)
        
        # Автоматический парсинг при инициализации
        self._parse_document()
    
    def _parse_document(self):
        """
        Основной метод парсинга документа.
        
        Выполняет:
        1. Парсинг таблиц и изображений
        2. Парсинг абзацев с определением структуры
        3. Построение иерархии разделов и глав
        4. Вычисление примерных номеров страниц
        5. Построение полных путей до элементов
        """
        # Переменные для отслеживания текущей структуры
        current_section = None  # Текущий раздел
        current_chapter = None  # Текущая глава
        section_index = 0  # Счетчик разделов
        chapter_index = 0  # Счетчик глав
        paragraph_index = 0  # Счетчик абзацев
        estimated_page = 1  # Текущая примерная страница
        chars_per_page = 2000  # Примерное количество символов на страницу (для расчета)
        
        # Парсинг таблиц
        self._parse_tables()
        
        # Парсинг изображений
        self._parse_images()
        
        # Парсинг абзацев
        for para in self.document.paragraphs:
            text = para.text.strip()
            
            # Обновление иерархии при встрече заголовка
            if text:
                style_name = para.style.name if para.style else "Normal"
                heading_level = self._get_heading_level(style_name, para)
                element_type = self._classify_element(text, heading_level, style_name)
                
                # Игнорируем абзацы, которые являются подписями к рисункам/таблицам
                # (не добавляем их в иерархию)
                if not (text.lower().startswith('рисунок') or 
                        text.lower().startswith('таблица') or
                        text.lower().startswith('рис.') or
                        text.lower().startswith('табл.')):
                    if element_type in ["section", "chapter"]:
                        # Обновление стека иерархии
                        self._update_hierarchy_stack(text, heading_level, element_type)
            
            if not text:
                continue
            
            paragraph_index += 1
            
            # Определение стиля абзаца
            style_name = para.style.name if para.style else "Normal"
            
            # Определение уровня заголовка
            heading_level = self._get_heading_level(style_name, para)
            
            # Определение типа элемента
            element_type = self._classify_element(text, heading_level, style_name)
            
            if element_type == "section":
                section_index += 1
                current_section = {
                    "index": section_index,
                    "title": text,
                    "level": heading_level,
                    "paragraphs": []
                }
                self.sections.append(current_section)
                current_chapter = None
                chapter_index = 0
            
            elif element_type == "chapter":
                chapter_index += 1
                current_chapter = {
                    "index": chapter_index,
                    "title": text,
                    "level": heading_level,
                    "paragraphs": []
                }
                if current_section:
                    if "chapters" not in current_section:
                        current_section["chapters"] = []
                    current_section["chapters"].append(current_chapter)
                else:
                    self.chapters.append(current_chapter)
            
            # Вычисление примерной страницы
            total_chars = sum(len(p["text"]) for p in self.paragraphs)
            estimated_page = max(1, (total_chars // chars_per_page) + 1)
            
            # Построение полного пути
            full_path = self._build_full_path(element_type)
            
            # Добавление абзаца
            paragraph_data = {
                "text": text,
                "style": style_name,
                "level": heading_level,
                "type": element_type,
                "section_index": section_index if current_section else None,
                "chapter_index": chapter_index if current_chapter else None,
                "paragraph_index": paragraph_index,
                "full_path": full_path,
                "page": estimated_page
            }
            
            self.paragraphs.append(paragraph_data)
            
            # Добавление в текущий раздел/главу
            if current_chapter:
                current_chapter["paragraphs"].append(paragraph_data)
            elif current_section:
                current_section["paragraphs"].append(paragraph_data)
    
    def _get_heading_level(self, style_name: str, para) -> int:
        """
        Определение уровня заголовка.
        Поддерживает как стандартные стили Word, так и кастомные стили.
        
        Args:
            style_name: Название стиля
            para: Объект абзаца
            
        Returns:
            Уровень заголовка (0 - обычный текст, 1-9 - уровни заголовков)
        """
        # 1. Проверка outline level стиля (работает для кастомных стилей тоже)
        # Это самый надежный способ, так как работает с любыми стилями, у которых задан outline level
        try:
            if para.style:
                # Проверяем outline level через paragraph_format
                if hasattr(para.style, 'paragraph_format'):
                    pf = para.style.paragraph_format
                    if hasattr(pf, 'outline_level') and pf.outline_level is not None:
                        level = pf.outline_level
                        if level >= 1 and level <= 9:
                            return level
                
                # Альтернативный способ через XML (для кастомных стилей)
                try:
                    style_xml = para.style._element
                    if style_xml is not None:
                        # Ищем outline level в XML
                        outline_elem = style_xml.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}outlineLvl')
                        if outline_elem is not None and outline_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val') is not None:
                            level = int(outline_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'))
                            if level >= 0 and level <= 8:  # В XML уровни 0-8 соответствуют заголовкам 1-9
                                return level + 1
                except (AttributeError, ValueError, TypeError):
                    pass
        except (AttributeError, TypeError):
            pass
        
        # 2. Проверка стандартных стилей заголовков Word
        if "Heading" in style_name:
            match = re.search(r'Heading\s*(\d+)', style_name)
            if match:
                return int(match.group(1))
            return 1
        
        # 3. Проверка кастомных стилей, которые могут быть основаны на заголовках
        # Ищем паттерны типа "Заголовок 1", "Заголовок_1", "H1", "Title", "Subtitle" и т.д.
        heading_patterns = [
            (r'заголовок\s*(\d+)', True),  # Заголовок 1, Заголовок_1
            (r'h\s*(\d+)', True),  # H1, H 1
            (r'title', 1),  # Title
            (r'subtitle', 2),  # Subtitle
            (r'heading\s*(\d+)', True),  # heading 1 (разные регистры)
        ]
        
        style_lower = style_name.lower()
        for pattern, level in heading_patterns:
            if isinstance(level, bool) and level:
                match = re.search(pattern, style_lower)
                if match:
                    return int(match.group(1))
            elif isinstance(level, int):
                if re.search(pattern, style_lower):
                    return level
        
        # 4. Анализ форматирования текста (эвристика для кастомных стилей)
        if para.runs:
            first_run = para.runs[0]
            text = para.text.strip()
            
            # Проверка размера шрифта (заголовки обычно крупнее)
            font_size = None
            try:
                if first_run.font and first_run.font.size:
                    font_size = first_run.font.size.pt
            except:
                pass
            
            # Проверка форматирования
            is_bold = first_run.bold
            is_italic = first_run.italic
            
            # Эвристика: короткий жирный текст с большим шрифтом - вероятно заголовок
            if is_bold and len(text) < 150:
                # Если шрифт крупнее 14pt - вероятно заголовок уровня 1-2
                if font_size and font_size >= 14:
                    if font_size >= 18:
                        return 1
                    elif font_size >= 16:
                        return 2
                    else:
                        return 3
                # Если нет информации о размере, но текст короткий и жирный
                elif len(text) < 50:
                    return 1
                elif len(text) < 100:
                    return 2
        
        # 5. Проверка по структуре текста (нумерованные заголовки)
        text = para.text.strip()
        # Паттерн: "1. ", "1.1. ", "1.1.1. " и т.д.
        numbered_pattern = re.match(r'^(\d+(?:\.\d+)*)\.\s+', text)
        if numbered_pattern:
            numbers = numbered_pattern.group(1).split('.')
            level = len(numbers)
            if level <= 9:
                return level
        
        return 0
    
    def _classify_element(self, text: str, heading_level: int, style_name: str) -> str:
        """
        Классификация элемента документа.
        
        Args:
            text: Текст элемента
            heading_level: Уровень заголовка
            style_name: Название стиля
            
        Returns:
            Тип элемента: "section", "chapter", "paragraph"
        """
        if heading_level >= 1:
            # Заголовки уровня 1-2 считаем разделами
            if heading_level <= 2:
                return "section"
            # Заголовки уровня 3+ считаем главами
            else:
                return "chapter"
        
        # Проверка на заголовок по формату (нумерованные заголовки)
        if re.match(r'^\d+\.?\s+[А-ЯЁA-Z]', text) or re.match(r'^[А-ЯЁA-Z][а-яёa-z]+\s+\d+', text):
            if len(text) < 100:
                return "section"
        
        return "paragraph"
    
    def get_all_paragraphs(self) -> List[Dict]:
        """
        Получить все абзацы документа.
        
        Returns:
            Список словарей с данными абзацев
        """
        return self.paragraphs
    
    def get_sections(self) -> List[Dict]:
        """
        Получить все разделы документа.
        
        Returns:
            Список разделов
        """
        return self.sections
    
    def get_chapters(self) -> List[Dict]:
        """
        Получить все главы документа.
        
        Returns:
            Список глав
        """
        return self.chapters
    
    def get_paragraphs_by_section(self, section_index: int) -> List[Dict]:
        """
        Получить абзацы конкретного раздела.
        
        Args:
            section_index: Индекс раздела
            
        Returns:
            Список абзацев раздела
        """
        if 0 <= section_index < len(self.sections):
            return self.sections[section_index].get("paragraphs", [])
        return []
    
    def get_paragraphs_by_chapter(self, chapter_index: int) -> List[Dict]:
        """
        Получить абзацы конкретной главы.
        
        Args:
            chapter_index: Индекс главы
            
        Returns:
            Список абзацев главы
        """
        if 0 <= chapter_index < len(self.chapters):
            return self.chapters[chapter_index].get("paragraphs", [])
        return []
    
    def _update_hierarchy_stack(self, text: str, level: int, element_type: str):
        """
        Обновление стека иерархии при встрече заголовка.
        
        Args:
            text: Текст заголовка
            level: Уровень заголовка
            element_type: Тип элемента
        """
        # Удаляем элементы с уровнем >= текущего
        while self.hierarchy_stack and self.hierarchy_stack[-1]["level"] >= level:
            self.hierarchy_stack.pop()
        
        # Добавляем новый элемент
        self.hierarchy_stack.append({
            "text": text,
            "level": level,
            "type": element_type
        })
    
    def _build_full_path(self, element_type: str) -> str:
        """
        Построение полного пути до элемента.
        
        Args:
            element_type: Тип элемента
            
        Returns:
            Полный путь в формате "Раздел 1 > Подраздел 1.1 > Пункт 1.1.1"
        """
        if not self.hierarchy_stack:
            return ""
        
        path_parts = []
        for item in self.hierarchy_stack:
            text = item["text"]
            level = item["level"]
            
            # Извлечение номера из текста (если есть)
            number_match = re.match(r'^(\d+(?:\.\d+)*)', text)
            if number_match:
                number = number_match.group(1)
                # Определение типа по уровню и номеру
                if level <= 2:
                    path_parts.append(f"Раздел {number}")
                elif level <= 4:
                    path_parts.append(f"Подраздел {number}")
                else:
                    path_parts.append(f"Пункт {number}")
            else:
                # Если номера нет, используем текст
                short_text = text[:50] if len(text) > 50 else text
                if level <= 2:
                    path_parts.append(f"Раздел: {short_text}")
                elif level <= 4:
                    path_parts.append(f"Подраздел: {short_text}")
                else:
                    path_parts.append(f"Пункт: {short_text}")
        
        return " > ".join(path_parts) if path_parts else ""
    
    def _parse_tables(self):
        """
        Парсинг всех таблиц из документа.
        
        Для каждой таблицы извлекает:
        - Индекс таблицы
        - Содержимое всех ячеек (по строкам и столбцам)
        - Количество строк и столбцов
        - Текстовое представление для сравнения
        - Хеш содержимого для быстрого сравнения
        """
        for table_idx, table in enumerate(self.document.tables):
            # Базовая информация о таблице
            table_data = {
                "index": table_idx + 1,  # Номер таблицы (начиная с 1)
                "rows": [],  # Список строк с данными ячеек
                "row_count": len(table.rows),  # Количество строк
                "col_count": len(table.columns) if table.rows else 0  # Количество столбцов
            }
            
            # Извлечение данных из всех ячеек
            for row in table.rows:
                row_data = []
                for cell in row.cells:
                    cell_text = cell.text.strip()  # Текст ячейки без пробелов
                    row_data.append(cell_text)
                table_data["rows"].append(row_data)
            
            # Создание текстового представления для сравнения
            # Формат: каждая строка через табуляцию, строки через перенос
            table_text = "\n".join(["\t".join(row) for row in table_data["rows"]])
            table_data["text"] = table_text
            
            # Хеш для быстрого сравнения идентичных таблиц
            table_data["hash"] = hashlib.md5(table_text.encode()).hexdigest()
            
            self.tables.append(table_data)
    
    def _parse_images(self):
        """
        Парсинг всех изображений из документа.
        
        Использует два метода:
        1. Через relationships (основной метод) - находит изображения через связи документа
        2. Через XML (резервный) - ищет изображения в XML структуре абзацев
        
        Для каждого изображения извлекает:
        - Индекс изображения
        - Размер файла (если доступен)
        - Хеш содержимого для сравнения
        - Индекс абзаца, в котором находится изображение
        """
        image_idx = 0
        processed_rels = set()  # Множество обработанных relationships для избежания дубликатов
        
        # Метод 1: Поиск изображений через relationships (основной метод)
        try:
            for rel in self.document.part.rels.values():
                # Проверяем, является ли связь изображением
                if rel.target_ref and "image" in str(rel.target_ref).lower():
                    rel_id = getattr(rel, 'rId', None)
                    if rel_id and rel_id not in processed_rels:
                        processed_rels.add(rel_id)
                        
                        image_data = {
                            "index": image_idx + 1,
                            "type": "image",
                            "relationship_id": rel_id
                        }
                        
                        # Попытка получить данные изображения для создания хеша
                        try:
                            if hasattr(rel, 'target_part'):
                                image_part = rel.target_part
                                if hasattr(image_part, 'blob'):
                                    # Получаем размер и хеш изображения
                                    image_data["size"] = len(image_part.blob)
                                    image_data["hash"] = hashlib.md5(image_part.blob).hexdigest()
                                else:
                                    # Если blob недоступен, создаем уникальный хеш
                                    image_data["hash"] = f"img_{rel_id}_{image_idx}"
                            else:
                                image_data["hash"] = f"img_{rel_id}_{image_idx}"
                        except Exception:
                            # В случае ошибки создаем уникальный идентификатор
                            image_data["hash"] = f"img_{rel_id}_{image_idx}"
                        
                        self.images.append(image_data)
                        image_idx += 1
        except Exception:
            # Метод 2: Резервный метод через XML (если relationships не сработал)
            try:
                for para_idx, para in enumerate(self.document.paragraphs):
                    for run in para.runs:
                        xml_str = str(run._element.xml)
                        # Ищем изображения в XML структуре
                        if 'pic:pic' in xml_str or 'w:drawing' in xml_str:
                            image_data = {
                                "index": image_idx + 1,
                                "paragraph_index": para_idx + 1,  # Абзац, в котором найдено изображение
                                "type": "image",
                                "hash": f"image_{para_idx}_{image_idx}"  # Уникальный идентификатор
                            }
                            self.images.append(image_data)
                            image_idx += 1
                            break  # Одно изображение на абзац
            except Exception:
                pass  # Если и XML метод не сработал, просто пропускаем
    
    def get_tables(self) -> List[Dict]:
        """
        Получить все таблицы документа.
        
        Returns:
            Список таблиц
        """
        return self.tables
    
    def get_images(self) -> List[Dict]:
        """
        Получить все изображения документа.
        
        Returns:
            Список изображений
        """
        return self.images
    
    def get_structure_info(self) -> Dict:
        """
        Получить информацию о структуре документа.
        
        Returns:
            Словарь с информацией о структуре
        """
        return {
            "total_paragraphs": len(self.paragraphs),
            "total_sections": len(self.sections),
            "total_chapters": len(self.chapters),
            "total_tables": len(self.tables),
            "total_images": len(self.images),
            "sections": [
                {
                    "index": s["index"],
                    "title": s["title"],
                    "paragraphs_count": len(s.get("paragraphs", [])),
                    "chapters_count": len(s.get("chapters", []))
                }
                for s in self.sections
            ]
        }

