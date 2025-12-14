"""
Модуль для сравнения двух DOCX документов.
Класс Compare предоставляет функционал для сравнения документов по абзацам.
Сравнение выполняется по содержимому текста, игнорируя стили форматирования.
"""

from typing import List, Dict, Tuple, Set
from docx_file import DocxFile
import difflib
import re


class Compare:
    """Класс для сравнения двух DOCX документов."""
    
    def __init__(self, file1_path: str, file2_path: str):
        """
        Инициализация класса сравнения.
        
        Args:
            file1_path: Путь к первому DOCX файлу
            file2_path: Путь ко второму DOCX файлу
        """
        self.file1 = DocxFile(file1_path)
        self.file2 = DocxFile(file2_path)
        self.comparison_results = []
        self.table_changes = []
        self.image_changes = []
        self._compare_documents()
        self._compare_tables()
        self._compare_images()
    
    def _normalize_text(self, text: str) -> str:
        """
        Нормализация текста для сравнения.
        Убирает лишние пробелы, приводит к единому виду, игнорируя форматирование.
        
        Args:
            text: Исходный текст
            
        Returns:
            Нормализованный текст
        """
        if not text:
            return ""
        
        # Приведение к нижнему регистру для сравнения (опционально, но лучше оставить регистр)
        # text = text.lower()
        
        # Замена множественных пробелов на один
        text = re.sub(r'\s+', ' ', text)
        
        # Удаление пробелов в начале и конце
        text = text.strip()
        
        # Нормализация переносов строк (замена на пробелы)
        text = re.sub(r'[\n\r]+', ' ', text)
        
        # Удаление специальных символов форматирования (неразрывные пробелы и т.д.)
        text = text.replace('\xa0', ' ')  # Неразрывный пробел
        text = text.replace('\u2009', ' ')  # Тонкий пробел
        text = text.replace('\u2008', ' ')  # Пунктуационный пробел
        
        # Удаление множественных пробелов после нормализации
        text = re.sub(r'\s+', ' ', text)
        
        return text.strip()
    
    def _get_text_fingerprint(self, text: str) -> str:
        """
        Создание "отпечатка" текста для быстрого сравнения.
        Используется для поиска идентичных абзацев независимо от стилей.
        
        Args:
            text: Исходный текст
            
        Returns:
            Отпечаток текста
        """
        normalized = self._normalize_text(text)
        
        # Создание отпечатка на основе первых и последних слов
        # Это помогает найти похожие абзацы даже при небольших изменениях
        words = normalized.split()
        if len(words) == 0:
            return ""
        
        # Берем первые 5 и последние 5 слов
        first_words = ' '.join(words[:5])
        last_words = ' '.join(words[-5:]) if len(words) > 5 else ''
        
        # Также используем длину и количество слов
        length = len(normalized)
        word_count = len(words)
        
        return f"{first_words}|{last_words}|{length}|{word_count}"
    
    def _compare_documents(self):
        """
        Основной метод сравнения документов.
        Сравнение выполняется по содержимому текста, игнорируя стили форматирования.
        """
        paragraphs1 = self.file1.get_all_paragraphs()
        paragraphs2 = self.file2.get_all_paragraphs()
        
        # Нормализация текстов для сравнения (игнорируя стили)
        normalized_texts1 = [self._normalize_text(p["text"]) for p in paragraphs1]
        normalized_texts2 = [self._normalize_text(p["text"]) for p in paragraphs2]
        
        # Создание отпечатков для быстрого поиска идентичных абзацев
        fingerprints1 = {i: self._get_text_fingerprint(p["text"]) for i, p in enumerate(paragraphs1)}
        fingerprints2 = {i: self._get_text_fingerprint(p["text"]) for i, p in enumerate(paragraphs2)}
        
        # Создание индекса по отпечаткам для второго документа
        fingerprint_index2 = {}
        for idx2, fp in fingerprints2.items():
            if fp:  # Игнорируем пустые отпечатки
                if fp not in fingerprint_index2:
                    fingerprint_index2[fp] = []
                fingerprint_index2[fp].append(idx2)
        
        # Использование SequenceMatcher для поиска похожих абзацев по нормализованному тексту
        matcher = difflib.SequenceMatcher(None, normalized_texts1, normalized_texts2)
        
        # Получение совпадающих блоков
        matching_blocks = matcher.get_matching_blocks()
        
        # Создание карты соответствий
        matched_indices_1 = set()
        matched_indices_2 = set()
        match_map = {}  # Прямое соответствие по позиции
        
        # Обработка совпадающих блоков
        for match in matching_blocks:
            if match.size > 0:
                for offset in range(match.size):
                    idx1 = match.a + offset
                    idx2 = match.b + offset
                    if idx1 < len(paragraphs1) and idx2 < len(paragraphs2):
                        matched_indices_1.add(idx1)
                        matched_indices_2.add(idx2)
                        match_map[idx1] = idx2
        
        # Дополнительный поиск идентичных абзацев по отпечаткам (независимо от позиции)
        # Это помогает найти одинаковые абзацы, которые находятся в разных местах
        for idx1, fp1 in fingerprints1.items():
            if fp1 and fp1 in fingerprint_index2:
                # Найден абзац с таким же отпечатком
                candidates = fingerprint_index2[fp1]
                for idx2 in candidates:
                    if idx2 not in matched_indices_2:
                        # Проверяем, действительно ли тексты идентичны
                        norm1 = normalized_texts1[idx1]
                        norm2 = normalized_texts2[idx2]
                        if norm1 == norm2 or self._calculate_similarity(norm1, norm2) >= 0.95:
                            if idx1 not in match_map:
                                match_map[idx1] = idx2
                                matched_indices_1.add(idx1)
                                matched_indices_2.add(idx2)
                                break
        
        # Сравнение каждого абзаца из первого документа
        for idx1, para1 in enumerate(paragraphs1):
            result = {
                "index_1": idx1 + 1,
                "text_1": para1["text"],
                "section_1": para1.get("section_index"),
                "chapter_1": para1.get("chapter_index"),
                "type_1": para1.get("type", "paragraph"),
                "full_path_1": para1.get("full_path", ""),
                "page_1": para1.get("page", None),
                "index_2": None,
                "text_2": None,
                "section_2": None,
                "chapter_2": None,
                "type_2": None,
                "full_path_2": None,
                "page_2": None,
                "status": "deleted",  # deleted, added, modified, identical
                "similarity": 0.0,
                "differences": [],
                "change_description": "",
                "change_type": ""  # Тип исправления
            }
            
            if idx1 in match_map:
                idx2 = match_map[idx1]
                para2 = paragraphs2[idx2]
                result["index_2"] = idx2 + 1
                result["text_2"] = para2["text"]
                result["section_2"] = para2.get("section_index")
                result["chapter_2"] = para2.get("chapter_index")
                result["type_2"] = para2.get("type", "paragraph")
                result["full_path_2"] = para2.get("full_path", "")
                result["page_2"] = para2.get("page", None)
                
                # Вычисление схожести
                similarity = self._calculate_similarity(para1["text"], para2["text"])
                result["similarity"] = similarity
                
                if similarity == 1.0:
                    result["status"] = "identical"
                    result["change_type"] = "Без изменений"
                elif similarity >= 0.8:
                    result["status"] = "modified"
                    result["differences"] = self._get_differences(para1["text"], para2["text"])
                    result["change_description"] = self._build_change_description(result)
                    result["change_type"] = self._determine_change_type(para1["text"], para2["text"])
                else:
                    result["status"] = "modified"
                    result["differences"] = self._get_differences(para1["text"], para2["text"])
                    result["change_description"] = self._build_change_description(result)
                    result["change_type"] = self._determine_change_type(para1["text"], para2["text"])
            else:
                # Поиск наиболее похожего абзаца по содержимому (игнорируя стили)
                # Используем нормализованный текст для сравнения
                normalized_text1 = normalized_texts1[idx1]
                best_match_idx, best_similarity = self._find_best_match_by_content(
                    normalized_text1, paragraphs2, normalized_texts2, matched_indices_2
                )
                
                if best_similarity >= 0.6:
                    para2 = paragraphs2[best_match_idx]
                    result["index_2"] = best_match_idx + 1
                    result["text_2"] = para2["text"]
                    result["section_2"] = para2.get("section_index")
                    result["chapter_2"] = para2.get("chapter_index")
                    result["type_2"] = para2.get("type", "paragraph")
                    result["full_path_2"] = para2.get("full_path", "")
                    result["page_2"] = para2.get("page", None)
                    result["similarity"] = best_similarity
                    result["status"] = "modified"
                    result["differences"] = self._get_differences(para1["text"], para2["text"])
                    result["change_description"] = self._build_change_description(result)
                    result["change_type"] = self._determine_change_type(para1["text"], para2["text"])
                    matched_indices_2.add(best_match_idx)
                else:
                    result["change_description"] = self._build_change_description(result)
                    result["change_type"] = "Удален"
            
            self.comparison_results.append(result)
        
        # Обработка абзацев из второго документа, которых нет в первом
        for idx2, para2 in enumerate(paragraphs2):
            if idx2 not in matched_indices_2:
                result = {
                    "index_1": None,
                    "text_1": None,
                    "section_1": None,
                    "chapter_1": None,
                    "type_1": None,
                    "full_path_1": None,
                    "page_1": None,
                    "index_2": idx2 + 1,
                    "text_2": para2["text"],
                    "section_2": para2.get("section_index"),
                    "chapter_2": para2.get("chapter_index"),
                    "type_2": para2.get("type", "paragraph"),
                    "full_path_2": para2.get("full_path", ""),
                    "page_2": para2.get("page", None),
                    "status": "added",
                    "similarity": 0.0,
                    "differences": [],
                    "change_description": ""
                }
                result["change_description"] = self._build_change_description(result)
                result["change_type"] = "Добавлен"
                self.comparison_results.append(result)
    
    def _calculate_similarity(self, text1: str, text2: str) -> float:
        """
        Вычисление схожести двух текстов.
        Использует нормализованные тексты для сравнения по содержимому.
        
        Args:
            text1: Первый текст
            text2: Второй текст
            
        Returns:
            Коэффициент схожести от 0.0 до 1.0
        """
        # Нормализуем тексты перед сравнением
        norm1 = self._normalize_text(text1)
        norm2 = self._normalize_text(text2)
        
        # Если тексты идентичны после нормализации
        if norm1 == norm2:
            return 1.0
        
        # Используем SequenceMatcher для вычисления схожести
        return difflib.SequenceMatcher(None, norm1, norm2).ratio()
    
    def _find_best_match(self, text: str, paragraphs: List[Dict], 
                        excluded_indices: set) -> Tuple[int, float]:
        """
        Поиск наиболее похожего абзаца (старый метод, используется для обратной совместимости).
        
        Args:
            text: Текст для поиска
            paragraphs: Список абзацев для поиска
            excluded_indices: Индексы, которые уже использованы
            
        Returns:
            Кортеж (индекс, схожесть)
        """
        normalized_text = self._normalize_text(text)
        best_idx = -1
        best_similarity = 0.0
        
        for idx, para in enumerate(paragraphs):
            if idx in excluded_indices:
                continue
            
            # Сравниваем нормализованные тексты
            normalized_para = self._normalize_text(para["text"])
            similarity = difflib.SequenceMatcher(None, normalized_text, normalized_para).ratio()
            
            if similarity > best_similarity:
                best_similarity = similarity
                best_idx = idx
        
        return best_idx, best_similarity
    
    def _find_best_match_by_content(self, normalized_text: str, paragraphs: List[Dict],
                                   normalized_texts: List[str], excluded_indices: Set[int]) -> Tuple[int, float]:
        """
        Поиск наиболее похожего абзаца по содержимому (игнорируя стили).
        
        Args:
            normalized_text: Нормализованный текст для поиска
            paragraphs: Список абзацев для поиска
            normalized_texts: Список нормализованных текстов
            excluded_indices: Индексы, которые уже использованы
            
        Returns:
            Кортеж (индекс, схожесть)
        """
        best_idx = -1
        best_similarity = 0.0
        
        for idx, normalized_para_text in enumerate(normalized_texts):
            if idx in excluded_indices:
                continue
            
            # Сравниваем нормализованные тексты
            similarity = difflib.SequenceMatcher(None, normalized_text, normalized_para_text).ratio()
            
            if similarity > best_similarity:
                best_similarity = similarity
                best_idx = idx
        
        return best_idx, best_similarity
    
    def _get_differences(self, text1: str, text2: str) -> List[str]:
        """
        Получение списка различий между двумя текстами.
        Возвращает полный текст предложений/абзацев с изменениями.
        
        Args:
            text1: Первый текст
            text2: Второй текст
            
        Returns:
            Список строк с описанием различий (полный текст)
        """
        differences = []
        
        # Нормализуем тексты для сравнения
        norm1 = self._normalize_text(text1)
        norm2 = self._normalize_text(text2)
        
        # Если тексты идентичны после нормализации, но отличаются форматированием
        if norm1 == norm2 and text1 != text2:
            differences.append(f"Изменено только форматирование. Текст: '{text1}'")
            return differences
        
        # Разбиваем на предложения для более точного сравнения
        sentences1 = re.split(r'[.!?]\s+', norm1)
        sentences2 = re.split(r'[.!?]\s+', norm2)
        
        # Если тексты короткие, сравниваем целиком
        if len(norm1) < 100 and len(norm2) < 100:
            if norm1 != norm2:
                differences.append(f"Старый текст: '{text1}'")
                differences.append(f"Новый текст: '{text2}'")
            return differences
        
        # Сравниваем предложения
        matcher = difflib.SequenceMatcher(None, sentences1, sentences2)
        matching_blocks = matcher.get_matching_blocks()
        
        matched_sentences_1 = set()
        matched_sentences_2 = set()
        
        for match in matching_blocks:
            if match.size > 0:
                for i in range(match.a, match.a + match.size):
                    matched_sentences_1.add(i)
                for j in range(match.b, match.b + match.size):
                    matched_sentences_2.add(j)
        
        # Находим удаленные предложения
        for i, sent in enumerate(sentences1):
            if i not in matched_sentences_1 and sent.strip():
                differences.append(f"Удалено предложение: '{sent.strip()}'")
        
        # Находим добавленные предложения
        for j, sent in enumerate(sentences2):
            if j not in matched_sentences_2 and sent.strip():
                differences.append(f"Добавлено предложение: '{sent.strip()}'")
        
        # Если не нашли различий по предложениям, сравниваем по словам
        if not differences:
            words1 = norm1.split()
            words2 = norm2.split()
            
            # Находим измененные фразы
            matcher = difflib.SequenceMatcher(None, words1, words2)
            opcodes = matcher.get_opcodes()
            
            for tag, i1, i2, j1, j2 in opcodes:
                if tag == 'delete':
                    removed = ' '.join(words1[i1:i2])
                    if removed:
                        differences.append(f"Удалено: '{removed}'")
                elif tag == 'insert':
                    added = ' '.join(words2[j1:j2])
                    if added:
                        differences.append(f"Добавлено: '{added}'")
                elif tag == 'replace':
                    removed = ' '.join(words1[i1:i2])
                    added = ' '.join(words2[j1:j2])
                    if removed and added:
                        differences.append(f"'{removed}' изменено на '{added}'")
        
        # Если все еще нет различий, возвращаем полные тексты
        if not differences:
            differences.append(f"Старый текст: '{text1}'")
            differences.append(f"Новый текст: '{text2}'")
        
        return differences[:10]  # Увеличиваем лимит для полных текстов
    
    def get_comparison_results(self) -> List[Dict]:
        """
        Получить результаты сравнения.
        
        Returns:
            Список результатов сравнения
        """
        return self.comparison_results
    
    def _compare_tables(self):
        """Сравнение таблиц в документах с детальным описанием изменений."""
        tables1 = self.file1.get_tables()
        tables2 = self.file2.get_tables()
        
        # Получаем названия таблиц из предыдущих абзацев
        def get_table_name(table_index, paragraphs, is_first_doc):
            # Ищем абзац перед таблицей с текстом "Таблица" или номером
            # Ищем в последних 10 абзацах перед предполагаемой позицией таблицы
            search_start = max(0, table_index * 3 - 10)
            for para in reversed(paragraphs[search_start:table_index * 3]):
                text = para.get("text", "").strip()
                if text and ("таблица" in text.lower() or "табл." in text.lower()):
                    return text
            return f"Таблица {table_index}"
        
        # Создание карты таблиц по хешу
        tables1_map = {t["hash"]: t for t in tables1}
        tables2_map = {t["hash"]: t for t in tables2}
        
        # Поиск идентичных таблиц
        matched_hashes = set()
        for hash_val in tables1_map:
            if hash_val in tables2_map:
                matched_hashes.add(hash_val)
                t1 = tables1_map[hash_val]
                t2 = tables2_map[hash_val]
                name1 = get_table_name(t1["index"], self.file1.get_all_paragraphs(), True)
                name2 = get_table_name(t2["index"], self.file2.get_all_paragraphs(), False)
                self.table_changes.append({
                    "status": "identical",
                    "table_1_index": t1["index"],
                    "table_2_index": t2["index"],
                    "table_1_name": name1,
                    "table_2_name": name2,
                    "description": f"{name1} идентична {name2}",
                    "change_description": "Без изменений"
                })
        
        # Поиск измененных таблиц (похожих, но не идентичных)
        for t1 in tables1:
            if t1["hash"] not in matched_hashes:
                best_match = None
                best_similarity = 0.0
                
                for t2 in tables2:
                    if t2["hash"] not in matched_hashes:
                        similarity = self._calculate_similarity(t1["text"], t2["text"])
                        if similarity > best_similarity and similarity >= 0.5:
                            best_similarity = similarity
                            best_match = t2
                
                if best_match:
                    name1 = get_table_name(t1["index"], self.file1.get_all_paragraphs(), True)
                    name2 = get_table_name(best_match["index"], self.file2.get_all_paragraphs(), False)
                    # Находим конкретные изменения в ячейках
                    cell_changes = self._find_table_cell_changes(t1, best_match)
                    change_desc = self._build_table_change_description(t1, best_match, cell_changes)
                    
                    self.table_changes.append({
                        "status": "modified",
                        "table_1_index": t1["index"],
                        "table_2_index": best_match["index"],
                        "table_1_name": name1,
                        "table_2_name": name2,
                        "similarity": best_similarity,
                        "description": f"{name1} изменена (схожесть {best_similarity*100:.1f}%)",
                        "change_description": change_desc,
                        "cell_changes": cell_changes
                    })
                    matched_hashes.add(best_match["hash"])
                else:
                    name1 = get_table_name(t1["index"], self.file1.get_all_paragraphs(), True)
                    self.table_changes.append({
                        "status": "deleted",
                        "table_1_index": t1["index"],
                        "table_2_index": None,
                        "table_1_name": name1,
                        "table_2_name": "",
                        "description": f"{name1} удалена",
                        "change_description": f"Удалена таблица: '{name1}'"
                    })
        
        # Поиск добавленных таблиц
        for t2 in tables2:
            if t2["hash"] not in matched_hashes:
                name2 = get_table_name(t2["index"], self.file2.get_all_paragraphs(), False)
                self.table_changes.append({
                    "status": "added",
                    "table_1_index": None,
                    "table_2_index": t2["index"],
                    "table_1_name": "",
                    "table_2_name": name2,
                    "description": f"{name2} добавлена",
                    "change_description": f"Добавлена таблица: '{name2}'"
                })
    
    def _find_table_cell_changes(self, table1: Dict, table2: Dict) -> List[Dict]:
        """Находит изменения в ячейках таблиц."""
        changes = []
        rows1 = table1.get("rows", [])
        rows2 = table2.get("rows", [])
        
        max_rows = max(len(rows1), len(rows2))
        max_cols = max(
            max(len(row) for row in rows1) if rows1 else 0,
            max(len(row) for row in rows2) if rows2 else 0
        )
        
        for row_idx in range(max_rows):
            row1 = rows1[row_idx] if row_idx < len(rows1) else []
            row2 = rows2[row_idx] if row_idx < len(rows2) else []
            
            max_col = max(len(row1), len(row2))
            for col_idx in range(max_col):
                cell1 = row1[col_idx] if col_idx < len(row1) else ""
                cell2 = row2[col_idx] if col_idx < len(row2) else ""
                
                if cell1 != cell2:
                    changes.append({
                        "row": row_idx + 1,
                        "col": col_idx + 1,
                        "old_value": cell1,
                        "new_value": cell2
                    })
        
        return changes
    
    def _build_table_change_description(self, table1: Dict, table2: Dict, cell_changes: List[Dict]) -> str:
        """Строит описание изменений в таблице."""
        if not cell_changes:
            return "Изменения не обнаружены"
        
        descriptions = []
        for change in cell_changes[:5]:  # Ограничиваем количество
            old_val = change["old_value"][:50] if change["old_value"] else "пусто"
            new_val = change["new_value"][:50] if change["new_value"] else "пусто"
            descriptions.append(
                f"Строка {change['row']}, столбец {change['col']}: '{old_val}' изменено на '{new_val}'"
            )
        
        if len(cell_changes) > 5:
            descriptions.append(f"... и еще {len(cell_changes) - 5} изменений")
        
        return "; ".join(descriptions)
    
    def _compare_images(self):
        """Сравнение изображений в документах с названиями."""
        images1 = self.file1.get_images()
        images2 = self.file2.get_images()
        
        # Получаем названия изображений из предыдущих абзацев
        def get_image_name(img_index, paragraphs, is_first_doc):
            # Ищем абзац перед изображением с текстом "Рисунок"
            # Ищем в последних 10 абзацах перед предполагаемой позицией изображения
            search_start = max(0, img_index * 3 - 10)
            for para in reversed(paragraphs[search_start:img_index * 3]):
                text = para.get("text", "").strip()
                if text and ("рисунок" in text.lower() or "рис." in text.lower()):
                    return text
            return f"Рисунок {img_index}"
        
        # Создание карты изображений по хешу
        images1_map = {img.get("hash", f"img_{i}"): img for i, img in enumerate(images1)}
        images2_map = {img.get("hash", f"img_{i}"): img for i, img in enumerate(images2)}
        
        # Поиск идентичных изображений
        matched_hashes = set()
        for hash_val in images1_map:
            if hash_val in images2_map:
                matched_hashes.add(hash_val)
                img1 = images1_map[hash_val]
                img2 = images2_map[hash_val]
                name1 = get_image_name(img1["index"], self.file1.get_all_paragraphs(), True)
                name2 = get_image_name(img2["index"], self.file2.get_all_paragraphs(), False)
                self.image_changes.append({
                    "status": "identical",
                    "image_1_index": img1["index"],
                    "image_2_index": img2["index"],
                    "image_1_name": name1,
                    "image_2_name": name2,
                    "description": f"{name1} идентично {name2}",
                    "change_description": "Без изменений"
                })
        
        # Поиск удаленных изображений
        for img1 in images1:
            hash_val = img1.get("hash", f"img_{img1['index']}")
            if hash_val not in matched_hashes:
                name1 = get_image_name(img1["index"], self.file1.get_all_paragraphs(), True)
                self.image_changes.append({
                    "status": "deleted",
                    "image_1_index": img1["index"],
                    "image_2_index": None,
                    "image_1_name": name1,
                    "image_2_name": "",
                    "description": f"{name1} удалено",
                    "change_description": f"Удалено изображение: '{name1}'"
                })
        
        # Поиск добавленных изображений
        for img2 in images2:
            hash_val = img2.get("hash", f"img_{img2['index']}")
            if hash_val not in matched_hashes:
                name2 = get_image_name(img2["index"], self.file2.get_all_paragraphs(), False)
                self.image_changes.append({
                    "status": "added",
                    "image_1_index": None,
                    "image_2_index": img2["index"],
                    "image_1_name": "",
                    "image_2_name": name2,
                    "description": f"{name2} добавлено",
                    "change_description": f"Добавлено изображение: '{name2}'"
                })
        
        # Поиск измененных изображений (разные хеши, но похожие индексы)
        for img1 in images1:
            hash_val1 = img1.get("hash", f"img_{img1['index']}")
            if hash_val1 not in matched_hashes:
                for img2 in images2:
                    hash_val2 = img2.get("hash", f"img_{img2['index']}")
                    if hash_val2 not in matched_hashes and img1["index"] == img2["index"]:
                        name1 = get_image_name(img1["index"], self.file1.get_all_paragraphs(), True)
                        name2 = get_image_name(img2["index"], self.file2.get_all_paragraphs(), False)
                        self.image_changes.append({
                            "status": "modified",
                            "image_1_index": img1["index"],
                            "image_2_index": img2["index"],
                            "image_1_name": name1,
                            "image_2_name": name2,
                            "description": f"{name1} изменено",
                            "change_description": f"'{name1}' изменено на '{name2}'"
                        })
                        matched_hashes.add(hash_val1)
                        matched_hashes.add(hash_val2)
                        break
    
    def _build_change_description(self, result: Dict) -> str:
        """
        Построение текстового описания изменений в формате "старое" изменено на "новое".
        
        Args:
            result: Результат сравнения
            
        Returns:
            Текстовое описание изменений
        """
        status = result["status"]
        description_parts = []
        
        if status == "identical":
            return ""
        
        # Полный путь (приоритет новому документу, исключаем "Рисунок")
        full_path = result.get("full_path_2") or result.get("full_path_1") or ""
        if full_path and not full_path.lower().startswith('рисунок'):
            # Извлекаем номер пункта из пути
            path_match = re.search(r'Пункт\s+(\d+(?:\.\d+)*)', full_path)
            if path_match:
                point_num = path_match.group(1)
                description_parts.append(f"Пункт {point_num}")
            elif not any(x in full_path.lower() for x in ['рисунок', 'таблица', 'рис.', 'табл.']):
                description_parts.append(full_path)
        
        # Страница нового документа
        page_2 = result.get("page_2")
        if page_2:
            description_parts.append(f"страница {page_2}")
        
        # Описание изменений в формате "старое" изменено на "новое"
        text1 = result.get("text_1", "")
        text2 = result.get("text_2", "")
        
        if status == "added":
            description_parts.append("Добавлен новый абзац")
            if text2:
                # Полный текст или превью
                if len(text2) <= 200:
                    description_parts.append(f"'{text2}'")
                else:
                    text_preview = text2[:150].replace("\n", " ").strip() + "..."
                    description_parts.append(f"'{text_preview}'")
        
        elif status == "deleted":
            description_parts.append("Удален абзац")
            if text1:
                if len(text1) <= 200:
                    description_parts.append(f"'{text1}'")
                else:
                    text_preview = text1[:150].replace("\n", " ").strip() + "..."
                    description_parts.append(f"'{text_preview}'")
        
        elif status == "modified":
            if text1 and text2:
                # Нормализуем для сравнения
                norm1 = self._normalize_text(text1)
                norm2 = self._normalize_text(text2)
                
                # Если тексты идентичны после нормализации
                if norm1 == norm2:
                    description_parts.append("Изменено только форматирование")
                else:
                    # Поиск конкретных изменений
                    # Версии
                    version1_match = re.search(r'верси[ияею]\s+([\d.]+)', text1, re.IGNORECASE)
                    version2_match = re.search(r'верси[ияею]\s+([\d.]+)', text2, re.IGNORECASE)
                    
                    if version1_match and version2_match and version1_match.group(1) != version2_match.group(1):
                        description_parts.append(
                            f"'{version1_match.group(1)}' изменено на '{version2_match.group(1)}'"
                        )
                    else:
                        # Находим измененные фразы
                        words1 = norm1.split()
                        words2 = norm2.split()
                        matcher = difflib.SequenceMatcher(None, words1, words2)
                        opcodes = matcher.get_opcodes()
                        
                        changes_found = False
                        for tag, i1, i2, j1, j2 in opcodes:
                            if tag == 'replace' and not changes_found:
                                removed = ' '.join(words1[i1:i2])
                                added = ' '.join(words2[j1:j2])
                                if removed and added and len(removed) < 100 and len(added) < 100:
                                    description_parts.append(f"'{removed}' изменено на '{added}'")
                                    changes_found = True
                                    break
                        
                        if not changes_found:
                            # Если не нашли конкретных изменений, используем полные тексты
                            if len(text1) <= 100 and len(text2) <= 100:
                                description_parts.append(f"'{text1}' изменено на '{text2}'")
                            else:
                                description_parts.append("Текст абзаца изменен")
        
        return ". ".join(description_parts) if description_parts else ""
    
    def _determine_change_type(self, text1: str, text2: str) -> str:
        """
        Определение типа исправления.
        
        Args:
            text1: Первый текст
            text2: Второй текст
            
        Returns:
            Тип исправления
        """
        norm1 = self._normalize_text(text1)
        norm2 = self._normalize_text(text2)
        
        if norm1 == norm2:
            return "Изменение форматирования"
        
        # Проверка на грамматические изменения
        words1 = set(norm1.lower().split())
        words2 = set(norm2.lower().split())
        
        if words1 == words2:
            return "Изменение порядка слов"
        
        # Проверка на добавление/удаление текста
        added = words2 - words1
        removed = words1 - words2
        
        if added and not removed:
            return "Добавление текста"
        elif removed and not added:
            return "Удаление текста"
        elif len(added) > len(removed) * 2:
            return "Значительное добавление"
        elif len(removed) > len(added) * 2:
            return "Значительное удаление"
        else:
            return "Изменение содержания"
    
    def _get_text_changes(self, text1: str, text2: str) -> str:
        """
        Получение краткого описания изменений в тексте.
        Использует нормализованный текст для сравнения по содержимому.
        
        Args:
            text1: Первый текст
            text2: Второй текст
            
        Returns:
            Краткое описание изменений
        """
        # Нормализуем тексты для сравнения
        norm1 = self._normalize_text(text1)
        norm2 = self._normalize_text(text2)
        
        # Если тексты идентичны после нормализации
        if norm1 == norm2:
            return "изменено только форматирование"
        
        words1 = norm1.split()
        words2 = norm2.split()
        
        # Простое сравнение: что добавлено/удалено
        added_words = set(words2) - set(words1)
        removed_words = set(words1) - set(words2)
        
        changes = []
        if removed_words:
            sample_removed = list(removed_words)[:3]
            changes.append(f"удалено: {', '.join(sample_removed)}")
        if added_words:
            sample_added = list(added_words)[:3]
            changes.append(f"добавлено: {', '.join(sample_added)}")
        
        return "; ".join(changes) if changes else "текст изменен"
    
    def get_table_changes(self) -> List[Dict]:
        """
        Получить изменения в таблицах.
        
        Returns:
            Список изменений таблиц
        """
        return self.table_changes
    
    def get_image_changes(self) -> List[Dict]:
        """
        Получить изменения в изображениях.
        
        Returns:
            Список изменений изображений
        """
        return self.image_changes
    
    def get_statistics(self) -> Dict:
        """
        Получить статистику сравнения.
        
        Returns:
            Словарь со статистикой
        """
        total = len(self.comparison_results)
        identical = sum(1 for r in self.comparison_results if r["status"] == "identical")
        modified = sum(1 for r in self.comparison_results if r["status"] == "modified")
        added = sum(1 for r in self.comparison_results if r["status"] == "added")
        deleted = sum(1 for r in self.comparison_results if r["status"] == "deleted")
        
        return {
            "total": total,
            "identical": identical,
            "modified": modified,
            "added": added,
            "deleted": deleted,
            "identical_percent": (identical / total * 100) if total > 0 else 0,
            "modified_percent": (modified / total * 100) if total > 0 else 0,
            "added_percent": (added / total * 100) if total > 0 else 0,
            "deleted_percent": (deleted / total * 100) if total > 0 else 0,
            "tables_changed": len([t for t in self.table_changes if t["status"] != "identical"]),
            "images_changed": len([i for i in self.image_changes if i["status"] != "identical"])
        }

