"""
Модуль для сравнения двух DOCX документов.

Класс Compare предоставляет функционал для:
- Сравнения документов по содержимому (игнорируя стили форматирования)
- Нормализации текстов перед сравнением
- Поиска идентичных и похожих абзацев
- Сравнения таблиц с детальным анализом изменений в ячейках
- Сравнения изображений
- Построения текстовых описаний изменений
- Определения типов исправлений

Особенности:
- Сравнение по содержимому, а не по стилям
- Поддержка кастомных стилей
- Детальный анализ изменений в таблицах (строка, столбец)
- Полные тексты в описании различий
"""

from typing import List, Dict, Tuple, Set, Optional
from docx_file import DocxFile
import difflib
import re
from config import config
from logger_config import logger
from exceptions import ComparisonError


class Compare:
    """
    Класс для сравнения двух DOCX документов.
    
    Выполняет полное сравнение документов:
    - Абзацы сравниваются по содержимому (нормализованный текст)
    - Таблицы сравниваются по содержимому ячеек
    - Изображения сравниваются по хешу
    - Строятся детальные описания всех изменений
    """
    
    def __init__(self, file1_path: str, file2_path: str, llm_adapter=None):
        """
        Инициализация класса сравнения.
        
        При инициализации автоматически выполняется:
        1. Загрузка и парсинг обоих документов
        2. Сравнение абзацев
        3. Сравнение таблиц
        4. Сравнение изображений
        5. Дополнительный анализ изменений через LLM (если адаптер предоставлен)
        
        Args:
            file1_path: Путь к первому DOCX файлу (базовый документ)
            file2_path: Путь ко второму DOCX файлу (измененный документ)
            llm_adapter: Опциональный адаптер LLM для дополнительного анализа изменений
        """
        # Загрузка документов
        self.file1 = DocxFile(file1_path)  # Базовый документ
        self.file2 = DocxFile(file2_path)  # Измененный документ
        
        # LLM адаптер для дополнительного анализа
        self.llm_adapter = llm_adapter
        
        # Результаты сравнения
        self.comparison_results = []  # Результаты сравнения абзацев
        self.table_changes = []  # Результаты сравнения таблиц
        self.image_changes = []  # Результаты сравнения изображений
        
        # Автоматическое выполнение сравнения при инициализации
        try:
            logger.info("Начало сравнения документов")
            self._compare_documents()  # Сравнение абзацев
            logger.debug(f"Сравнение абзацев завершено: {len(self.comparison_results)} результатов")
            
            self._compare_tables()  # Сравнение таблиц
            logger.debug(f"Сравнение таблиц завершено: {len(self.table_changes)} результатов")
            
            self._compare_images()  # Сравнение изображений
            logger.debug(f"Сравнение изображений завершено: {len(self.image_changes)} результатов")
            
            # Дополнительный анализ через LLM для измененных элементов
            if self.llm_adapter and self.llm_adapter.is_enabled():
                logger.info("Начало анализа изменений через LLM")
                self._analyze_changes_with_llm()
                logger.info("Анализ изменений через LLM завершен")
            
            logger.info("Сравнение документов завершено успешно")
        except Exception as e:
            logger.error(f"Ошибка при сравнении документов: {e}")
            raise ComparisonError(str(e))
    
    def _normalize_text(self, text: str) -> str:
        """
        Нормализация текста для сравнения по содержимому.
        
        Убирает различия в форматировании, оставляя только содержимое:
        - Множественные пробелы → один пробел
        - Переносы строк → пробелы
        - Специальные пробелы → обычные пробелы
        - Лишние пробелы в начале/конце удаляются
        
        Это позволяет сравнивать тексты независимо от стилей форматирования.
        
        Args:
            text: Исходный текст
            
        Returns:
            Нормализованный текст для сравнения
            
        Пример:
            "Настоящее   техническое\nзадание" → "Настоящее техническое задание"
        """
        if not text:
            return ""
        
        # Примечание: регистр сохраняется для точности сравнения
        # Используется настройка из конфигурации
        if config.comparison.normalize_case:
            text = text.lower()
        
        # Шаг 1: Замена множественных пробелов на один
        text = re.sub(r'\s+', ' ', text)
        
        # Шаг 2: Удаление пробелов в начале и конце
        text = text.strip()
        
        # Шаг 3: Нормализация переносов строк (замена на пробелы)
        text = re.sub(r'[\n\r]+', ' ', text)
        
        # Шаг 4: Удаление специальных символов форматирования
        text = text.replace('\xa0', ' ')  # Неразрывный пробел (U+00A0)
        text = text.replace('\u2009', ' ')  # Тонкий пробел (U+2009)
        text = text.replace('\u2008', ' ')  # Пунктуационный пробел (U+2008)
        
        # Шаг 5: Финальная очистка множественных пробелов после нормализации
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
        
        # Берем первые и последние слова согласно конфигурации
        first_words = ' '.join(words[:config.comparison.fingerprint_first_words])
        last_words = ' '.join(words[-config.comparison.fingerprint_last_words:]) if len(words) > config.comparison.fingerprint_last_words else ''
        
        # Также используем длину и количество слов
        length = len(normalized)
        word_count = len(words)
        
        return f"{first_words}|{last_words}|{length}|{word_count}"
    
    def _compare_documents(self):
        """
        Основной метод сравнения документов.
        
        Алгоритм сравнения (6 шагов):
        1. Нормализация всех текстов (игнорируя стили)
        2. Создание отпечатков для быстрого поиска
        3. Сопоставление по позиции (SequenceMatcher)
        4. Дополнительный поиск по отпечаткам (независимо от позиции)
        5. Анализ каждого абзаца и построение описаний изменений
        6. Обработка добавленных/удаленных абзацев
        
        Сравнение выполняется по содержимому текста, игнорируя стили форматирования.
        """
        # Получение всех абзацев из обоих документов
        paragraphs1 = self.file1.get_all_paragraphs()
        paragraphs2 = self.file2.get_all_paragraphs()
        
        # Шаг 1: Нормализация текстов для сравнения (игнорируя стили)
        # Это позволяет сравнивать тексты независимо от форматирования
        normalized_texts1 = [self._normalize_text(p["text"]) for p in paragraphs1]
        normalized_texts2 = [self._normalize_text(p["text"]) for p in paragraphs2]
        
        # Шаг 2: Создание отпечатков для быстрого поиска идентичных абзацев
        # Отпечаток = первые/последние слова + длина + количество слов
        fingerprints1 = {i: self._get_text_fingerprint(p["text"]) for i, p in enumerate(paragraphs1)}
        fingerprints2 = {i: self._get_text_fingerprint(p["text"]) for i, p in enumerate(paragraphs2)}
        
        # Шаг 3: Создание индекса по отпечаткам для второго документа
        # Это позволяет быстро находить абзацы с одинаковыми отпечатками
        fingerprint_index2 = {}
        for idx2, fp in fingerprints2.items():
            if fp:  # Игнорируем пустые отпечатки
                if fp not in fingerprint_index2:
                    fingerprint_index2[fp] = []
                fingerprint_index2[fp].append(idx2)
        
        # Шаг 4: Использование SequenceMatcher для поиска похожих последовательностей
        # Находит блоки идентичных абзацев по позиции
        matcher = difflib.SequenceMatcher(None, normalized_texts1, normalized_texts2)
        matching_blocks = matcher.get_matching_blocks()
        
        # Шаг 5: Создание карты соответствий по позиции
        matched_indices_1 = set()  # Индексы абзацев из первого документа, которые уже сопоставлены
        matched_indices_2 = set()  # Индексы абзацев из второго документа, которые уже сопоставлены
        match_map = {}  # Прямое соответствие: индекс1 -> индекс2
        
        # Обработка совпадающих блоков (абзацы на одинаковых позициях)
        for match in matching_blocks:
            if match.size > 0:
                for offset in range(match.size):
                    idx1 = match.a + offset
                    idx2 = match.b + offset
                    if idx1 < len(paragraphs1) and idx2 < len(paragraphs2):
                        matched_indices_1.add(idx1)
                        matched_indices_2.add(idx2)
                        match_map[idx1] = idx2
        
        # Шаг 6: Дополнительный поиск идентичных абзацев по отпечаткам (независимо от позиции)
        # Это помогает найти одинаковые абзацы, которые находятся в разных местах документа
        for idx1, fp1 in fingerprints1.items():
            if fp1 and fp1 in fingerprint_index2:
                # Найден абзац с таким же отпечатком
                candidates = fingerprint_index2[fp1]
                for idx2 in candidates:
                    if idx2 not in matched_indices_2:
                        # Проверяем, действительно ли тексты идентичны
                        norm1 = normalized_texts1[idx1]
                        norm2 = normalized_texts2[idx2]
                        similarity_check = self._calculate_similarity(norm1, norm2)
                        if norm1 == norm2 or similarity_check >= config.comparison.similarity_threshold_high:
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
                "change_type": "",  # Тип исправления
                "llm_response": ""  # Ответ от LLM
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
                
                # Использование порогов из конфигурации
                if similarity >= config.comparison.similarity_threshold_identical:
                    result["status"] = "identical"
                    result["change_type"] = "Без изменений"
                elif similarity >= config.comparison.similarity_threshold_medium:
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
                
                if best_similarity >= config.comparison.similarity_threshold_low:
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
                    "change_description": "",
                    "llm_response": ""  # Ответ от LLM
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
        """
        Находит изменения в конкретных ячейках таблиц.
        
        Сравнивает таблицы построчно и по столбцам, находя все измененные ячейки.
        Возвращает список изменений с указанием строки и столбца.
        
        Args:
            table1: Данные первой таблицы (словарь с ключом "rows")
            table2: Данные второй таблицы (словарь с ключом "rows")
            
        Returns:
            Список словарей с изменениями:
            [
                {"row": 2, "col": 1, "old_value": "1000", "new_value": "1500"},
                ...
            ]
        """
        changes = []
        rows1 = table1.get("rows", [])
        rows2 = table2.get("rows", [])
        
        # Определение максимальных размеров для обработки всех ячеек
        max_rows = max(len(rows1), len(rows2))
        max_cols = max(
            max(len(row) for row in rows1) if rows1 else 0,
            max(len(row) for row in rows2) if rows2 else 0
        )
        
        # Построчное сравнение ячеек
        for row_idx in range(max_rows):
            row1 = rows1[row_idx] if row_idx < len(rows1) else []
            row2 = rows2[row_idx] if row_idx < len(rows2) else []
            
            max_col = max(len(row1), len(row2))
            for col_idx in range(max_col):
                # Получение содержимого ячеек (пустая строка, если ячейка отсутствует)
                cell1 = row1[col_idx] if col_idx < len(row1) else ""
                cell2 = row2[col_idx] if col_idx < len(row2) else ""
                
                # Если содержимое ячеек различается - это изменение
                if cell1 != cell2:
                    changes.append({
                        "row": row_idx + 1,  # Номер строки (начиная с 1)
                        "col": col_idx + 1,  # Номер столбца (начиная с 1)
                        "old_value": cell1,  # Старое значение
                        "new_value": cell2   # Новое значение
                    })
        
        return changes
    
    def _build_table_change_description(self, table1: Dict, table2: Dict, cell_changes: List[Dict]) -> str:
        """
        Строит текстовое описание изменений в таблице.
        
        Формат описания: "Строка X, столбец Y: 'старое' изменено на 'новое'"
        Ограничивает количество изменений для читаемости (первые 5).
        
        Args:
            table1: Данные первой таблицы
            table2: Данные второй таблицы
            cell_changes: Список изменений в ячейках
            
        Returns:
            Текстовое описание изменений в формате "'старое' изменено на 'новое'"
            
        Пример:
            "Строка 2, столбец 1: '1000' изменено на '1500'; Строка 3, столбец 2: 'текст' изменено на 'новый текст'"
        """
        if not cell_changes:
            return "Изменения не обнаружены"
        
        descriptions = []
        # Ограничиваем количество для читаемости (используем конфигурацию)
        max_display = config.excel.max_differences_display
        max_length = config.excel.max_cell_value_length
        
        for change in cell_changes[:max_display]:
            # Обрезаем длинные значения для компактности
            old_val = change["old_value"][:max_length] if change["old_value"] else "пусто"
            new_val = change["new_value"][:max_length] if change["new_value"] else "пусто"
            descriptions.append(
                f"Строка {change['row']}, столбец {change['col']}: '{old_val}' изменено на '{new_val}'"
            )
        
        # Если изменений больше максимума, добавляем информацию об остальных
        if len(cell_changes) > max_display:
            descriptions.append(f"... и еще {len(cell_changes) - max_display} изменений")
        
        return "; ".join(descriptions)
    
    def _compare_images(self):
        """Сравнение изображений в документах с названиями."""
        images1 = self.file1.get_images()
        images2 = self.file2.get_images()
        
        # Получаем названия изображений из предыдущих абзацев
        def get_image_name(img_index, paragraphs, is_first_doc):
            # Ищем абзац перед изображением с текстом "Рисунок"
            # Ищем в последних N абзацах перед предполагаемой позицией изображения
            search_backward = config.document.search_backward_paragraphs
            search_start = max(0, img_index * 3 - search_backward)
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
    
    def _analyze_changes_with_llm(self):
        """
        Дополнительный анализ изменений через LLM для измененных, добавленных и удаленных элементов.
        
        Вызывается автоматически после основного сравнения, если LLM адаптер доступен.
        Анализирует только элементы с изменениями (modified, added, deleted),
        пропуская идентичные элементы для оптимизации.
        """
        if not self.llm_adapter or not self.llm_adapter.is_enabled():
            return
        
        print("\nВыполнение дополнительного анализа изменений через LLM...")
        
        # Фильтруем только элементы с изменениями
        changed_results = [
            r for r in self.comparison_results 
            if r["status"] in ["modified", "added", "deleted"]
        ]
        
        total_changed = len(changed_results)
        if total_changed == 0:
            print("Нет изменений для анализа через LLM.")
            return
        
        print(f"Анализ {total_changed} измененных элементов...")
        
        # Анализ каждого измененного элемента
        for idx, result in enumerate(changed_results, 1):
            if idx % 10 == 0:
                print(f"Обработано {idx}/{total_changed} элементов...")
            
            status = result["status"]
            text1 = result.get("text_1", "") or ""
            text2 = result.get("text_2", "") or ""
            
            # Формирование контекста для LLM
            context_parts = []
            full_path = result.get("full_path_2") or result.get("full_path_1") or ""
            if full_path:
                context_parts.append(f"Путь: {full_path}")
            
            page = result.get("page_2") or result.get("page_1")
            if page:
                context_parts.append(f"Страница: {page}")
            
            context = "; ".join(context_parts) if context_parts else None
            
            # Анализ в зависимости от статуса
            if status == "modified" and text1 and text2:
                # Для измененных элементов - сравниваем оба текста
                llm_response = self.llm_adapter.analyze_changes(text1, text2, context)
                result["llm_response"] = llm_response
                
            elif status == "added" and text2:
                # Для добавленных элементов - анализируем новый текст
                llm_response = self.llm_adapter.analyze_changes("", text2, context)
                result["llm_response"] = llm_response
                
            elif status == "deleted" and text1:
                # Для удаленных элементов - анализируем старый текст
                llm_response = self.llm_adapter.analyze_changes(text1, "", context)
                result["llm_response"] = llm_response
        
        print(f"Анализ через LLM завершен. Обработано {total_changed} элементов.")
    
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
        
        # Статистика по LLM анализу
        llm_analyzed = sum(1 for r in self.comparison_results if r.get("llm_response"))
        
        # Статистика по таблицам
        tables1 = self.file1.get_tables()
        tables2 = self.file2.get_tables()
        tables_changed = len([t for t in self.table_changes if t["status"] != "identical"])
        
        # Статистика по изображениям
        images1 = self.file1.get_images()
        images2 = self.file2.get_images()
        images_changed = len([i for i in self.image_changes if i["status"] != "identical"])
        
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
            "tables_total_1": len(tables1),
            "tables_total_2": len(tables2),
            "tables_changed": tables_changed,
            "images_total_1": len(images1),
            "images_total_2": len(images2),
            "images_changed": images_changed,
            "llm_analyzed": llm_analyzed
        }

