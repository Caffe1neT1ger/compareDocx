"""
Модуль для сравнения двух DOCX документов.
Класс Compare предоставляет функционал для сравнения документов по абзацам.
"""

from typing import List, Dict, Tuple
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
    
    def _compare_documents(self):
        """Основной метод сравнения документов."""
        paragraphs1 = self.file1.get_all_paragraphs()
        paragraphs2 = self.file2.get_all_paragraphs()
        
        # Использование SequenceMatcher для поиска похожих абзацев
        matcher = difflib.SequenceMatcher(None, 
                                         [p["text"] for p in paragraphs1],
                                         [p["text"] for p in paragraphs2])
        
        # Получение совпадающих блоков
        matching_blocks = matcher.get_matching_blocks()
        
        # Создание карты соответствий
        matched_indices_1 = set()
        matched_indices_2 = set()
        
        for match in matching_blocks:
            if match.size > 0:
                for i in range(match.a, match.a + match.size):
                    matched_indices_1.add(i)
                for j in range(match.b, match.b + match.size):
                    matched_indices_2.add(j)
        
        # Обработка совпадающих абзацев
        match_map = {}
        for match in matching_blocks:
            if match.size > 0:
                for offset in range(match.size):
                    idx1 = match.a + offset
                    idx2 = match.b + offset
                    if idx1 < len(paragraphs1) and idx2 < len(paragraphs2):
                        match_map[idx1] = idx2
        
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
                "change_description": ""
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
                elif similarity >= 0.8:
                    result["status"] = "modified"
                    result["differences"] = self._get_differences(para1["text"], para2["text"])
                    result["change_description"] = self._build_change_description(result)
                else:
                    result["status"] = "modified"
                    result["differences"] = self._get_differences(para1["text"], para2["text"])
                    result["change_description"] = self._build_change_description(result)
            else:
                # Поиск наиболее похожего абзаца
                best_match_idx, best_similarity = self._find_best_match(
                    para1["text"], paragraphs2, matched_indices_2
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
                    matched_indices_2.add(best_match_idx)
                else:
                    result["change_description"] = self._build_change_description(result)
            
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
                self.comparison_results.append(result)
    
    def _calculate_similarity(self, text1: str, text2: str) -> float:
        """
        Вычисление схожести двух текстов.
        
        Args:
            text1: Первый текст
            text2: Второй текст
            
        Returns:
            Коэффициент схожести от 0.0 до 1.0
        """
        return difflib.SequenceMatcher(None, text1, text2).ratio()
    
    def _find_best_match(self, text: str, paragraphs: List[Dict], 
                        excluded_indices: set) -> Tuple[int, float]:
        """
        Поиск наиболее похожего абзаца.
        
        Args:
            text: Текст для поиска
            paragraphs: Список абзацев для поиска
            excluded_indices: Индексы, которые уже использованы
            
        Returns:
            Кортеж (индекс, схожесть)
        """
        best_idx = -1
        best_similarity = 0.0
        
        for idx, para in enumerate(paragraphs):
            if idx in excluded_indices:
                continue
            
            similarity = self._calculate_similarity(text, para["text"])
            if similarity > best_similarity:
                best_similarity = similarity
                best_idx = idx
        
        return best_idx, best_similarity
    
    def _get_differences(self, text1: str, text2: str) -> List[str]:
        """
        Получение списка различий между двумя текстами.
        
        Args:
            text1: Первый текст
            text2: Второй текст
            
        Returns:
            Список строк с описанием различий
        """
        differences = []
        
        # Использование difflib для получения различий
        diff = difflib.unified_diff(
            text1.splitlines(keepends=True),
            text2.splitlines(keepends=True),
            lineterm='',
            n=0
        )
        
        diff_list = list(diff)
        if len(diff_list) > 2:
            # Пропускаем заголовки
            for line in diff_list[2:]:
                if line.startswith('+') and not line.startswith('+++'):
                    differences.append(f"Добавлено: {line[1:].strip()}")
                elif line.startswith('-') and not line.startswith('---'):
                    differences.append(f"Удалено: {line[1:].strip()}")
        
        return differences[:5]  # Ограничиваем количество различий
    
    def get_comparison_results(self) -> List[Dict]:
        """
        Получить результаты сравнения.
        
        Returns:
            Список результатов сравнения
        """
        return self.comparison_results
    
    def _compare_tables(self):
        """Сравнение таблиц в документах."""
        tables1 = self.file1.get_tables()
        tables2 = self.file2.get_tables()
        
        # Создание карты таблиц по хешу
        tables1_map = {t["hash"]: t for t in tables1}
        tables2_map = {t["hash"]: t for t in tables2}
        
        # Поиск идентичных таблиц
        matched_hashes = set()
        for hash_val in tables1_map:
            if hash_val in tables2_map:
                matched_hashes.add(hash_val)
                self.table_changes.append({
                    "status": "identical",
                    "table_1_index": tables1_map[hash_val]["index"],
                    "table_2_index": tables2_map[hash_val]["index"],
                    "description": f"Таблица {tables1_map[hash_val]['index']} идентична таблице {tables2_map[hash_val]['index']}"
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
                    self.table_changes.append({
                        "status": "modified",
                        "table_1_index": t1["index"],
                        "table_2_index": best_match["index"],
                        "similarity": best_similarity,
                        "description": f"Таблица {t1['index']} изменена (схожесть {best_similarity*100:.1f}%)"
                    })
                    matched_hashes.add(best_match["hash"])
                else:
                    self.table_changes.append({
                        "status": "deleted",
                        "table_1_index": t1["index"],
                        "table_2_index": None,
                        "description": f"Таблица {t1['index']} удалена"
                    })
        
        # Поиск добавленных таблиц
        for t2 in tables2:
            if t2["hash"] not in matched_hashes:
                self.table_changes.append({
                    "status": "added",
                    "table_1_index": None,
                    "table_2_index": t2["index"],
                    "description": f"Таблица {t2['index']} добавлена"
                })
    
    def _compare_images(self):
        """Сравнение изображений в документах."""
        images1 = self.file1.get_images()
        images2 = self.file2.get_images()
        
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
                self.image_changes.append({
                    "status": "identical",
                    "image_1_index": img1["index"],
                    "image_2_index": img2["index"],
                    "description": f"Изображение {img1['index']} идентично изображению {img2['index']}"
                })
        
        # Поиск удаленных изображений
        for img1 in images1:
            hash_val = img1.get("hash", f"img_{img1['index']}")
            if hash_val not in matched_hashes:
                self.image_changes.append({
                    "status": "deleted",
                    "image_1_index": img1["index"],
                    "image_2_index": None,
                    "description": f"Изображение {img1['index']} удалено"
                })
        
        # Поиск добавленных изображений
        for img2 in images2:
            hash_val = img2.get("hash", f"img_{img2['index']}")
            if hash_val not in matched_hashes:
                self.image_changes.append({
                    "status": "added",
                    "image_1_index": None,
                    "image_2_index": img2["index"],
                    "description": f"Изображение {img2['index']} добавлено"
                })
    
    def _build_change_description(self, result: Dict) -> str:
        """
        Построение текстового описания изменений.
        
        Args:
            result: Результат сравнения
            
        Returns:
            Текстовое описание изменений
        """
        status = result["status"]
        description_parts = []
        
        if status == "identical":
            return ""
        
        # Полный путь (приоритет новому документу)
        full_path = result.get("full_path_2") or result.get("full_path_1") or ""
        if full_path:
            # Извлекаем номер пункта из пути
            path_match = re.search(r'Пункт\s+(\d+(?:\.\d+)*)', full_path)
            if path_match:
                point_num = path_match.group(1)
                description_parts.append(f"Пункт {point_num}")
            else:
                description_parts.append(full_path)
        
        # Страница нового документа
        page_2 = result.get("page_2")
        if page_2:
            description_parts.append(f"страница {page_2}")
        
        # Описание изменений
        if status == "added":
            description_parts.append("Добавлен новый абзац")
            if result.get("text_2"):
                # Извлечение ключевой информации из текста
                text2 = result["text_2"]
                # Поиск важных изменений (версии, даты, числа)
                version_match = re.search(r'верси[ияею]\s+([\d.]+)', text2, re.IGNORECASE)
                if version_match:
                    description_parts.append(f"версия системы {version_match.group(1)}")
                else:
                    text_preview = text2[:80].replace("\n", " ").strip()
                    if len(text2) > 80:
                        text_preview += "..."
                    if text_preview:
                        description_parts.append(f"'{text_preview}'")
        
        elif status == "deleted":
            description_parts.append("Удален абзац")
            if result.get("text_1"):
                text1 = result["text_1"]
                text_preview = text1[:80].replace("\n", " ").strip()
                if len(text1) > 80:
                    text_preview += "..."
                if text_preview:
                    description_parts.append(f"'{text_preview}'")
        
        elif status == "modified":
            description_parts.append("Изменен абзац")
            
            # Анализ конкретных изменений
            text1 = result.get("text_1", "")
            text2 = result.get("text_2", "")
            
            if text1 and text2:
                # Поиск изменений версий
                version1_match = re.search(r'верси[ияею]\s+([\d.]+)', text1, re.IGNORECASE)
                version2_match = re.search(r'верси[ияею]\s+([\d.]+)', text2, re.IGNORECASE)
                
                if version1_match and version2_match:
                    if version1_match.group(1) != version2_match.group(1):
                        description_parts.append(
                            f"Изменена версия системы с {version1_match.group(1)} на {version2_match.group(1)}"
                        )
                
                # Поиск других изменений
                diff_info = self._get_text_changes(text1, text2)
                if diff_info and not (version1_match and version2_match):
                    description_parts.append(f"Изменения: {diff_info}")
        
        return ". ".join(description_parts) if description_parts else ""
    
    def _get_text_changes(self, text1: str, text2: str) -> str:
        """
        Получение краткого описания изменений в тексте.
        
        Args:
            text1: Первый текст
            text2: Второй текст
            
        Returns:
            Краткое описание изменений
        """
        words1 = text1.split()
        words2 = text2.split()
        
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

