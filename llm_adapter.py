"""
Модуль для интеграции с OpenAI LLM для дополнительного анализа изменений.

Класс LLMAdapter предоставляет функционал для:
- Отправки запросов к OpenAI API с двумя фрагментами текста
- Получения ответов в деловом стиле об изменениях
- Обработки ошибок и таймаутов
- Опционального использования (если API ключ не задан, LLM не используется)
- Чтения промптов из файлов в папке prompts/
- Чтения конфигурации из .env файла

Требования:
- Установленный пакет openai: pip install openai
- Установленный пакет python-dotenv: pip install python-dotenv
- API ключ OpenAI (в .env файле или переменной окружения)
"""

from typing import Optional, Dict
import os
import time
import re
from pathlib import Path

# Попытка загрузить переменные окружения из .env файла
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # Если python-dotenv не установлен, работаем только с переменными окружения системы
    pass

from config import config
from logger_config import logger
from exceptions import LLMError


def _remove_markdown_bold(text: str) -> str:
    """
    Удаляет markdown форматирование жирного текста (**текст**) из строки.
    
    Args:
        text: Текст с возможным markdown форматированием
        
    Returns:
        Текст без markdown форматирования жирного текста
    """
    if not text:
        return text
    # Удаляем **текст** и заменяем на просто текст
    # Паттерн для **текст** или ** текст **
    text = re.sub(r'\*\*([^*]+?)\*\*', r'\1', text)
    return text


class LLMAdapter:
    """
    Адаптер для работы с OpenAI LLM для анализа изменений в документах.
    
    Отправляет два фрагмента текста (из старого и нового документа) в LLM
    и получает ответ в деловом стиле о характере изменений.
    
    Особенности:
    - Использует системный промпт для делового стиля ответов
    - Обрабатывает ошибки и таймауты
    - Может работать без API ключа (возвращает пустые ответы)
    """
    
    def _load_system_prompt(self) -> str:
        """
        Загрузка системного промпта из файла.
        
        Returns:
            Системный промпт из файла или дефолтный промпт, если файл не найден
        """
        prompt_file = Path(__file__).parent / "prompts" / "system_prompt.txt"
        
        try:
            if prompt_file.exists():
                with open(prompt_file, 'r', encoding='utf-8') as f:
                    return f.read().strip()
            else:
                print(f"Предупреждение: файл промпта не найден: {prompt_file}")
                print("Используется дефолтный промпт.")
                return self._get_default_system_prompt()
        except Exception as e:
            print(f"Ошибка при загрузке промпта из файла: {e}")
            print("Используется дефолтный промпт.")
            return self._get_default_system_prompt()
    
    def _get_default_system_prompt(self) -> str:
        """
        Возвращает дефолтный системный промпт (fallback).
        
        Returns:
            Дефолтный системный промпт
        """
        return """Вы - профессиональный аналитик документов, специализирующийся на сравнении технических и деловых документов.

Ваша задача - проанализировать два фрагмента текста из разных версий документа и предоставить краткое, точное описание изменений в деловом стиле.

Требования к ответу:
1. Используйте деловой, формальный стиль изложения
2. Будьте конкретны и точны в описании изменений
3. Укажите, что именно изменилось (текст, данные, формулировки)
4. Если изменения незначительны (опечатки, пунктуация), укажите это
5. Если изменения существенны, опишите их характер (добавление информации, изменение формулировки, обновление данных)
6. Ответ должен быть кратким (1-3 предложения) и информативным
7. Избегайте общих фраз, будьте конкретны
8. Используйте профессиональную терминологию

Примеры хороших ответов:
- "Обновлена версия системы с 2.0.3 на 2.0.4 в описании технических характеристик"
- "Добавлено уточнение о требованиях к производительности сервера"
- "Изменена формулировка пункта о сроках выполнения работ с 'не более 30 дней' на 'в течение 30 рабочих дней'"
- "Исправлены опечатки и пунктуационные ошибки в тексте"

Примеры плохих ответов:
- "Текст изменен" (слишком общо)
- "Есть различия" (неинформативно)
- "Что-то поменялось" (неформально)

Проанализируйте предоставленные фрагменты текста и опишите изменения в соответствии с требованиями выше."""
    
    def _load_user_prompt_template(self) -> str:
        """
        Загрузка шаблона пользовательского промпта из файла.
        
        Returns:
            Шаблон пользовательского промпта из файла или дефолтный шаблон
        """
        template_file = Path(__file__).parent / "prompts" / "user_prompt_template.txt"
        
        try:
            if template_file.exists():
                with open(template_file, 'r', encoding='utf-8') as f:
                    return f.read().strip()
            else:
                return self._get_default_user_prompt_template()
        except Exception as e:
            print(f"Ошибка при загрузке шаблона промпта из файла: {e}")
            return self._get_default_user_prompt_template()
    
    def _get_default_user_prompt_template(self) -> str:
        """
        Возвращает дефолтный шаблон пользовательского промпта.
        
        Returns:
            Дефолтный шаблон пользовательского промпта
        """
        return """Проанализируйте изменения между двумя фрагментами текста из разных версий документа.

Старый текст:
{old_text}

Новый текст:
{new_text}
{context_section}

Опишите изменения в деловом стиле согласно требованиям системного промпта."""

    def __init__(self, api_key: Optional[str] = None, api_url: Optional[str] = None,
                 model: Optional[str] = None, temperature: Optional[float] = None, 
                 max_tokens: Optional[int] = None):
        """
        Инициализация адаптера LLM.
        
        Параметры могут быть заданы через аргументы, переменные окружения или .env файл.
        Приоритет: аргументы > переменные окружения > .env файл > значения по умолчанию
        
        Args:
            api_key: API ключ OpenAI. Если не указан, будет использоваться OPENAI_API_KEY из .env или окружения.
            api_url: URL API. Если не указан, будет использоваться OPENAI_API_URL из .env или стандартный URL OpenAI.
            model: Название модели OpenAI. Если не указан, будет использоваться OPENAI_MODEL из .env или gpt-3.5-turbo.
            temperature: Температура для генерации (0.0-1.0). Если не указан, будет использоваться OPENAI_TEMPERATURE из .env или 0.3.
            max_tokens: Максимальное количество токенов в ответе. Если не указан, будет использоваться OPENAI_MAX_TOKENS из .env или 200.
        """
        # Загрузка конфигурации из .env или переменных окружения
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        self.api_url = api_url or os.getenv("OPENAI_API_URL") or None
        model_name = model or os.getenv("OPENAI_MODEL", "gpt-3.5-turbo")
        # Нормализация имени модели: замена обратных слэшей на прямые (для Windows)
        self.model = model_name.replace("\\", "/") if model_name else "gpt-3.5-turbo"
        
        # Парсинг числовых значений из окружения
        try:
            self.temperature = float(temperature) if temperature is not None else float(os.getenv("OPENAI_TEMPERATURE", "0.3"))
        except (ValueError, TypeError):
            self.temperature = 0.3
        
        try:
            self.max_tokens = int(max_tokens) if max_tokens is not None else int(os.getenv("OPENAI_MAX_TOKENS", "200"))
        except (ValueError, TypeError):
            self.max_tokens = 200
        
        # Загрузка промптов из файлов
        self.system_prompt = self._load_system_prompt()
        self.user_prompt_template = self._load_user_prompt_template()
        
        self.client = None
        self.enabled = False
        
        # Попытка инициализации клиента OpenAI
        if self.api_key:
            try:
                from openai import OpenAI
                
                # Создание клиента с кастомным URL, если указан
                client_kwargs = {"api_key": self.api_key}
                if self.api_url:
                    client_kwargs["base_url"] = self.api_url
                
                self.client = OpenAI(**client_kwargs)
                self.enabled = True
            except ImportError:
                print("Предупреждение: пакет 'openai' не установлен. LLM функции будут отключены.")
                print("Установите пакет: pip install openai")
            except Exception as e:
                print(f"Предупреждение: не удалось инициализировать OpenAI клиент: {e}")
                print("LLM функции будут отключены.")
        else:
            print("Предупреждение: API ключ OpenAI не задан. LLM функции будут отключены.")
            print("Установите переменную окружения OPENAI_API_KEY или создайте файл .env с настройками.")
    
    def analyze_changes(self, old_text: str, new_text: str, 
                       context: Optional[str] = None) -> str:
        """
        Анализ изменений между двумя фрагментами текста с помощью LLM.
        
        Отправляет запрос к OpenAI API с двумя фрагментами текста и получает
        ответ в деловом стиле о характере изменений.
        
        Включает retry логику и таймауты для повышения надежности.
        
        Args:
            old_text: Текст из старого документа (базовая версия)
            new_text: Текст из нового документа (измененная версия)
            context: Дополнительный контекст в формате "Путь: ...; Страница: ..."
                    Может быть использован для более точного анализа
        
        Returns:
            Ответ LLM в деловом стиле о характере изменений с путем в начале.
            Формат: "Путь > Подпуть. страница X. [ответ LLM]"
            Если LLM недоступен или произошла ошибка, возвращает пустую строку.
        """
        if not self.enabled or not self.client:
            return ""
        
        # Извлечение пути и страницы из context для добавления в начало ответа
        path_prefix = ""
        if context:
            # Парсим context для извлечения пути и страницы
            path_part = None
            page_part = None
            
            if "Путь:" in context:
                path_start = context.find("Путь:") + len("Путь:")
                path_end = context.find(";", path_start)
                if path_end == -1:
                    path_end = len(context)
                path_part = context[path_start:path_end].strip()
            
            if "Страница:" in context:
                page_start = context.find("Страница:") + len("Страница:")
                page_part = context[page_start:].strip()
            
            # Формируем префикс пути
            if path_part:
                # Заменяем разделители на " > " для единообразия
                path_formatted = path_part.replace(" > ", " > ").replace(" → ", " > ")
                path_prefix = f"{path_formatted}"
                if page_part:
                    path_prefix += f". страница {page_part}."
                else:
                    path_prefix += "."
                path_prefix += "\n\n"  # Пустая строка между путем и текстом
        
        # Формирование пользовательского промпта из шаблона
        context_section = f"\n\nКонтекст: {context}" if context else ""
        user_prompt = self.user_prompt_template.format(
            old_text=old_text,
            new_text=new_text,
            context_section=context_section
        )
        
        # Retry логика с экспоненциальной задержкой
        max_retries = config.llm.max_retries
        retry_delay = config.llm.retry_delay_seconds
        timeout = config.llm.timeout_seconds
        
        for attempt in range(max_retries):
            try:
                # Отправка запроса к OpenAI API с таймаутом
                # Поддержка дополнительных параметров для совместимости с различными провайдерами
                request_params = {
                    "model": self.model,
                    "messages": [
                        {"role": "system", "content": self.system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    "temperature": self.temperature,
                    "max_tokens": self.max_tokens,
                    "timeout": timeout
                }
                
                # Дополнительные параметры из переменных окружения (для Cloud.ru и других провайдеров)
                presence_penalty = os.getenv("OPENAI_PRESENCE_PENALTY")
                if presence_penalty:
                    try:
                        request_params["presence_penalty"] = float(presence_penalty)
                    except (ValueError, TypeError):
                        pass
                
                top_p = os.getenv("OPENAI_TOP_P")
                if top_p:
                    try:
                        request_params["top_p"] = float(top_p)
                    except (ValueError, TypeError):
                        pass
                
                response = self.client.chat.completions.create(**request_params)
                
                # Извлечение ответа
                if response.choices and len(response.choices) > 0:
                    content = response.choices[0].message.content
                    if content:
                        llm_response = content.strip()
                        # Убираем markdown форматирование жирного текста
                        llm_response = _remove_markdown_bold(llm_response)
                        # Добавляем путь в начало ответа с пустой строкой, если он есть
                        if path_prefix:
                            return f"{path_prefix}{llm_response}"
                        return llm_response
                    else:
                        return "Без изменений"
                else:
                    return "Без изменений"
                    
            except Exception as e:
                error_msg = str(e)
                logger.warning(f"Попытка {attempt + 1}/{max_retries} не удалась: {error_msg}")
                
                # Если это ошибка модели (422), не повторяем запросы
                if "422" in error_msg or "Invalid parameter" in error_msg or "model" in error_msg.lower():
                    logger.error(f"Ошибка модели: {error_msg}")
                    logger.error(f"Используемая модель: {self.model}")
                    logger.error(f"Проверьте правильность названия модели в .env файле (OPENAI_MODEL)")
                    logger.error(f"Для стандартного OpenAI API используйте: gpt-3.5-turbo, gpt-4, gpt-4-turbo-preview")
                    return ""
                
                # Если это последняя попытка, логируем ошибку и возвращаем пустую строку
                if attempt == max_retries - 1:
                    logger.error(f"Все попытки исчерпаны. Ошибка LLM: {error_msg}")
                    return ""
                
                # Экспоненциальная задержка перед следующей попыткой
                delay = retry_delay * (2 ** attempt)
                logger.debug(f"Ожидание {delay} секунд перед следующей попыткой...")
                time.sleep(delay)
        
        return ""
    
    def analyze_multiple_changes(self, text_pairs: list[tuple[str, str]], 
                                contexts: Optional[list[str]] = None) -> list[str]:
        """
        Анализ множественных изменений за один запрос (батчинг).
        
        Может быть использован для оптимизации, если API поддерживает множественные запросы.
        В текущей реализации выполняет последовательные запросы.
        
        Args:
            text_pairs: Список кортежей (old_text, new_text)
            contexts: Опциональный список контекстов для каждой пары
        
        Returns:
            Список ответов LLM для каждой пары текстов
        """
        if not self.enabled or not self.client:
            return [""] * len(text_pairs)
        
        results = []
        contexts = contexts or [None] * len(text_pairs)
        
        for (old_text, new_text), context in zip(text_pairs, contexts):
            result = self.analyze_changes(old_text, new_text, context)
            results.append(result)
        
        return results
    
    def is_enabled(self) -> bool:
        """
        Проверка, доступен ли LLM адаптер.
        
        Returns:
            True, если LLM доступен и может использоваться
        """
        return self.enabled
    
    def get_model_info(self) -> Dict[str, str]:
        """
        Получение информации о настройках модели.
        
        Returns:
            Словарь с информацией о модели и настройках
        """
        return {
            "model": self.model,
            "temperature": str(self.temperature),
            "max_tokens": str(self.max_tokens),
            "enabled": str(self.enabled)
        }
    
    def generate_summary(self, llm_responses: list) -> str:
        """
        Генерация краткого смыслового описания всех изменений на основе LLM ответов.
        
        Собирает все LLM ответы, группирует их и отправляет к LLM для генерации
        краткого описания в формате нумерованного списка.
        
        Args:
            llm_responses: Список словарей с ключами "response" (текст ответа) и "page" (номер страницы),
                          или список строк (для обратной совместимости)
        
        Returns:
            Краткое смысловое описание изменений в формате нумерованного списка
        """
        if not self.enabled or not self.client:
            return "Общие правки."
        
        # Обрабатываем входные данные: могут быть словари или строки
        processed_responses = []
        for resp in llm_responses:
            if isinstance(resp, dict):
                # Новый формат: словарь с "response" и "page"
                response_text = resp.get("response", "")
                page = resp.get("page")
            else:
                # Старый формат: просто строка (для обратной совместимости)
                response_text = resp
                page = None
            
            # Фильтруем пустые ответы и "Без изменений"
            if response_text and response_text.strip() and response_text.strip() != "Без изменений":
                processed_responses.append({
                    "response": response_text,
                    "page": page
                })
        
        logger.debug(f"Получено {len(llm_responses)} LLM ответов, после фильтрации: {len(processed_responses)}")
        
        if not processed_responses:
            logger.warning("Нет LLM ответов для генерации краткого описания")
            return "Общие правки."
        
        # Загружаем промпт для краткого описания
        summary_prompt_file = Path(__file__).parent / "prompts" / "summary_prompt.txt"
        try:
            if summary_prompt_file.exists():
                with open(summary_prompt_file, 'r', encoding='utf-8') as f:
                    summary_system_prompt = f.read().strip()
            else:
                # Дефолтный промпт для краткого описания
                summary_system_prompt = """Вы - профессиональный аналитик документов. 
Проанализируйте список изменений и составьте краткое смысловое описание в формате нумерованного списка.
Группируйте похожие изменения вместе. Указывайте конкретные места изменений (разделы, пункты, таблицы) с номерами страниц.
Если есть общие правки, укажите их отдельным пунктом."""
        except Exception as e:
            logger.warning(f"Ошибка при загрузке промпта краткого описания: {e}")
            summary_system_prompt = "Проанализируйте список изменений и составьте краткое смысловое описание."
        
        # Ограничиваем количество ответов для анализа (первые 15 наиболее важных)
        # Сортируем по длине ответа (более длинные ответы обычно содержат больше информации)
        sorted_responses = sorted(processed_responses, key=lambda x: len(x["response"]), reverse=True)[:15]
        
        # Упрощаем ответы: убираем пути, но сохраняем информацию о страницах
        simplified_responses = []
        for item in sorted_responses:
            resp = item["response"]
            page = item.get("page")
            
            # Убираем путь из начала ответа, оставляем только суть изменений
            if "\n\n" in resp:
                _, response_text = resp.split("\n\n", 1)
            else:
                response_text = resp
            
            # Сокращаем длину каждого ответа до 200 символов
            if len(response_text) > 200:
                response_text = response_text[:200] + "..."
            
            # Формируем строку с информацией о странице
            if page:
                simplified_responses.append(f"{response_text.strip()} (страница {page})")
            else:
                simplified_responses.append(response_text.strip())
        
        # Формируем список изменений для анализа
        changes_list = "\n".join([f"{i+1}. {resp}" for i, resp in enumerate(simplified_responses)])
        
        logger.debug(f"Отправка {len(simplified_responses)} изменений к LLM для генерации краткого описания")
        
        user_prompt = f"""Проанализируйте следующие изменения в документе и составьте краткое смысловое описание в формате нумерованного списка:

{changes_list}

Составьте краткое описание, группируя похожие изменения вместе. Указывайте конкретные места изменений (разделы, пункты, таблицы) с их номерами. ОБЯЗАТЕЛЬНО включайте номера страниц для каждого изменения (если они указаны в скобках). Формат ответа - нумерованный список."""
        
        # Retry логика
        max_retries = config.llm.max_retries
        retry_delay = config.llm.retry_delay_seconds
        timeout = config.llm.timeout_seconds
        
        for attempt in range(max_retries):
            try:
                request_params = {
                    "model": self.model,
                    "messages": [
                        {"role": "system", "content": summary_system_prompt},
                        {"role": "user", "content": user_prompt}
                    ],
                    "temperature": self.temperature,
                    "max_tokens": min(self.max_tokens * 8, 2500),  # Увеличиваем лимит для краткого описания
                    "timeout": timeout * 3  # Увеличиваем таймаут для более сложного запроса
                }
                
                # Дополнительные параметры из переменных окружения
                presence_penalty = os.getenv("OPENAI_PRESENCE_PENALTY")
                if presence_penalty:
                    try:
                        request_params["presence_penalty"] = float(presence_penalty)
                    except (ValueError, TypeError):
                        pass
                
                top_p = os.getenv("OPENAI_TOP_P")
                if top_p:
                    try:
                        request_params["top_p"] = float(top_p)
                    except (ValueError, TypeError):
                        pass
                
                response = self.client.chat.completions.create(**request_params)
                
                logger.debug(f"LLM ответ получен: choices={len(response.choices) if response.choices else 0}")
                
                if response.choices and len(response.choices) > 0:
                    choice = response.choices[0]
                    # Проверяем наличие content разными способами для совместимости
                    content = None
                    if hasattr(choice, 'message'):
                        if hasattr(choice.message, 'content'):
                            content = choice.message.content
                        elif isinstance(choice.message, dict):
                            content = choice.message.get('content')
                        # Дополнительная проверка через getattr
                        if content is None:
                            content = getattr(choice.message, 'content', None)
                    
                    finish_reason = getattr(choice, 'finish_reason', None)
                    
                    logger.debug(f"Содержимое ответа: {repr(content[:100]) if content else 'None'}")
                    logger.debug(f"Finish reason: {finish_reason}, тип choice: {type(choice)}")
                    
                    if content:
                        result = content.strip()
                        # Убираем markdown форматирование жирного текста
                        result = _remove_markdown_bold(result)
                        if result:
                            logger.info(f"Краткое описание успешно сгенерировано ({len(result)} символов)")
                            if finish_reason == 'length':
                                logger.warning("Ответ LLM был обрезан из-за лимита токенов, но часть ответа получена")
                            return result
                        else:
                            logger.warning("LLM вернул пустой ответ после strip() для краткого описания")
                            return "Общие правки."
                    else:
                        logger.warning(f"LLM вернул None для краткого описания. Finish reason: {finish_reason}")
                        # Если ответ был обрезан из-за лимита токенов, увеличиваем лимит и пробуем снова
                        if finish_reason == 'length' and attempt < max_retries - 1:
                            new_max_tokens = min(self.max_tokens * 10, 3000)
                            logger.info(f"Увеличиваем max_tokens до {new_max_tokens} и повторяем попытку {attempt + 2}")
                            request_params["max_tokens"] = new_max_tokens
                            # Также сокращаем входные данные еще больше
                            if len(simplified_responses) > 10:
                                simplified_responses = simplified_responses[:10]
                                changes_list = "\n".join([f"{i+1}. {resp}" for i, resp in enumerate(simplified_responses)])
                                request_params["messages"][1]["content"] = f"""Проанализируйте следующие изменения в документе и составьте краткое смысловое описание в формате нумерованного списка:

{changes_list}

Составьте краткое описание, группируя похожие изменения вместе. Указывайте конкретные места изменений (разделы, пункты, таблицы) с их номерами. Формат ответа - нумерованный список."""
                            delay = retry_delay * (2 ** attempt)
                            time.sleep(delay)
                            continue
                        return "Общие правки."
                else:
                    logger.warning("LLM не вернул choices для краткого описания")
                    return "Общие правки."
                    
            except Exception as e:
                error_msg = str(e)
                logger.warning(f"Попытка генерации краткого описания {attempt + 1}/{max_retries} не удалась: {error_msg}")
                
                if attempt == max_retries - 1:
                    logger.error(f"Не удалось сгенерировать краткое описание: {error_msg}")
                    return "Общие правки."
                
                delay = retry_delay * (2 ** attempt)
                time.sleep(delay)
        
        return "Общие правки."

