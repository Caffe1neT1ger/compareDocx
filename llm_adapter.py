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
from pathlib import Path

# Попытка загрузить переменные окружения из .env файла
try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    # Если python-dotenv не установлен, работаем только с переменными окружения системы
    pass


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
        self.model = model or os.getenv("OPENAI_MODEL", "gpt-3.5-turbo")
        
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
        
        Args:
            old_text: Текст из старого документа (базовая версия)
            new_text: Текст из нового документа (измененная версия)
            context: Дополнительный контекст (например, путь к элементу, страница)
                    Может быть использован для более точного анализа
        
        Returns:
            Ответ LLM в деловом стиле о характере изменений.
            Если LLM недоступен или произошла ошибка, возвращает пустую строку.
        """
        if not self.enabled or not self.client:
            return ""
        
        # Формирование пользовательского промпта из шаблона
        context_section = f"\n\nКонтекст: {context}" if context else ""
        user_prompt = self.user_prompt_template.format(
            old_text=old_text,
            new_text=new_text,
            context_section=context_section
        )
        
        try:
            # Отправка запроса к OpenAI API
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": self.system_prompt},
                    {"role": "user", "content": user_prompt}
                ],
                temperature=self.temperature,
                max_tokens=self.max_tokens
            )
            
            # Извлечение ответа
            if response.choices and len(response.choices) > 0:
                return response.choices[0].message.content.strip()
            else:
                return ""
                
        except Exception as e:
            # Обработка ошибок без прерывания работы программы
            print(f"Ошибка при обращении к LLM: {e}")
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

