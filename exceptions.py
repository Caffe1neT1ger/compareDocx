"""
Кастомные исключения для проекта.

Предоставляет специфичные исключения для лучшей обработки ошибок
и более информативных сообщений об ошибках.
"""


class CompareDocxError(Exception):
    """Базовое исключение для всех ошибок проекта."""
    pass


class DocumentLoadError(CompareDocxError):
    """Ошибка при загрузке документа."""
    
    def __init__(self, file_path: str, reason: str = ""):
        message = f"Не удалось загрузить документ: {file_path}"
        if reason:
            message += f". Причина: {reason}"
        super().__init__(message)
        self.file_path = file_path
        self.reason = reason


class DocumentParseError(CompareDocxError):
    """Ошибка при парсинге документа."""
    
    def __init__(self, file_path: str, reason: str = ""):
        message = f"Ошибка парсинга документа: {file_path}"
        if reason:
            message += f". Причина: {reason}"
        super().__init__(message)
        self.file_path = file_path
        self.reason = reason


class FileSizeError(CompareDocxError):
    """Ошибка: файл слишком большой."""
    
    def __init__(self, file_path: str, size_mb: float, max_size_mb: int):
        message = (
            f"Файл слишком большой: {file_path} "
            f"({size_mb:.2f} МБ). "
            f"Максимальный размер: {max_size_mb} МБ"
        )
        super().__init__(message)
        self.file_path = file_path
        self.size_mb = size_mb
        self.max_size_mb = max_size_mb


class ValidationError(CompareDocxError):
    """Ошибка валидации входных данных."""
    
    def __init__(self, message: str):
        super().__init__(f"Ошибка валидации: {message}")


class ComparisonError(CompareDocxError):
    """Ошибка при сравнении документов."""
    
    def __init__(self, reason: str = ""):
        message = "Ошибка при сравнении документов"
        if reason:
            message += f". Причина: {reason}"
        super().__init__(message)
        self.reason = reason


class ExportError(CompareDocxError):
    """Ошибка при экспорте результатов."""
    
    def __init__(self, output_path: str, reason: str = ""):
        message = f"Ошибка при экспорте в файл: {output_path}"
        if reason:
            message += f". Причина: {reason}"
        super().__init__(message)
        self.output_path = output_path
        self.reason = reason


class LLMError(CompareDocxError):
    """Ошибка при работе с LLM."""
    
    def __init__(self, reason: str = ""):
        message = "Ошибка при обращении к LLM"
        if reason:
            message += f". Причина: {reason}"
        super().__init__(message)
        self.reason = reason

