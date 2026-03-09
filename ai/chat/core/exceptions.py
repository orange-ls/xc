"""
Custom exceptions for the AI Chat Module
"""

from dataclasses import dataclass
from datetime import datetime
from typing import Dict, Any, Optional


class ChatError(Exception):
    """Base exception for chat module errors"""
    
    def __init__(self, message: str, code: str = "CHAT_ERROR", details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.details = details or {}
        self.timestamp = datetime.now()


class ValidationError(ChatError):
    """Exception for input validation errors"""
    
    def __init__(self, message: str, field: Optional[str] = None):
        super().__init__(message, "VALIDATION_ERROR", {"field": field})
        self.field = field


class ModelError(ChatError):
    """Exception for AI model related errors"""
    
    def __init__(self, message: str, model_name: Optional[str] = None):
        super().__init__(message, "MODEL_ERROR", {"model_name": model_name})
        self.model_name = model_name


class RetryableError(ChatError):
    """Exception for errors that can be retried"""
    
    def __init__(self, message: str, retry_after: Optional[float] = None):
        super().__init__(message, "RETRYABLE_ERROR", {"retry_after": retry_after})
        self.retry_after = retry_after


class StorageError(ChatError):
    """Exception for storage related errors"""
    
    def __init__(self, message: str, operation: Optional[str] = None):
        super().__init__(message, "STORAGE_ERROR", {"operation": operation})
        self.operation = operation


class SessionError(ChatError):
    """Exception for session related errors"""
    
    def __init__(self, message: str, session_id: Optional[str] = None):
        super().__init__(message, "SESSION_ERROR", {"session_id": session_id})
        self.session_id = session_id


class ConfigurationError(ChatError):
    """Exception for configuration related errors"""
    
    def __init__(self, message: str, config_key: Optional[str] = None):
        super().__init__(message, "CONFIGURATION_ERROR", {"config_key": config_key})
        self.config_key = config_key


class AllModelsFailedError(ModelError):
    """Exception when all AI models fail"""
    
    def __init__(self, message: str = "All AI models are unavailable"):
        super().__init__(message, "all_models")
        self.code = "ALL_MODELS_FAILED"