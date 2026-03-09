"""
Core interfaces and abstract base classes for the AI Chat Module
"""

from abc import ABC, abstractmethod
from typing import List, Optional, Dict, Any
from .models import Message, AIResponse, Session, ChatConfig


class StorageBackend(ABC):
    """Abstract base class for storage backends"""
    
    @abstractmethod
    async def save_message(self, session_id: str, message: Message) -> None:
        """Save a message to storage"""
        pass
    
    @abstractmethod
    async def get_messages(self, session_id: str, limit: Optional[int] = None) -> List[Message]:
        """Retrieve messages from storage"""
        pass
    
    @abstractmethod
    async def save_session(self, session: Session) -> None:
        """Save session metadata"""
        pass
    
    @abstractmethod
    async def get_session(self, session_id: str) -> Optional[Session]:
        """Retrieve session metadata"""
        pass
    
    @abstractmethod
    async def delete_session(self, session_id: str) -> bool:
        """Delete a session and its messages"""
        pass
    
    @abstractmethod
    async def backup_data(self) -> bool:
        """Backup all data"""
        pass


class AIModelProvider(ABC):
    """Abstract base class for AI model providers"""
    
    @abstractmethod
    async def generate_response(self, context: List[Message], message: str) -> AIResponse:
        """Generate AI response given context and message"""
        pass
    
    @abstractmethod
    def get_model_name(self) -> str:
        """Get the name of this AI model"""
        pass
    
    @abstractmethod
    def is_available(self) -> bool:
        """Check if the model is currently available"""
        pass
    
    @abstractmethod
    async def health_check(self) -> bool:
        """Perform health check on the model"""
        pass


class MessageValidator(ABC):
    """Abstract base class for message validators"""
    
    @abstractmethod
    def validate(self, message: str) -> bool:
        """Validate a message"""
        pass
    
    @abstractmethod
    def get_validation_errors(self, message: str) -> List[str]:
        """Get validation errors for a message"""
        pass


class ContextManager(ABC):
    """Abstract base class for context management"""
    
    @abstractmethod
    async def get_context(self, session_id: str, limit: int = 10) -> List[Message]:
        """Get conversation context"""
        pass
    
    @abstractmethod
    async def add_message(self, session_id: str, message: Message) -> None:
        """Add message to context"""
        pass
    
    @abstractmethod
    def manage_context_window(self, session_id: str) -> None:
        """Manage context window size"""
        pass


class ConfigurationProvider(ABC):
    """Abstract base class for configuration providers"""
    
    @abstractmethod
    def load_config(self) -> ChatConfig:
        """Load configuration"""
        pass
    
    @abstractmethod
    def reload_config(self) -> ChatConfig:
        """Reload configuration"""
        pass
    
    @abstractmethod
    def validate_config(self, config: ChatConfig) -> bool:
        """Validate configuration"""
        pass