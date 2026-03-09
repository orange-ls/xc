"""
Core data models for the AI Chat Module
"""

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Dict, Any, Optional, List
import uuid


class MessageRole(Enum):
    """Message role enumeration"""
    USER = "user"
    ASSISTANT = "assistant"
    SYSTEM = "system"


class SessionStatus(Enum):
    """Session status enumeration"""
    ACTIVE = "active"
    INACTIVE = "inactive"
    ENDED = "ended"


@dataclass
class Message:
    """Message data structure"""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    session_id: str = ""
    content: str = ""
    role: MessageRole = MessageRole.USER
    timestamp: datetime = field(default_factory=datetime.now)
    metadata: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        """Validate message after initialization"""
        if not self.content.strip():
            raise ValueError("Message content cannot be empty")
        if len(self.content) > 4000:
            raise ValueError("Message content exceeds maximum length of 4000 characters")


@dataclass
class ChatResponse:
    """Chat response data structure"""
    message: str
    session_id: str
    response_time: float
    model_used: str = "default"
    confidence: float = 1.0
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class Session:
    """Conversation session model"""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    created_at: datetime = field(default_factory=datetime.now)
    last_activity: datetime = field(default_factory=datetime.now)
    status: SessionStatus = SessionStatus.ACTIVE
    metadata: Dict[str, Any] = field(default_factory=dict)
    message_count: int = 0

    def is_expired(self, timeout_hours: int = 24) -> bool:
        """Check if session is expired"""
        return (datetime.now() - self.last_activity).total_seconds() > (timeout_hours * 3600)

    def update_activity(self):
        """Update last activity timestamp"""
        self.last_activity = datetime.now()


@dataclass
class TokenUsage:
    """Token usage statistics"""
    prompt_tokens: int = 0
    completion_tokens: int = 0
    total_tokens: int = 0


@dataclass
class AIResponse:
    """AI model response"""
    content: str
    model_name: str
    confidence_score: float = 1.0
    processing_time: float = 0.0
    token_usage: TokenUsage = field(default_factory=TokenUsage)
    error: Optional[str] = None


@dataclass
class ChatConfig:
    """Chat module configuration"""
    max_message_length: int = 4000
    context_window_size: int = 10
    response_timeout: int = 5
    storage_backend: str = "file"
    data_dir: str = "ai/chat/data"
    ai_models: List[str] = field(default_factory=lambda: ["mock"])
    fallback_models: List[str] = field(default_factory=lambda: ["mock"])
    log_level: str = "INFO"
    session_timeout_hours: int = 24
    max_concurrent_sessions: int = 100
    
    # Backup configuration
    max_backups: int = 10
    backup_interval_hours: int = 24
    auto_backup_enabled: bool = True
    verify_backups: bool = True


@dataclass
class ValidationResult:
    """Message validation result"""
    is_valid: bool
    error_message: Optional[str] = None
    warnings: List[str] = field(default_factory=list)


@dataclass
class Intent:
    """Message intent extraction result"""
    name: str
    confidence: float
    parameters: Dict[str, Any] = field(default_factory=dict)