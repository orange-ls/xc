"""
AI Chat Module

A modular conversational AI system for handling user interactions,
maintaining conversation context, and providing intelligent responses.
"""

from .core.chat_module import ChatModule
from .core.models import Message, ChatResponse, ChatConfig
from .core.exceptions import ChatError, ValidationError, ModelError

__version__ = "1.0.0"
__all__ = [
    "ChatModule",
    "Message", 
    "ChatResponse",
    "ChatConfig",
    "ChatError",
    "ValidationError", 
    "ModelError"
]