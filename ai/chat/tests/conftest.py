"""
Pytest configuration and fixtures for AI Chat Module tests
"""

import pytest
import asyncio
from typing import Generator, List
from ..core.models import ChatConfig, Session
from ..core.chat_module import ChatModule


@pytest.fixture(scope="session")
def event_loop() -> Generator:
    """Create an instance of the default event loop for the test session."""
    loop = asyncio.get_event_loop_policy().new_event_loop()
    yield loop
    loop.close()


@pytest.fixture
def test_config() -> ChatConfig:
    """Provide test configuration"""
    return ChatConfig(
        max_message_length=1000,
        context_window_size=5,
        response_timeout=10,
        storage_backend="memory",
        ai_models=["mock"],
        fallback_models=["mock"],
        log_level="DEBUG",
        session_timeout_hours=1,
        max_concurrent_sessions=10
    )


@pytest.fixture
def chat_module(test_config: ChatConfig) -> ChatModule:
    """Provide ChatModule instance for testing"""
    return ChatModule(test_config)


@pytest.fixture
def test_session() -> Session:
    """Provide test session"""
    return Session(id="test-session-123")


@pytest.fixture
def sample_messages() -> List[str]:
    """Provide sample messages for testing"""
    return [
        "Hello, how are you?",
        "What's the weather like today?",
        "Can you help me with a programming question?",
        "Tell me a joke",
        "Goodbye!"
    ]