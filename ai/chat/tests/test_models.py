"""
Unit tests for core data models
"""

import pytest
from datetime import datetime, timedelta
from ..core.models import Message, MessageRole, Session, SessionStatus, ChatConfig, ValidationResult


class TestMessage:
    """Test Message model"""
    
    def test_message_creation(self):
        """Test basic message creation"""
        message = Message(
            session_id="test-session",
            content="Hello world",
            role=MessageRole.USER
        )
        
        assert message.session_id == "test-session"
        assert message.content == "Hello world"
        assert message.role == MessageRole.USER
        assert message.id is not None
        assert isinstance(message.timestamp, datetime)
    
    def test_message_validation_empty_content(self):
        """Test message validation with empty content"""
        with pytest.raises(ValueError, match="Message content cannot be empty"):
            Message(content="")
    
    def test_message_validation_whitespace_only(self):
        """Test message validation with whitespace-only content"""
        with pytest.raises(ValueError, match="Message content cannot be empty"):
            Message(content="   ")
    
    def test_message_validation_too_long(self):
        """Test message validation with content too long"""
        long_content = "x" * 4001
        with pytest.raises(ValueError, match="Message content exceeds maximum length"):
            Message(content=long_content)
    
    def test_message_max_length_boundary(self):
        """Test message at maximum length boundary"""
        max_content = "x" * 4000
        message = Message(content=max_content)
        assert len(message.content) == 4000


class TestSession:
    """Test Session model"""
    
    def test_session_creation(self):
        """Test basic session creation"""
        session = Session()
        
        assert session.id is not None
        assert isinstance(session.created_at, datetime)
        assert isinstance(session.last_activity, datetime)
        assert session.status == SessionStatus.ACTIVE
        assert session.message_count == 0
    
    def test_session_expiry_not_expired(self):
        """Test session expiry check for active session"""
        session = Session()
        assert not session.is_expired(24)
    
    def test_session_expiry_expired(self):
        """Test session expiry check for old session"""
        session = Session()
        # Manually set old timestamp
        session.last_activity = datetime.now() - timedelta(hours=25)
        assert session.is_expired(24)
    
    def test_session_update_activity(self):
        """Test session activity update"""
        session = Session()
        old_activity = session.last_activity
        
        # Wait a small amount to ensure timestamp difference
        import time
        time.sleep(0.01)
        
        session.update_activity()
        assert session.last_activity > old_activity


class TestChatConfig:
    """Test ChatConfig model"""
    
    def test_default_config(self):
        """Test default configuration values"""
        config = ChatConfig()
        
        assert config.max_message_length == 4000
        assert config.context_window_size == 10
        assert config.response_timeout == 5
        assert config.storage_backend == "file"
        assert config.ai_models == ["mock"]
        assert config.fallback_models == ["mock"]
        assert config.log_level == "INFO"
        assert config.session_timeout_hours == 24
        assert config.max_concurrent_sessions == 100
    
    def test_custom_config(self):
        """Test custom configuration values"""
        config = ChatConfig(
            max_message_length=2000,
            context_window_size=5,
            ai_models=["gpt-4", "claude"],
            log_level="DEBUG"
        )
        
        assert config.max_message_length == 2000
        assert config.context_window_size == 5
        assert config.ai_models == ["gpt-4", "claude"]
        assert config.log_level == "DEBUG"


class TestValidationResult:
    """Test ValidationResult model"""
    
    def test_valid_result(self):
        """Test valid validation result"""
        result = ValidationResult(is_valid=True)
        
        assert result.is_valid is True
        assert result.error_message is None
        assert result.warnings == []
    
    def test_invalid_result(self):
        """Test invalid validation result"""
        result = ValidationResult(
            is_valid=False,
            error_message="Test error",
            warnings=["Test warning"]
        )
        
        assert result.is_valid is False
        assert result.error_message == "Test error"
        assert result.warnings == ["Test warning"]