"""
Unit tests for ChatModule core functionality
"""

import pytest
from ..core.chat_module import ChatModule
from ..core.models import ChatConfig
from ..core.exceptions import ValidationError, SessionError


class TestChatModule:
    """Test ChatModule functionality"""
    
    @pytest.mark.asyncio
    async def test_chat_module_initialization(self, test_config):
        """Test ChatModule initialization"""
        chat_module = ChatModule(test_config)
        
        assert chat_module.config == test_config
        assert chat_module.message_handler is not None
        assert chat_module.session_manager is not None
        assert chat_module.conversation_manager is not None
        assert chat_module.response_generator is not None
    
    def test_create_session(self, chat_module):
        """Test session creation"""
        session_id = chat_module.create_session()
        
        assert session_id is not None
        assert isinstance(session_id, str)
        assert len(session_id) > 0
    
    @pytest.mark.asyncio
    async def test_process_message_invalid_session(self, chat_module):
        """Test processing message with invalid session"""
        with pytest.raises(SessionError):
            await chat_module.process_message("invalid-session", "Hello")
    
    @pytest.mark.asyncio
    async def test_process_message_empty_message(self, chat_module):
        """Test processing empty message"""
        session_id = chat_module.create_session()
        
        with pytest.raises(ValidationError):
            await chat_module.process_message(session_id, "")
    
    @pytest.mark.asyncio
    async def test_process_message_too_long(self, chat_module):
        """Test processing message that's too long"""
        session_id = chat_module.create_session()
        long_message = "x" * (chat_module.config.max_message_length + 1)
        
        with pytest.raises(ValidationError):
            await chat_module.process_message(session_id, long_message)
    
    @pytest.mark.asyncio
    async def test_process_message_success(self, chat_module):
        """Test successful message processing"""
        session_id = chat_module.create_session()
        message = "Hello, how are you?"
        
        response = await chat_module.process_message(session_id, message)
        
        assert response is not None
        assert response.message is not None
        assert response.session_id == session_id
        assert response.response_time > 0
        assert response.model_used is not None
    
    @pytest.mark.asyncio
    async def test_end_session(self, chat_module):
        """Test ending a session"""
        session_id = chat_module.create_session()
        
        result = await chat_module.end_session(session_id)
        assert result is True
    
    @pytest.mark.asyncio
    async def test_end_nonexistent_session(self, chat_module):
        """Test ending a nonexistent session"""
        result = await chat_module.end_session("nonexistent-session")
        assert result is False
    
    @pytest.mark.asyncio
    async def test_get_session_info(self, chat_module):
        """Test getting session information"""
        session_id = chat_module.create_session()
        
        session_info = await chat_module.get_session_info(session_id)
        
        assert session_info is not None
        assert session_info.id == session_id
    
    @pytest.mark.asyncio
    async def test_list_active_sessions(self, chat_module):
        """Test listing active sessions"""
        session_id1 = chat_module.create_session()
        session_id2 = chat_module.create_session()
        
        active_sessions = await chat_module.list_active_sessions()
        
        assert session_id1 in active_sessions
        assert session_id2 in active_sessions
    
    @pytest.mark.asyncio
    async def test_health_check(self, chat_module):
        """Test system health check"""
        health_status = await chat_module.health_check()
        
        assert "status" in health_status
        assert "components" in health_status
        assert "timestamp" in health_status
        assert health_status["status"] in ["healthy", "degraded", "unhealthy"]