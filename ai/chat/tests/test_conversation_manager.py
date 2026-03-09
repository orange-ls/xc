"""
对话管理器的单元测试
"""

import pytest
from datetime import datetime, timedelta
from ..managers.conversation_manager import ConversationManager
from ..core.models import ChatConfig, Message, MessageRole
from ..core.exceptions import SessionError


class TestConversationManager:
    """测试ConversationManager功能"""
    
    @pytest.fixture
    def config(self):
        """创建测试配置"""
        return ChatConfig(
            context_window_size=5,  # 较小的窗口用于测试
            storage_backend="memory"
        )
    
    @pytest.fixture
    def conversation_manager(self, config):
        """创建ConversationManager实例用于测试"""
        return ConversationManager(config)
    
    @pytest.mark.asyncio
    async def test_add_and_get_message(self, conversation_manager):
        """测试添加和获取消息"""
        session_id = "test-session-1"
        message = Message(
            content="Hello, world!",
            role=MessageRole.USER
        )
        
        # 添加消息
        await conversation_manager.add_message(session_id, message)
        
        # 获取上下文
        context = await conversation_manager.get_context(session_id)
        
        assert len(context) == 1
        assert context[0].content == "Hello, world!"
        assert context[0].role == MessageRole.USER
        assert context[0].session_id == session_id
    
    @pytest.mark.asyncio
    async def test_multiple_messages_chronological_order(self, conversation_manager):
        """测试多条消息按时间顺序排列"""
        session_id = "test-session-2"
        
        # 添加多条消息
        messages = [
            Message(content="First message", role=MessageRole.USER),
            Message(content="Second message", role=MessageRole.ASSISTANT),
            Message(content="Third message", role=MessageRole.USER)
        ]
        
        for i, msg in enumerate(messages):
            # 确保时间戳不同
            msg.timestamp = datetime.now() + timedelta(seconds=i)
            await conversation_manager.add_message(session_id, msg)
        
        # 获取上下文
        context = await conversation_manager.get_context(session_id)
        
        assert len(context) == 3
        assert context[0].content == "First message"
        assert context[1].content == "Second message"
        assert context[2].content == "Third message"
    
    @pytest.mark.asyncio
    async def test_context_window_limit(self, conversation_manager):
        """测试上下文窗口限制"""
        session_id = "test-session-3"
        
        # 添加超过窗口大小的消息
        for i in range(8):  # 超过配置的5条限制
            message = Message(
                content=f"Message {i+1}",
                role=MessageRole.USER,
                timestamp=datetime.now() + timedelta(seconds=i)
            )
            await conversation_manager.add_message(session_id, message)
        
        # 获取上下文（应该限制在窗口大小内）
        context = await conversation_manager.get_context(session_id)
        
        # 应该只返回最近的5条消息
        assert len(context) <= 5
    
    @pytest.mark.asyncio
    async def test_get_context_with_custom_limit(self, conversation_manager):
        """测试使用自定义限制获取上下文"""
        session_id = "test-session-4"
        
        # 添加多条消息
        for i in range(6):
            message = Message(
                content=f"Message {i+1}",
                role=MessageRole.USER,
                timestamp=datetime.now() + timedelta(seconds=i)
            )
            await conversation_manager.add_message(session_id, message)
        
        # 使用自定义限制获取上下文
        context = await conversation_manager.get_context(session_id, limit=3)
        
        assert len(context) == 3
    
    @pytest.mark.asyncio
    async def test_conversation_summary(self, conversation_manager):
        """测试对话摘要"""
        session_id = "test-session-5"
        
        # 添加不同角色的消息
        messages = [
            Message(content="User message 1", role=MessageRole.USER),
            Message(content="Assistant response 1", role=MessageRole.ASSISTANT),
            Message(content="User message 2", role=MessageRole.USER),
            Message(content="Assistant response 2", role=MessageRole.ASSISTANT),
            Message(content="System message", role=MessageRole.SYSTEM)
        ]
        
        for msg in messages:
            await conversation_manager.add_message(session_id, msg)
        
        # 获取对话摘要
        summary = await conversation_manager.get_conversation_summary(session_id)
        
        assert summary["session_id"] == session_id
        assert summary["total_messages"] == 5
        assert summary["user_messages"] == 2
        assert summary["assistant_messages"] == 2
        assert summary["first_message_time"] is not None
        assert summary["last_message_time"] is not None
    
    @pytest.mark.asyncio
    async def test_clear_context(self, conversation_manager):
        """测试清除上下文"""
        session_id = "test-session-6"
        
        # 添加一些消息
        message = Message(content="Test message", role=MessageRole.USER)
        await conversation_manager.add_message(session_id, message)
        
        # 验证消息存在
        context = await conversation_manager.get_context(session_id)
        assert len(context) == 1
        
        # 清除上下文
        result = await conversation_manager.clear_context(session_id)
        assert result is True
        
        # 验证上下文已清除
        context = await conversation_manager.get_context(session_id)
        assert len(context) == 0
    
    @pytest.mark.asyncio
    async def test_empty_session_context(self, conversation_manager):
        """测试空会话的上下文"""
        session_id = "empty-session"
        
        context = await conversation_manager.get_context(session_id)
        
        assert len(context) == 0
    
    @pytest.mark.asyncio
    async def test_empty_session_summary(self, conversation_manager):
        """测试空会话的摘要"""
        session_id = "empty-session"
        
        summary = await conversation_manager.get_conversation_summary(session_id)
        
        assert summary["session_id"] == session_id
        assert summary["total_messages"] == 0
        assert summary["user_messages"] == 0
        assert summary["assistant_messages"] == 0
        assert summary["first_message_time"] is None
        assert summary["last_message_time"] is None
    
    @pytest.mark.asyncio
    async def test_health_check(self, conversation_manager):
        """测试健康检查"""
        health_status = await conversation_manager.health_check()
        
        assert "status" in health_status
        assert health_status["status"] in ["healthy", "degraded", "unhealthy"]
        assert "storage" in health_status
        assert "context_window_size" in health_status
        assert health_status["context_window_size"] == 5
    
    @pytest.mark.asyncio
    async def test_message_session_id_assignment(self, conversation_manager):
        """测试消息会话ID的自动分配"""
        session_id = "test-session-7"
        message = Message(
            content="Test message",
            role=MessageRole.USER
        )
        
        # 消息最初没有session_id
        assert message.session_id == ""
        
        # 添加消息后应该自动分配session_id
        await conversation_manager.add_message(session_id, message)
        
        # 获取上下文验证session_id
        context = await conversation_manager.get_context(session_id)
        assert len(context) == 1
        assert context[0].session_id == session_id
    
    @pytest.mark.asyncio
    async def test_concurrent_sessions(self, conversation_manager):
        """测试并发会话的隔离"""
        session_id1 = "session-1"
        session_id2 = "session-2"
        
        # 向不同会话添加消息
        message1 = Message(content="Message for session 1", role=MessageRole.USER)
        message2 = Message(content="Message for session 2", role=MessageRole.USER)
        
        await conversation_manager.add_message(session_id1, message1)
        await conversation_manager.add_message(session_id2, message2)
        
        # 验证会话隔离
        context1 = await conversation_manager.get_context(session_id1)
        context2 = await conversation_manager.get_context(session_id2)
        
        assert len(context1) == 1
        assert len(context2) == 1
        assert context1[0].content == "Message for session 1"
        assert context2[0].content == "Message for session 2"
        assert context1[0].session_id == session_id1
        assert context2[0].session_id == session_id2