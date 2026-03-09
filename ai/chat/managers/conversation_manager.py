"""
Conversation manager for handling dialogue context and history
"""

from typing import List, Dict, Any
from ..core.models import Message, ChatConfig
from ..core.interfaces import ContextManager
from ..storage.context_store import ContextStore


class ConversationManager(ContextManager):
    """Manages conversation state and context"""
    
    def __init__(self, config: ChatConfig):
        self.config = config
        self.context_store = ContextStore(config)
        self.context_window_size = config.context_window_size
    
    async def get_context(self, session_id: str, limit: int = 10) -> List[Message]:
        """
        Get conversation context for a session
        
        Args:
            session_id: The session identifier
            limit: Maximum number of messages to retrieve
            
        Returns:
            List[Message]: List of messages in chronological order
        """
        # Use the configured context window size if limit is not specified
        effective_limit = min(limit, self.context_window_size)
        
        messages = await self.context_store.get_messages(session_id, effective_limit)
        
        # Ensure messages are in chronological order
        messages.sort(key=lambda m: m.timestamp)
        
        return messages
    
    async def add_message(self, session_id: str, message: Message) -> None:
        """
        Add a message to the conversation context
        
        Args:
            session_id: The session identifier
            message: The message to add
        """
        # Ensure message has correct session_id
        message.session_id = session_id
        
        # Save message to storage
        await self.context_store.save_message(session_id, message)
        
        # Manage context window if needed
        self.manage_context_window(session_id)
    
    def manage_context_window(self, session_id: str) -> None:
        """
        管理会话的上下文窗口大小
        
        Args:
            session_id: 会话标识符
        """
        # 这是一个异步操作的同步包装器
        # 在实际实现中，这会：
        # 1. 检查会话是否有太多消息
        # 2. 归档或删除超出窗口的旧消息
        # 3. 通过只保留最近的上下文来维持性能
        
        # 注意：由于这是从同步方法调用的，我们不能直接使用async/await
        # 在真实的实现中，这个方法应该是异步的，或者使用后台任务
        pass
    
    async def get_conversation_summary(self, session_id: str) -> Dict[str, Any]:
        """
        Get a summary of the conversation
        
        Args:
            session_id: The session identifier
            
        Returns:
            Dict[str, Any]: Conversation summary
        """
        messages = await self.context_store.get_messages(session_id)
        
        user_messages = [m for m in messages if m.role.value == "user"]
        assistant_messages = [m for m in messages if m.role.value == "assistant"]
        
        return {
            "session_id": session_id,
            "total_messages": len(messages),
            "user_messages": len(user_messages),
            "assistant_messages": len(assistant_messages),
            "first_message_time": messages[0].timestamp if messages else None,
            "last_message_time": messages[-1].timestamp if messages else None
        }
    
    async def clear_context(self, session_id: str) -> bool:
        """
        Clear conversation context for a session
        
        Args:
            session_id: The session identifier
            
        Returns:
            bool: True if context was cleared successfully
        """
        return await self.context_store.delete_session(session_id)
    
    async def health_check(self) -> Dict[str, Any]:
        """
        Perform health check on conversation manager
        
        Returns:
            Dict[str, Any]: Health status information
        """
        try:
            # Check if storage is accessible
            storage_health = await self.context_store.health_check()
            
            return {
                "status": "healthy" if storage_health["status"] == "healthy" else "degraded",
                "storage": storage_health,
                "context_window_size": self.context_window_size
            }
        except Exception as e:
            return {
                "status": "unhealthy",
                "error": str(e)
            }