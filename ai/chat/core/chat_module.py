"""
Main ChatModule class - the primary coordinator for the AI chat system
"""

import asyncio
import time
from typing import Optional, List, Dict, Any
from .models import ChatConfig, ChatResponse, Message, MessageRole, Session
from .exceptions import ChatError, SessionError, ValidationError
from ..handlers.message_handler import MessageHandler
from ..managers.conversation_manager import ConversationManager
from ..generators.response_generator import ResponseGenerator
from ..managers.session_manager import SessionManager


class ChatModule:
    """Main chat module coordinator"""
    
    def __init__(self, config: Optional[ChatConfig] = None):
        """Initialize the chat module with configuration"""
        self.config = config or ChatConfig()
        
        # Initialize core components (will be implemented in later tasks)
        self.message_handler = MessageHandler(self.config)
        self.session_manager = SessionManager(self.config)
        self.conversation_manager = ConversationManager(self.config)
        self.response_generator = ResponseGenerator(self.config)
        
        # Track active sessions
        self._active_sessions = {}
    
    async def process_message(self, session_id: str, message: str) -> ChatResponse:
        """
        Process a user message and return a response
        
        Args:
            session_id: The session identifier
            message: The user message content
            
        Returns:
            ChatResponse: The generated response
            
        Raises:
            ValidationError: If message validation fails
            SessionError: If session is invalid or expired
            ChatError: For other processing errors
        """
        start_time = time.time()
        
        try:
            # Validate the message
            validation_result = self.message_handler.validate_message(message)
            if not validation_result.is_valid:
                raise ValidationError(validation_result.error_message or "Invalid message")
            
            # Get or validate session
            session = await self.session_manager.get_session(session_id)
            if not session:
                raise SessionError(f"Session {session_id} not found", session_id)
            
            if session.is_expired(self.config.session_timeout_hours):
                raise SessionError(f"Session {session_id} has expired", session_id)
            
            # Preprocess the message
            processed_message = self.message_handler.preprocess_message(message)
            
            # Create message object
            user_message = Message(
                session_id=session_id,
                content=processed_message,
                role=MessageRole.USER
            )
            
            # Add message to conversation context
            await self.conversation_manager.add_message(session_id, user_message)
            
            # Get conversation context
            context = await self.conversation_manager.get_context(
                session_id, 
                self.config.context_window_size
            )
            
            # Generate AI response
            ai_response = await self.response_generator.generate_response(context, processed_message)
            
            # Create assistant message
            assistant_message = Message(
                session_id=session_id,
                content=ai_response.content,
                role=MessageRole.ASSISTANT
            )
            
            # Add assistant response to context
            await self.conversation_manager.add_message(session_id, assistant_message)
            
            # Update session activity
            session.update_activity()
            session.message_count += 2  # User message + assistant response
            await self.session_manager.update_session(session)
            
            # Calculate response time
            response_time = time.time() - start_time
            
            # Create chat response
            chat_response = ChatResponse(
                message=ai_response.content,
                session_id=session_id,
                response_time=response_time,
                model_used=ai_response.model_name,
                confidence=ai_response.confidence_score,
                metadata={
                    "token_usage": ai_response.token_usage.__dict__,
                    "processing_time": ai_response.processing_time
                }
            )
            
            return chat_response
            
        except Exception as e:
            response_time = time.time() - start_time
            if isinstance(e, (ValidationError, SessionError, ChatError)):
                raise
            else:
                raise ChatError(f"Unexpected error processing message: {str(e)}")
    
    def create_session(self) -> str:
        """
        Create a new conversation session
        
        Returns:
            str: The new session identifier
        """
        return self.session_manager.create_session()
    
    async def end_session(self, session_id: str) -> bool:
        """
        End a conversation session
        
        Args:
            session_id: The session identifier to end
            
        Returns:
            bool: True if session was successfully ended
        """
        return await self.session_manager.end_session(session_id)
    
    async def get_session_info(self, session_id: str) -> Optional[Session]:
        """
        Get session information
        
        Args:
            session_id: The session identifier
            
        Returns:
            Optional[Session]: Session information if found
        """
        return await self.session_manager.get_session(session_id)
    
    async def list_active_sessions(self) -> List[str]:
        """
        List all active session IDs
        
        Returns:
            list[str]: List of active session identifiers
        """
        return await self.session_manager.list_active_sessions()
    
    async def cleanup_expired_sessions(self) -> int:
        """
        Clean up expired sessions
        
        Returns:
            int: Number of sessions cleaned up
        """
        return await self.session_manager.cleanup_expired_sessions()
    
    async def health_check(self) -> Dict[str, Any]:
        """
        Perform system health check
        
        Returns:
            Dict[str, Any]: Health status information
        """
        health_status = {
            "status": "healthy",
            "components": {},
            "timestamp": time.time()
        }
        
        try:
            # Check response generator (AI models)
            ai_health = await self.response_generator.health_check()
            health_status["components"]["ai_models"] = ai_health
            
            # Check storage
            storage_health = await self.conversation_manager.health_check()
            health_status["components"]["storage"] = storage_health
            
            # Check session manager
            session_health = await self.session_manager.health_check()
            health_status["components"]["sessions"] = session_health
            
            # Overall status
            all_healthy = all(
                comp.get("status") == "healthy" 
                for comp in health_status["components"].values()
            )
            
            if not all_healthy:
                health_status["status"] = "degraded"
                
        except Exception as e:
            health_status["status"] = "unhealthy"
            health_status["error"] = str(e)
        
        return health_status