"""
Session manager for handling conversation sessions
"""

import uuid
from datetime import datetime
from typing import Optional, List, Dict, Any
from ..core.models import Session, SessionStatus, ChatConfig
from ..core.exceptions import SessionError


class SessionManager:
    """Manages conversation sessions"""
    
    def __init__(self, config: ChatConfig):
        self.config = config
        self.sessions: Dict[str, Session] = {}
        self.max_sessions = config.max_concurrent_sessions
    
    def create_session(self) -> str:
        """
        Create a new conversation session
        
        Returns:
            str: The new session identifier
            
        Raises:
            SessionError: If maximum sessions exceeded
        """
        if len(self.sessions) >= self.max_sessions:
            # Clean up expired sessions first
            self._cleanup_expired_sessions()
            
            # Check again after cleanup
            if len(self.sessions) >= self.max_sessions:
                raise SessionError("Maximum number of concurrent sessions exceeded")
        
        session_id = str(uuid.uuid4())
        session = Session(id=session_id)
        self.sessions[session_id] = session
        
        return session_id
    
    async def get_session(self, session_id: str) -> Optional[Session]:
        """
        Get session by ID
        
        Args:
            session_id: The session identifier
            
        Returns:
            Optional[Session]: The session if found
        """
        return self.sessions.get(session_id)
    
    async def update_session(self, session: Session) -> None:
        """
        Update session information
        
        Args:
            session: The session to update
        """
        if session.id in self.sessions:
            self.sessions[session.id] = session
    
    async def end_session(self, session_id: str) -> bool:
        """
        End a conversation session
        
        Args:
            session_id: The session identifier
            
        Returns:
            bool: True if session was successfully ended
        """
        if session_id in self.sessions:
            session = self.sessions[session_id]
            session.status = SessionStatus.ENDED
            # Keep session for a while for potential recovery
            return True
        return False
    
    async def delete_session(self, session_id: str) -> bool:
        """
        Permanently delete a session
        
        Args:
            session_id: The session identifier
            
        Returns:
            bool: True if session was deleted
        """
        if session_id in self.sessions:
            del self.sessions[session_id]
            return True
        return False
    
    async def list_active_sessions(self) -> List[str]:
        """
        List all active session IDs
        
        Returns:
            List[str]: List of active session identifiers
        """
        return [
            session_id for session_id, session in self.sessions.items()
            if session.status == SessionStatus.ACTIVE and not session.is_expired(self.config.session_timeout_hours)
        ]
    
    async def cleanup_expired_sessions(self) -> int:
        """
        Clean up expired sessions
        
        Returns:
            int: Number of sessions cleaned up
        """
        return self._cleanup_expired_sessions()
    
    def _cleanup_expired_sessions(self) -> int:
        """
        Internal method to clean up expired sessions
        
        Returns:
            int: Number of sessions cleaned up
        """
        expired_sessions = []
        
        for session_id, session in self.sessions.items():
            if session.is_expired(self.config.session_timeout_hours):
                expired_sessions.append(session_id)
        
        for session_id in expired_sessions:
            del self.sessions[session_id]
        
        return len(expired_sessions)
    
    async def health_check(self) -> Dict[str, Any]:
        """
        Perform health check on session manager
        
        Returns:
            Dict[str, Any]: Health status information
        """
        active_sessions = await self.list_active_sessions()
        
        return {
            "status": "healthy",
            "active_sessions": len(active_sessions),
            "total_sessions": len(self.sessions),
            "max_sessions": self.max_sessions,
            "utilization": len(self.sessions) / self.max_sessions
        }