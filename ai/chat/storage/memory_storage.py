"""
In-memory storage backend for development and testing
"""

from typing import List, Optional, Dict
from ..core.models import Message, Session
from ..core.interfaces import StorageBackend


class MemoryStorage(StorageBackend):
    """In-memory storage implementation"""
    
    def __init__(self):
        self.messages: Dict[str, List[Message]] = {}
        self.sessions: Dict[str, Session] = {}
    
    async def save_message(self, session_id: str, message: Message) -> None:
        """Save a message to memory storage"""
        if session_id not in self.messages:
            self.messages[session_id] = []
        
        self.messages[session_id].append(message)
    
    async def get_messages(self, session_id: str, limit: Optional[int] = None) -> List[Message]:
        """Retrieve messages from memory storage"""
        if session_id not in self.messages:
            return []
        
        messages = self.messages[session_id]
        
        if limit:
            # Return the most recent messages
            return messages[-limit:]
        
        return messages.copy()
    
    async def save_session(self, session: Session) -> None:
        """Save session metadata to memory storage"""
        self.sessions[session.id] = session
    
    async def get_session(self, session_id: str) -> Optional[Session]:
        """Retrieve session metadata from memory storage"""
        return self.sessions.get(session_id)
    
    async def delete_session(self, session_id: str) -> bool:
        """Delete a session and its messages from memory storage"""
        deleted = False
        
        if session_id in self.sessions:
            del self.sessions[session_id]
            deleted = True
        
        if session_id in self.messages:
            del self.messages[session_id]
            deleted = True
        
        return deleted
    
    async def backup_data(self) -> bool:
        """Backup data (no-op for memory storage)"""
        # Memory storage doesn't support backup
        # This would be implemented for persistent storage backends
        return True
    
    def get_stats(self) -> Dict[str, int]:
        """Get storage statistics"""
        total_messages = sum(len(msgs) for msgs in self.messages.values())
        
        return {
            "total_sessions": len(self.sessions),
            "total_messages": total_messages,
            "active_sessions": len([s for s in self.sessions.values() if s.status.value == "active"])
        }