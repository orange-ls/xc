"""
File-based storage backend for persistent conversation data
"""

import json
import os
import tempfile
import shutil
from datetime import datetime
from typing import List, Optional, Dict, Any
from pathlib import Path

from ..core.models import Message, Session, MessageRole, SessionStatus
from ..core.interfaces import StorageBackend
from ..core.exceptions import StorageError


class FileStorage(StorageBackend):
    """File-based storage implementation with atomic operations"""
    
    def __init__(self, data_dir: str = "ai/chat/data"):
        self.data_dir = Path(data_dir)
        self.sessions_dir = self.data_dir / "sessions"
        self.conversations_dir = self.data_dir / "conversations"
        self.backups_dir = self.data_dir / "backups"
        
        # Create directories if they don't exist
        self._ensure_directories()
    
    def _ensure_directories(self):
        """Ensure all required directories exist"""
        for directory in [self.data_dir, self.sessions_dir, self.conversations_dir, self.backups_dir]:
            directory.mkdir(parents=True, exist_ok=True)
    
    def _get_session_file(self, session_id: str) -> Path:
        """Get the file path for a session"""
        return self.sessions_dir / f"{session_id}.json"
    
    def _get_messages_file(self, session_id: str) -> Path:
        """Get the file path for session messages"""
        session_dir = self.conversations_dir / session_id
        session_dir.mkdir(exist_ok=True)
        return session_dir / "messages.jsonl"
    
    def _get_metadata_file(self, session_id: str) -> Path:
        """Get the file path for session metadata"""
        session_dir = self.conversations_dir / session_id
        session_dir.mkdir(exist_ok=True)
        return session_dir / "metadata.json"
    
    async def _atomic_write(self, file_path: Path, content: str) -> None:
        """
        Perform atomic write operation to prevent data corruption
        
        Args:
            file_path: Target file path
            content: Content to write
        """
        try:
            # Write to temporary file first
            with tempfile.NamedTemporaryFile(
                mode='w', 
                dir=file_path.parent, 
                delete=False,
                suffix='.tmp'
            ) as temp_file:
                temp_file.write(content)
                temp_file.flush()
                os.fsync(temp_file.fileno())  # Force write to disk
                temp_path = temp_file.name
            
            # Atomic move to final location
            shutil.move(temp_path, file_path)
            
        except Exception as e:
            # Clean up temporary file if it exists
            if 'temp_path' in locals() and os.path.exists(temp_path):
                os.unlink(temp_path)
            raise StorageError(f"Failed to write file {file_path}: {str(e)}")
    
    async def save_message(self, session_id: str, message: Message) -> None:
        """
        Save a message to file storage using JSONL format
        
        Args:
            session_id: The session identifier
            message: The message to save
        """
        try:
            messages_file = self._get_messages_file(session_id)
            
            # Convert message to JSON
            message_json = {
                "id": message.id,
                "session_id": message.session_id,
                "content": message.content,
                "role": message.role.value,
                "timestamp": message.timestamp.isoformat(),
                "metadata": message.metadata
            }
            
            # Append to JSONL file (one JSON object per line)
            json_line = json.dumps(message_json, ensure_ascii=False) + '\n'
            
            # For append operations, we'll use a simpler approach
            # In production, you might want to implement atomic append
            with open(messages_file, 'a', encoding='utf-8') as f:
                f.write(json_line)
                f.flush()
                os.fsync(f.fileno())
            
        except Exception as e:
            raise StorageError(f"Failed to save message for session {session_id}: {str(e)}")
    
    async def get_messages(self, session_id: str, limit: Optional[int] = None) -> List[Message]:
        """
        Retrieve messages from file storage
        
        Args:
            session_id: The session identifier
            limit: Maximum number of messages to retrieve
            
        Returns:
            List[Message]: List of messages in chronological order
        """
        try:
            messages_file = self._get_messages_file(session_id)
            
            if not messages_file.exists():
                return []
            
            messages = []
            
            with open(messages_file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    
                    try:
                        message_data = json.loads(line)
                        message = Message(
                            id=message_data["id"],
                            session_id=message_data["session_id"],
                            content=message_data["content"],
                            role=MessageRole(message_data["role"]),  # Convert string back to enum
                            timestamp=datetime.fromisoformat(message_data["timestamp"]),
                            metadata=message_data.get("metadata", {})
                        )
                        messages.append(message)
                    except (json.JSONDecodeError, KeyError, ValueError) as e:
                        # Log corrupted line but continue processing
                        print(f"Warning: Corrupted message line in {messages_file}: {e}")
                        continue
            
            # Sort by timestamp to ensure chronological order
            messages.sort(key=lambda m: m.timestamp)
            
            # Apply limit if specified (get most recent messages)
            if limit and len(messages) > limit:
                messages = messages[-limit:]
            
            return messages
            
        except Exception as e:
            raise StorageError(f"Failed to retrieve messages for session {session_id}: {str(e)}")
    
    async def save_session(self, session: Session) -> None:
        """
        Save session metadata to file storage
        
        Args:
            session: The session to save
        """
        try:
            session_file = self._get_session_file(session.id)
            
            session_data = {
                "id": session.id,
                "created_at": session.created_at.isoformat(),
                "last_activity": session.last_activity.isoformat(),
                "status": session.status.value,
                "metadata": session.metadata,
                "message_count": session.message_count
            }
            
            content = json.dumps(session_data, ensure_ascii=False, indent=2)
            await self._atomic_write(session_file, content)
            
        except Exception as e:
            raise StorageError(f"Failed to save session {session.id}: {str(e)}")
    
    async def get_session(self, session_id: str) -> Optional[Session]:
        """
        Retrieve session metadata from file storage
        
        Args:
            session_id: The session identifier
            
        Returns:
            Optional[Session]: The session if found
        """
        try:
            session_file = self._get_session_file(session_id)
            
            if not session_file.exists():
                return None
            
            with open(session_file, 'r', encoding='utf-8') as f:
                session_data = json.load(f)
            
            session = Session(
                id=session_data["id"],
                created_at=datetime.fromisoformat(session_data["created_at"]),
                last_activity=datetime.fromisoformat(session_data["last_activity"]),
                status=SessionStatus(session_data["status"]),  # Convert string back to enum
                metadata=session_data.get("metadata", {}),
                message_count=session_data.get("message_count", 0)
            )
            
            return session
            
        except Exception as e:
            raise StorageError(f"Failed to retrieve session {session_id}: {str(e)}")
    
    async def delete_session(self, session_id: str) -> bool:
        """
        Delete a session and its messages from file storage
        
        Args:
            session_id: The session identifier
            
        Returns:
            bool: True if deletion was successful
        """
        try:
            deleted = False
            
            # Delete session file
            session_file = self._get_session_file(session_id)
            if session_file.exists():
                session_file.unlink()
                deleted = True
            
            # Delete conversation directory
            conversation_dir = self.conversations_dir / session_id
            if conversation_dir.exists():
                shutil.rmtree(conversation_dir)
                deleted = True
            
            return deleted
            
        except Exception as e:
            raise StorageError(f"Failed to delete session {session_id}: {str(e)}")
    
    async def backup_data(self) -> bool:
        """
        Create a backup of all conversation data
        
        Returns:
            bool: True if backup was successful
        """
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")  # Add microseconds for uniqueness
            backup_dir = self.backups_dir / timestamp
            backup_dir.mkdir(exist_ok=True)
            
            # Copy sessions directory
            if self.sessions_dir.exists():
                shutil.copytree(self.sessions_dir, backup_dir / "sessions")
            
            # Copy conversations directory
            if self.conversations_dir.exists():
                shutil.copytree(self.conversations_dir, backup_dir / "conversations")
            
            # Create backup metadata
            backup_metadata = {
                "timestamp": timestamp,
                "created_at": datetime.now().isoformat(),
                "data_dir": str(self.data_dir),
                "backup_type": "full"
            }
            
            metadata_file = backup_dir / "backup_metadata.json"
            content = json.dumps(backup_metadata, ensure_ascii=False, indent=2)
            await self._atomic_write(metadata_file, content)
            
            return True
            
        except Exception as e:
            raise StorageError(f"Failed to backup data: {str(e)}")
    
    async def restore_from_backup(self, backup_timestamp: str) -> bool:
        """
        Restore data from a backup
        
        Args:
            backup_timestamp: The timestamp of the backup to restore
            
        Returns:
            bool: True if restore was successful
        """
        try:
            backup_dir = self.backups_dir / backup_timestamp
            
            if not backup_dir.exists():
                raise StorageError(f"Backup {backup_timestamp} not found")
            
            # Verify backup integrity
            metadata_file = backup_dir / "backup_metadata.json"
            if not metadata_file.exists():
                raise StorageError(f"Backup metadata not found for {backup_timestamp}")
            
            # Create temporary backup of current data
            temp_backup = self.backups_dir / f"temp_restore_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            temp_backup.mkdir(exist_ok=True)
            
            if self.sessions_dir.exists():
                shutil.copytree(self.sessions_dir, temp_backup / "sessions")
            if self.conversations_dir.exists():
                shutil.copytree(self.conversations_dir, temp_backup / "conversations")
            
            try:
                # Clear current data
                if self.sessions_dir.exists():
                    shutil.rmtree(self.sessions_dir)
                if self.conversations_dir.exists():
                    shutil.rmtree(self.conversations_dir)
                
                # Restore from backup
                backup_sessions = backup_dir / "sessions"
                backup_conversations = backup_dir / "conversations"
                
                if backup_sessions.exists():
                    shutil.copytree(backup_sessions, self.sessions_dir)
                if backup_conversations.exists():
                    shutil.copytree(backup_conversations, self.conversations_dir)
                
                # Clean up temporary backup
                shutil.rmtree(temp_backup)
                
                return True
                
            except Exception as e:
                # Restore from temporary backup on failure
                if self.sessions_dir.exists():
                    shutil.rmtree(self.sessions_dir)
                if self.conversations_dir.exists():
                    shutil.rmtree(self.conversations_dir)
                
                temp_sessions = temp_backup / "sessions"
                temp_conversations = temp_backup / "conversations"
                
                if temp_sessions.exists():
                    shutil.copytree(temp_sessions, self.sessions_dir)
                if temp_conversations.exists():
                    shutil.copytree(temp_conversations, self.conversations_dir)
                
                shutil.rmtree(temp_backup)
                raise e
            
        except Exception as e:
            raise StorageError(f"Failed to restore from backup {backup_timestamp}: {str(e)}")
    
    def get_backup_list(self) -> List[Dict[str, Any]]:
        """
        Get list of available backups
        
        Returns:
            List[Dict[str, Any]]: List of backup information
        """
        backups = []
        
        if not self.backups_dir.exists():
            return backups
        
        for backup_dir in self.backups_dir.iterdir():
            if backup_dir.is_dir():
                metadata_file = backup_dir / "backup_metadata.json"
                
                if metadata_file.exists():
                    try:
                        with open(metadata_file, 'r', encoding='utf-8') as f:
                            metadata = json.load(f)
                        backups.append(metadata)
                    except Exception:
                        # Skip corrupted backup metadata
                        continue
                else:
                    # Backup without metadata (legacy)
                    backups.append({
                        "timestamp": backup_dir.name,
                        "created_at": None,
                        "backup_type": "unknown"
                    })
        
        # Sort by timestamp (newest first)
        backups.sort(key=lambda b: b["timestamp"], reverse=True)
        return backups
    
    def get_stats(self) -> Dict[str, Any]:
        """
        Get storage statistics
        
        Returns:
            Dict[str, Any]: Storage statistics
        """
        stats = {
            "total_sessions": 0,
            "total_messages": 0,
            "storage_size_bytes": 0,
            "backup_count": 0
        }
        
        try:
            # Count sessions
            if self.sessions_dir.exists():
                stats["total_sessions"] = len(list(self.sessions_dir.glob("*.json")))
            
            # Count messages and calculate storage size
            if self.conversations_dir.exists():
                for session_dir in self.conversations_dir.iterdir():
                    if session_dir.is_dir():
                        messages_file = session_dir / "messages.jsonl"
                        if messages_file.exists():
                            with open(messages_file, 'r', encoding='utf-8') as f:
                                stats["total_messages"] += sum(1 for line in f if line.strip())
            
            # Calculate total storage size
            for root, dirs, files in os.walk(self.data_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    if os.path.exists(file_path):
                        stats["storage_size_bytes"] += os.path.getsize(file_path)
            
            # Count backups
            if self.backups_dir.exists():
                stats["backup_count"] = len([d for d in self.backups_dir.iterdir() if d.is_dir()])
            
        except Exception as e:
            stats["error"] = str(e)
        
        return stats