"""
Database storage backend for persistent conversation data using SQLite
"""

import sqlite3
import json
import os
from datetime import datetime
from typing import List, Optional, Dict, Any
from pathlib import Path
import aiosqlite

from ..core.models import Message, Session, MessageRole, SessionStatus
from ..core.interfaces import StorageBackend
from ..core.exceptions import StorageError


class DatabaseStorage(StorageBackend):
    """SQLite database storage implementation"""
    
    def __init__(self, db_path: str = "ai/chat/data/conversations.db"):
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Initialize database schema
        self._init_database()
    
    def _init_database(self):
        """Initialize database schema"""
        try:
            with sqlite3.connect(self.db_path) as conn:
                cursor = conn.cursor()
                
                # Create sessions table
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS sessions (
                        id TEXT PRIMARY KEY,
                        created_at TEXT NOT NULL,
                        last_activity TEXT NOT NULL,
                        status TEXT NOT NULL,
                        metadata TEXT,
                        message_count INTEGER DEFAULT 0
                    )
                ''')
                
                # Create messages table
                cursor.execute('''
                    CREATE TABLE IF NOT EXISTS messages (
                        id TEXT PRIMARY KEY,
                        session_id TEXT NOT NULL,
                        content TEXT NOT NULL,
                        role TEXT NOT NULL,
                        timestamp TEXT NOT NULL,
                        metadata TEXT,
                        FOREIGN KEY (session_id) REFERENCES sessions (id) ON DELETE CASCADE
                    )
                ''')
                
                # Create indexes for better performance
                cursor.execute('''
                    CREATE INDEX IF NOT EXISTS idx_messages_session_id 
                    ON messages (session_id)
                ''')
                
                cursor.execute('''
                    CREATE INDEX IF NOT EXISTS idx_messages_timestamp 
                    ON messages (timestamp)
                ''')
                
                cursor.execute('''
                    CREATE INDEX IF NOT EXISTS idx_sessions_status 
                    ON sessions (status)
                ''')
                
                cursor.execute('''
                    CREATE INDEX IF NOT EXISTS idx_sessions_last_activity 
                    ON sessions (last_activity)
                ''')
                
                conn.commit()
                
        except Exception as e:
            raise StorageError(f"Failed to initialize database: {str(e)}")
    
    async def save_message(self, session_id: str, message: Message) -> None:
        """
        Save a message to database storage
        
        Args:
            session_id: The session identifier
            message: The message to save
        """
        try:
            async with aiosqlite.connect(self.db_path) as db:
                await db.execute('''
                    INSERT OR REPLACE INTO messages 
                    (id, session_id, content, role, timestamp, metadata)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (
                    message.id,
                    message.session_id,
                    message.content,
                    message.role.value,
                    message.timestamp.isoformat(),
                    json.dumps(message.metadata, ensure_ascii=False)
                ))
                
                await db.commit()
                
        except Exception as e:
            raise StorageError(f"Failed to save message for session {session_id}: {str(e)}")
    
    async def get_messages(self, session_id: str, limit: Optional[int] = None) -> List[Message]:
        """
        Retrieve messages from database storage
        
        Args:
            session_id: The session identifier
            limit: Maximum number of messages to retrieve
            
        Returns:
            List[Message]: List of messages in chronological order
        """
        try:
            async with aiosqlite.connect(self.db_path) as db:
                if limit:
                    # Get the most recent messages
                    cursor = await db.execute('''
                        SELECT id, session_id, content, role, timestamp, metadata
                        FROM messages 
                        WHERE session_id = ?
                        ORDER BY timestamp DESC
                        LIMIT ?
                    ''', (session_id, limit))
                    
                    rows = await cursor.fetchall()
                    # Reverse to get chronological order
                    rows = list(reversed(rows))
                else:
                    cursor = await db.execute('''
                        SELECT id, session_id, content, role, timestamp, metadata
                        FROM messages 
                        WHERE session_id = ?
                        ORDER BY timestamp ASC
                    ''', (session_id,))
                    
                    rows = await cursor.fetchall()
                
                messages = []
                for row in rows:
                    try:
                        metadata = json.loads(row[5]) if row[5] else {}
                    except json.JSONDecodeError:
                        metadata = {}
                    
                    message = Message(
                        id=row[0],
                        session_id=row[1],
                        content=row[2],
                        role=MessageRole(row[3]),
                        timestamp=datetime.fromisoformat(row[4]),
                        metadata=metadata
                    )
                    messages.append(message)
                
                return messages
                
        except Exception as e:
            raise StorageError(f"Failed to retrieve messages for session {session_id}: {str(e)}")
    
    async def save_session(self, session: Session) -> None:
        """
        Save session metadata to database storage
        
        Args:
            session: The session to save
        """
        try:
            async with aiosqlite.connect(self.db_path) as db:
                await db.execute('''
                    INSERT OR REPLACE INTO sessions 
                    (id, created_at, last_activity, status, metadata, message_count)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (
                    session.id,
                    session.created_at.isoformat(),
                    session.last_activity.isoformat(),
                    session.status.value,
                    json.dumps(session.metadata, ensure_ascii=False),
                    session.message_count
                ))
                
                await db.commit()
                
        except Exception as e:
            raise StorageError(f"Failed to save session {session.id}: {str(e)}")
    
    async def get_session(self, session_id: str) -> Optional[Session]:
        """
        Retrieve session metadata from database storage
        
        Args:
            session_id: The session identifier
            
        Returns:
            Optional[Session]: The session if found
        """
        try:
            async with aiosqlite.connect(self.db_path) as db:
                cursor = await db.execute('''
                    SELECT id, created_at, last_activity, status, metadata, message_count
                    FROM sessions 
                    WHERE id = ?
                ''', (session_id,))
                
                row = await cursor.fetchone()
                
                if not row:
                    return None
                
                try:
                    metadata = json.loads(row[4]) if row[4] else {}
                except json.JSONDecodeError:
                    metadata = {}
                
                session = Session(
                    id=row[0],
                    created_at=datetime.fromisoformat(row[1]),
                    last_activity=datetime.fromisoformat(row[2]),
                    status=SessionStatus(row[3]),
                    metadata=metadata,
                    message_count=row[5] or 0
                )
                
                return session
                
        except Exception as e:
            raise StorageError(f"Failed to retrieve session {session_id}: {str(e)}")
    
    async def delete_session(self, session_id: str) -> bool:
        """
        Delete a session and its messages from database storage
        
        Args:
            session_id: The session identifier
            
        Returns:
            bool: True if deletion was successful
        """
        try:
            async with aiosqlite.connect(self.db_path) as db:
                # Delete messages first (foreign key constraint)
                await db.execute('DELETE FROM messages WHERE session_id = ?', (session_id,))
                
                # Delete session
                cursor = await db.execute('DELETE FROM sessions WHERE id = ?', (session_id,))
                
                await db.commit()
                
                # Return True if any rows were affected
                return cursor.rowcount > 0
                
        except Exception as e:
            raise StorageError(f"Failed to delete session {session_id}: {str(e)}")
    
    async def backup_data(self) -> bool:
        """
        Create a backup of the database
        
        Returns:
            bool: True if backup was successful
        """
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_dir = self.db_path.parent / "backups"
            backup_dir.mkdir(exist_ok=True)
            
            backup_path = backup_dir / f"conversations_backup_{timestamp}.db"
            
            # Copy database file
            async with aiosqlite.connect(self.db_path) as source:
                async with aiosqlite.connect(backup_path) as backup:
                    await source.backup(backup)
            
            # Create backup metadata
            metadata_path = backup_dir / f"backup_metadata_{timestamp}.json"
            backup_metadata = {
                "timestamp": timestamp,
                "created_at": datetime.now().isoformat(),
                "original_db": str(self.db_path),
                "backup_db": str(backup_path),
                "backup_type": "full"
            }
            
            with open(metadata_path, 'w', encoding='utf-8') as f:
                json.dump(backup_metadata, f, ensure_ascii=False, indent=2)
            
            return True
            
        except Exception as e:
            raise StorageError(f"Failed to backup database: {str(e)}")
    
    async def restore_from_backup(self, backup_timestamp: str) -> bool:
        """
        Restore database from a backup
        
        Args:
            backup_timestamp: The timestamp of the backup to restore
            
        Returns:
            bool: True if restore was successful
        """
        try:
            backup_dir = self.db_path.parent / "backups"
            backup_path = backup_dir / f"conversations_backup_{backup_timestamp}.db"
            
            if not backup_path.exists():
                raise StorageError(f"Backup file not found: {backup_path}")
            
            # Create temporary backup of current database
            temp_backup = self.db_path.with_suffix('.db.temp')
            if self.db_path.exists():
                async with aiosqlite.connect(self.db_path) as source:
                    async with aiosqlite.connect(temp_backup) as temp:
                        await source.backup(temp)
            
            try:
                # Remove current database
                if self.db_path.exists():
                    os.remove(self.db_path)
                
                # Restore from backup
                async with aiosqlite.connect(backup_path) as source:
                    async with aiosqlite.connect(self.db_path) as target:
                        await source.backup(target)
                
                # Clean up temporary backup
                if temp_backup.exists():
                    os.remove(temp_backup)
                
                return True
                
            except Exception as e:
                # Restore from temporary backup on failure
                if temp_backup.exists():
                    if self.db_path.exists():
                        os.remove(self.db_path)
                    
                    async with aiosqlite.connect(temp_backup) as source:
                        async with aiosqlite.connect(self.db_path) as target:
                            await source.backup(target)
                    
                    os.remove(temp_backup)
                
                raise e
            
        except Exception as e:
            raise StorageError(f"Failed to restore from backup {backup_timestamp}: {str(e)}")
    
    def get_backup_list(self) -> List[Dict[str, Any]]:
        """
        Get list of available database backups
        
        Returns:
            List[Dict[str, Any]]: List of backup information
        """
        backups = []
        backup_dir = self.db_path.parent / "backups"
        
        if not backup_dir.exists():
            return backups
        
        # Look for backup metadata files
        for metadata_file in backup_dir.glob("backup_metadata_*.json"):
            try:
                with open(metadata_file, 'r', encoding='utf-8') as f:
                    metadata = json.load(f)
                backups.append(metadata)
            except Exception:
                # Skip corrupted metadata files
                continue
        
        # Look for backup database files without metadata (legacy)
        for backup_file in backup_dir.glob("conversations_backup_*.db"):
            timestamp = backup_file.stem.replace("conversations_backup_", "")
            
            # Check if we already have metadata for this backup
            if not any(b["timestamp"] == timestamp for b in backups):
                backups.append({
                    "timestamp": timestamp,
                    "created_at": None,
                    "backup_type": "unknown",
                    "backup_db": str(backup_file)
                })
        
        # Sort by timestamp (newest first)
        backups.sort(key=lambda b: b["timestamp"], reverse=True)
        return backups
    
    async def get_stats(self) -> Dict[str, Any]:
        """
        Get database storage statistics
        
        Returns:
            Dict[str, Any]: Storage statistics
        """
        stats = {
            "total_sessions": 0,
            "total_messages": 0,
            "active_sessions": 0,
            "inactive_sessions": 0,
            "ended_sessions": 0,
            "database_size_bytes": 0,
            "backup_count": 0
        }
        
        try:
            async with aiosqlite.connect(self.db_path) as db:
                # Count sessions by status
                cursor = await db.execute('''
                    SELECT status, COUNT(*) 
                    FROM sessions 
                    GROUP BY status
                ''')
                
                status_counts = await cursor.fetchall()
                for status, count in status_counts:
                    if status == "active":
                        stats["active_sessions"] = count
                    elif status == "inactive":
                        stats["inactive_sessions"] = count
                    elif status == "ended":
                        stats["ended_sessions"] = count
                
                stats["total_sessions"] = sum([
                    stats["active_sessions"],
                    stats["inactive_sessions"],
                    stats["ended_sessions"]
                ])
                
                # Count total messages
                cursor = await db.execute('SELECT COUNT(*) FROM messages')
                result = await cursor.fetchone()
                stats["total_messages"] = result[0] if result else 0
            
            # Get database file size
            if self.db_path.exists():
                stats["database_size_bytes"] = os.path.getsize(self.db_path)
            
            # Count backups
            backup_dir = self.db_path.parent / "backups"
            if backup_dir.exists():
                stats["backup_count"] = len(list(backup_dir.glob("conversations_backup_*.db")))
            
        except Exception as e:
            stats["error"] = str(e)
        
        return stats
    
    async def vacuum_database(self) -> bool:
        """
        Optimize database by running VACUUM command
        
        Returns:
            bool: True if vacuum was successful
        """
        try:
            async with aiosqlite.connect(self.db_path) as db:
                await db.execute('VACUUM')
                await db.commit()
            return True
        except Exception as e:
            raise StorageError(f"Failed to vacuum database: {str(e)}")
    
    async def get_session_count_by_date(self, days: int = 30) -> Dict[str, int]:
        """
        Get session creation count by date for the last N days
        
        Args:
            days: Number of days to look back
            
        Returns:
            Dict[str, int]: Date to session count mapping
        """
        try:
            async with aiosqlite.connect(self.db_path) as db:
                cursor = await db.execute('''
                    SELECT DATE(created_at) as date, COUNT(*) as count
                    FROM sessions 
                    WHERE created_at >= datetime('now', '-{} days')
                    GROUP BY DATE(created_at)
                    ORDER BY date
                '''.format(days))
                
                rows = await cursor.fetchall()
                return {row[0]: row[1] for row in rows}
                
        except Exception as e:
            raise StorageError(f"Failed to get session count by date: {str(e)}")