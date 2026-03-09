"""
Context store for managing conversation data persistence
"""

from typing import List, Optional, Dict, Any
from ..core.models import Message, Session, ChatConfig
from ..core.interfaces import StorageBackend
from .memory_storage import MemoryStorage
from .file_storage import FileStorage
from .database_storage import DatabaseStorage


class ContextStore:
    """Manages conversation context storage"""
    
    def __init__(self, config: ChatConfig):
        self.config = config
        
        # Initialize storage backend based on configuration
        if config.storage_backend == "memory":
            self.storage = MemoryStorage()
        elif config.storage_backend == "file":
            self.storage = FileStorage(config.data_dir)
        elif config.storage_backend == "database":
            db_path = f"{config.data_dir}/conversations.db"
            self.storage = DatabaseStorage(db_path)
        else:
            # Default to memory storage for unknown backends
            self.storage = MemoryStorage()
        
        # Memory cache for frequently accessed data
        self.memory_cache: Dict[str, List[Message]] = {}
        self.cache_size_limit = 100  # Maximum sessions to cache
        
        # Initialize backup manager (lazy loading to avoid circular imports)
        self._backup_manager = None
    
    @property
    def backup_manager(self):
        """Get backup manager instance (lazy loading)"""
        if self._backup_manager is None:
            from .backup_manager import BackupManager
            self._backup_manager = BackupManager(self, self.config)
        return self._backup_manager
    
    async def save_message(self, session_id: str, message: Message) -> None:
        """
        Save a message to storage
        
        Args:
            session_id: The session identifier
            message: The message to save
        """
        # Save to persistent storage
        await self.storage.save_message(session_id, message)
        
        # Update memory cache
        if session_id not in self.memory_cache:
            self.memory_cache[session_id] = []
        
        self.memory_cache[session_id].append(message)
        
        # Manage cache size
        self._manage_cache_size()
    
    async def get_messages(self, session_id: str, limit: Optional[int] = None) -> List[Message]:
        """
        Retrieve messages from storage
        
        Args:
            session_id: The session identifier
            limit: Maximum number of messages to retrieve
            
        Returns:
            List[Message]: List of messages
        """
        # Try cache first
        if session_id in self.memory_cache:
            messages = self.memory_cache[session_id]
            if limit:
                messages = messages[-limit:]  # Get most recent messages
            return messages.copy()
        
        # Fallback to persistent storage
        messages = await self.storage.get_messages(session_id, limit)
        
        # Update cache
        if len(self.memory_cache) < self.cache_size_limit:
            self.memory_cache[session_id] = messages.copy()
        
        return messages
    
    async def save_session(self, session: Session) -> None:
        """
        Save session metadata
        
        Args:
            session: The session to save
        """
        await self.storage.save_session(session)
    
    async def get_session(self, session_id: str) -> Optional[Session]:
        """
        Retrieve session metadata
        
        Args:
            session_id: The session identifier
            
        Returns:
            Optional[Session]: The session if found
        """
        return await self.storage.get_session(session_id)
    
    async def delete_session(self, session_id: str) -> bool:
        """
        Delete a session and its messages
        
        Args:
            session_id: The session identifier
            
        Returns:
            bool: True if deletion was successful
        """
        # Remove from cache
        if session_id in self.memory_cache:
            del self.memory_cache[session_id]
        
        # Delete from persistent storage
        return await self.storage.delete_session(session_id)
    
    async def backup_data(self) -> bool:
        """
        Backup all conversation data
        
        Returns:
            bool: True if backup was successful
        """
        return await self.storage.backup_data()
    
    async def restore_from_backup(self, backup_timestamp: str) -> bool:
        """
        Restore data from a backup
        
        Args:
            backup_timestamp: The timestamp of the backup to restore
            
        Returns:
            bool: True if restore was successful
        """
        if hasattr(self.storage, 'restore_from_backup'):
            # Clear memory cache since data is being restored
            self.memory_cache.clear()
            return await self.storage.restore_from_backup(backup_timestamp)
        return False
    
    def get_backup_list(self) -> List[Dict[str, Any]]:
        """
        Get list of available backups
        
        Returns:
            List[Dict[str, Any]]: List of backup information
        """
        return self.backup_manager.list_backups()
    
    async def create_backup(self, backup_type: str = "manual") -> Dict[str, Any]:
        """
        Create a backup of all conversation data
        
        Args:
            backup_type: Type of backup ("manual", "scheduled", "auto")
            
        Returns:
            Dict[str, Any]: Backup information
        """
        return await self.backup_manager.create_backup(backup_type)
    
    async def restore_backup(self, backup_id: str, verify_before_restore: bool = True) -> bool:
        """
        Restore data from a backup
        
        Args:
            backup_id: The backup identifier to restore from
            verify_before_restore: Whether to verify backup integrity before restore
            
        Returns:
            bool: True if restore was successful
        """
        return await self.backup_manager.restore_from_backup(backup_id, verify_before_restore)
    
    async def schedule_auto_backup(self) -> None:
        """Schedule automatic backup if enabled"""
        await self.backup_manager.schedule_auto_backup()
    
    def _manage_cache_size(self):
        """Manage memory cache size to prevent excessive memory usage"""
        if len(self.memory_cache) > self.cache_size_limit:
            # Remove oldest entries (simple LRU-like behavior)
            # In a real implementation, this would be more sophisticated
            oldest_session = next(iter(self.memory_cache))
            del self.memory_cache[oldest_session]
    
    async def health_check(self) -> Dict[str, Any]:
        """
        Perform health check on storage system
        
        Returns:
            Dict[str, Any]: Health status information
        """
        try:
            # Check storage backend health
            storage_healthy = True
            error_message = None
            
            try:
                # Try a simple operation to test storage
                test_session = Session(id="health_check_test")
                await self.storage.save_session(test_session)
                await self.storage.get_session("health_check_test")
                await self.storage.delete_session("health_check_test")
            except Exception as e:
                storage_healthy = False
                error_message = str(e)
            
            # Check backup manager health
            backup_health = await self.backup_manager.health_check()
            
            overall_status = "healthy"
            if not storage_healthy or backup_health["status"] != "healthy":
                overall_status = "unhealthy" if not storage_healthy else backup_health["status"]
            
            return {
                "status": overall_status,
                "storage_backend": self.config.storage_backend,
                "cache_size": len(self.memory_cache),
                "cache_limit": self.cache_size_limit,
                "storage_error": error_message,
                "backup_system": backup_health
            }
        except Exception as e:
            return {
                "status": "error",
                "error": str(e)
            }