"""
Backup and recovery manager for conversation data
"""

import json
import os
import shutil
import hashlib
from datetime import datetime, timedelta
from typing import List, Dict, Any, Optional
from pathlib import Path
import asyncio

from ..core.models import ChatConfig
from ..core.exceptions import StorageError
from .context_store import ContextStore


class BackupManager:
    """Manages backup and recovery operations for conversation data"""
    
    def __init__(self, context_store: ContextStore, config: ChatConfig):
        self.context_store = context_store
        self.config = config
        self.backup_dir = Path(config.data_dir) / "backups"
        self.backup_dir.mkdir(parents=True, exist_ok=True)
        
        # Backup configuration
        self.max_backups = getattr(config, 'max_backups', 10)
        self.backup_interval_hours = getattr(config, 'backup_interval_hours', 24)
        self.auto_backup_enabled = getattr(config, 'auto_backup_enabled', True)
        
        # Integrity check configuration
        self.verify_backups = getattr(config, 'verify_backups', True)
    
    async def create_backup(self, backup_type: str = "manual") -> Dict[str, Any]:
        """
        Create a backup of all conversation data
        
        Args:
            backup_type: Type of backup ("manual", "scheduled", "auto")
            
        Returns:
            Dict[str, Any]: Backup information
        """
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")  # Add microseconds for uniqueness
            backup_id = f"backup_{timestamp}"
            
            # Create backup using the storage backend
            success = await self.context_store.backup_data()
            
            if not success:
                raise StorageError("Storage backend backup failed")
            
            # Create backup metadata
            backup_info = {
                "backup_id": backup_id,
                "timestamp": timestamp,
                "created_at": datetime.now().isoformat(),
                "backup_type": backup_type,
                "storage_backend": self.config.storage_backend,
                "data_dir": self.config.data_dir,
                "status": "completed",
                "size_bytes": 0,
                "checksum": None,
                "verified": False
            }
            
            # Calculate backup size and checksum if possible
            if hasattr(self.context_store.storage, 'get_stats'):
                if asyncio.iscoroutinefunction(self.context_store.storage.get_stats):
                    stats = await self.context_store.storage.get_stats()
                else:
                    stats = self.context_store.storage.get_stats()
                
                backup_info["size_bytes"] = stats.get("storage_size_bytes", 0)
            
            # Verify backup integrity if enabled
            if self.verify_backups:
                backup_info["verified"] = await self._verify_backup_integrity(backup_id)
                backup_info["checksum"] = await self._calculate_backup_checksum(backup_id)
            
            # Save backup metadata
            metadata_file = self.backup_dir / f"{backup_id}_metadata.json"
            with open(metadata_file, 'w', encoding='utf-8') as f:
                json.dump(backup_info, f, ensure_ascii=False, indent=2)
            
            # Clean up old backups
            await self._cleanup_old_backups()
            
            return backup_info
            
        except Exception as e:
            raise StorageError(f"Failed to create backup: {str(e)}")
    
    async def restore_from_backup(self, backup_id: str, verify_before_restore: bool = True) -> bool:
        """
        Restore data from a backup
        
        Args:
            backup_id: The backup identifier to restore from
            verify_before_restore: Whether to verify backup integrity before restore
            
        Returns:
            bool: True if restore was successful
        """
        try:
            # Load backup metadata
            metadata_file = self.backup_dir / f"{backup_id}_metadata.json"
            
            if not metadata_file.exists():
                raise StorageError(f"Backup metadata not found for {backup_id}")
            
            with open(metadata_file, 'r', encoding='utf-8') as f:
                backup_info = json.load(f)
            
            # Verify backup integrity if requested
            if verify_before_restore:
                if not await self._verify_backup_integrity(backup_id):
                    raise StorageError(f"Backup {backup_id} failed integrity check")
            
            # Extract timestamp from backup_id for storage backend
            timestamp = backup_id.replace("backup_", "")
            
            # Restore using the storage backend
            if hasattr(self.context_store.storage, 'restore_from_backup'):
                success = await self.context_store.storage.restore_from_backup(timestamp)
                
                if success:
                    # Clear context store cache after restore
                    self.context_store.memory_cache.clear()
                    
                    # Update backup metadata with restore information
                    backup_info["last_restored"] = datetime.now().isoformat()
                    backup_info["restore_count"] = backup_info.get("restore_count", 0) + 1
                    
                    with open(metadata_file, 'w', encoding='utf-8') as f:
                        json.dump(backup_info, f, ensure_ascii=False, indent=2)
                    
                    return True
                else:
                    raise StorageError("Storage backend restore failed")
            else:
                raise StorageError("Storage backend does not support restore operations")
            
        except Exception as e:
            raise StorageError(f"Failed to restore from backup {backup_id}: {str(e)}")
    
    def list_backups(self) -> List[Dict[str, Any]]:
        """
        Get list of available backups
        
        Returns:
            List[Dict[str, Any]]: List of backup information
        """
        backups = []
        
        # Get backups from storage backend
        if hasattr(self.context_store.storage, 'get_backup_list'):
            storage_backups = self.context_store.storage.get_backup_list()
            backups.extend(storage_backups)
        
        # Get backups from backup manager metadata
        for metadata_file in self.backup_dir.glob("backup_*_metadata.json"):
            try:
                with open(metadata_file, 'r', encoding='utf-8') as f:
                    backup_info = json.load(f)
                
                # Check if this backup is already in the list
                if not any(b.get("backup_id") == backup_info.get("backup_id") for b in backups):
                    backups.append(backup_info)
                    
            except Exception:
                # Skip corrupted metadata files
                continue
        
        # Sort by timestamp (newest first)
        backups.sort(key=lambda b: b.get("timestamp", ""), reverse=True)
        return backups
    
    async def delete_backup(self, backup_id: str) -> bool:
        """
        Delete a backup
        
        Args:
            backup_id: The backup identifier to delete
            
        Returns:
            bool: True if deletion was successful
        """
        try:
            deleted = False
            
            # Delete backup metadata
            metadata_file = self.backup_dir / f"{backup_id}_metadata.json"
            if metadata_file.exists():
                metadata_file.unlink()
                deleted = True
            
            # Delete backup from storage backend if supported
            timestamp = backup_id.replace("backup_", "")
            if hasattr(self.context_store.storage, 'delete_backup'):
                backend_deleted = await self.context_store.storage.delete_backup(timestamp)
                deleted = deleted or backend_deleted
            
            return deleted
            
        except Exception as e:
            raise StorageError(f"Failed to delete backup {backup_id}: {str(e)}")
    
    async def schedule_auto_backup(self) -> None:
        """Schedule automatic backup if enabled"""
        if not self.auto_backup_enabled:
            return
        
        try:
            # Check if it's time for a backup
            last_backup = await self._get_last_backup_time()
            
            if last_backup is None or (datetime.now() - last_backup).total_seconds() > (self.backup_interval_hours * 3600):
                await self.create_backup(backup_type="scheduled")
                
        except Exception as e:
            # Log error but don't raise - auto backup failures shouldn't break the system
            print(f"Auto backup failed: {str(e)}")
    
    async def verify_all_backups(self) -> Dict[str, bool]:
        """
        Verify integrity of all backups
        
        Returns:
            Dict[str, bool]: Backup ID to verification result mapping
        """
        results = {}
        backups = self.list_backups()
        
        for backup in backups:
            backup_id = backup.get("backup_id")
            if backup_id:
                try:
                    results[backup_id] = await self._verify_backup_integrity(backup_id)
                except Exception:
                    results[backup_id] = False
        
        return results
    
    async def get_backup_stats(self) -> Dict[str, Any]:
        """
        Get backup system statistics
        
        Returns:
            Dict[str, Any]: Backup statistics
        """
        backups = self.list_backups()
        
        total_size = sum(b.get("size_bytes", 0) for b in backups)
        verified_count = sum(1 for b in backups if b.get("verified", False))
        
        stats = {
            "total_backups": len(backups),
            "verified_backups": verified_count,
            "total_size_bytes": total_size,
            "total_size_mb": round(total_size / (1024 * 1024), 2),
            "auto_backup_enabled": self.auto_backup_enabled,
            "backup_interval_hours": self.backup_interval_hours,
            "max_backups": self.max_backups,
            "storage_backend": self.config.storage_backend
        }
        
        if backups:
            latest_backup = backups[0]
            stats["latest_backup"] = {
                "backup_id": latest_backup.get("backup_id"),
                "created_at": latest_backup.get("created_at"),
                "backup_type": latest_backup.get("backup_type"),
                "verified": latest_backup.get("verified", False)
            }
        
        return stats
    
    async def _verify_backup_integrity(self, backup_id: str) -> bool:
        """
        Verify backup integrity
        
        Args:
            backup_id: The backup identifier to verify
            
        Returns:
            bool: True if backup is valid
        """
        try:
            # Load backup metadata
            metadata_file = self.backup_dir / f"{backup_id}_metadata.json"
            
            if not metadata_file.exists():
                return False
            
            with open(metadata_file, 'r', encoding='utf-8') as f:
                backup_info = json.load(f)
            
            # Basic metadata validation
            required_fields = ["backup_id", "timestamp", "created_at", "storage_backend"]
            if not all(field in backup_info for field in required_fields):
                return False
            
            # Verify backup exists in storage backend
            timestamp = backup_id.replace("backup_", "")
            if hasattr(self.context_store.storage, 'get_backup_list'):
                storage_backups = self.context_store.storage.get_backup_list()
                if not any(b.get("timestamp") == timestamp for b in storage_backups):
                    return False
            
            # Additional integrity checks could be added here
            # (e.g., checksum verification, file existence checks)
            
            return True
            
        except Exception:
            return False
    
    async def _calculate_backup_checksum(self, backup_id: str) -> Optional[str]:
        """
        Calculate checksum for backup verification
        
        Args:
            backup_id: The backup identifier
            
        Returns:
            Optional[str]: Backup checksum or None if calculation fails
        """
        try:
            # This is a simplified checksum calculation
            # In a real implementation, you would calculate checksums of actual backup files
            
            metadata_file = self.backup_dir / f"{backup_id}_metadata.json"
            if not metadata_file.exists():
                return None
            
            with open(metadata_file, 'rb') as f:
                content = f.read()
            
            return hashlib.sha256(content).hexdigest()
            
        except Exception:
            return None
    
    async def _get_last_backup_time(self) -> Optional[datetime]:
        """
        Get the timestamp of the last backup
        
        Returns:
            Optional[datetime]: Last backup time or None if no backups exist
        """
        backups = self.list_backups()
        
        if not backups:
            return None
        
        latest_backup = backups[0]  # Already sorted by timestamp
        created_at = latest_backup.get("created_at")
        
        if created_at:
            try:
                return datetime.fromisoformat(created_at)
            except ValueError:
                pass
        
        return None
    
    async def _cleanup_old_backups(self) -> None:
        """Clean up old backups to maintain the maximum backup count"""
        try:
            backups = self.list_backups()
            
            # Only count backup manager backups (those with backup_id)
            manager_backups = [b for b in backups if b.get("backup_id")]
            
            if len(manager_backups) <= self.max_backups:
                return
            
            # Delete oldest backups
            backups_to_delete = manager_backups[self.max_backups:]
            
            for backup in backups_to_delete:
                backup_id = backup.get("backup_id")
                if backup_id:
                    try:
                        await self.delete_backup(backup_id)
                    except Exception as e:
                        print(f"Failed to delete old backup {backup_id}: {str(e)}")
            
        except Exception as e:
            print(f"Failed to cleanup old backups: {str(e)}")
    
    async def health_check(self) -> Dict[str, Any]:
        """
        Perform health check on backup system
        
        Returns:
            Dict[str, Any]: Health status information
        """
        try:
            health_status = {
                "status": "healthy",
                "backup_dir_exists": self.backup_dir.exists(),
                "backup_dir_writable": os.access(self.backup_dir, os.W_OK),
                "auto_backup_enabled": self.auto_backup_enabled,
                "storage_backend_supports_backup": hasattr(self.context_store.storage, 'backup_data'),
                "storage_backend_supports_restore": hasattr(self.context_store.storage, 'restore_from_backup'),
                "total_backups": len(self.list_backups()),
                "error": None
            }
            
            # Check if backup directory is accessible
            if not health_status["backup_dir_exists"]:
                health_status["status"] = "warning"
                health_status["error"] = "Backup directory does not exist"
            elif not health_status["backup_dir_writable"]:
                health_status["status"] = "warning"
                health_status["error"] = "Backup directory is not writable"
            
            # Check storage backend support
            if not health_status["storage_backend_supports_backup"]:
                health_status["status"] = "warning"
                health_status["error"] = "Storage backend does not support backup operations"
            
            return health_status
            
        except Exception as e:
            return {
                "status": "error",
                "error": str(e)
            }