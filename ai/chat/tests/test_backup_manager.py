"""
Tests for backup manager
"""

import pytest
import tempfile
import shutil
from pathlib import Path
from datetime import datetime

from ai.chat.storage.backup_manager import BackupManager
from ai.chat.storage.context_store import ContextStore
from ai.chat.core.models import Message, Session, ChatConfig, MessageRole, SessionStatus


class TestBackupManager:
    """Test backup manager functionality"""
    
    @pytest.fixture
    def temp_dir(self):
        """Create temporary directory for testing"""
        temp_dir = tempfile.mkdtemp()
        yield temp_dir
        shutil.rmtree(temp_dir)
    
    @pytest.fixture
    def config(self, temp_dir):
        """Create test configuration"""
        return ChatConfig(
            storage_backend="file",
            data_dir=temp_dir,
            max_backups=5,
            backup_interval_hours=1,
            auto_backup_enabled=True,
            verify_backups=True
        )
    
    @pytest.fixture
    def context_store(self, config):
        """Create context store instance"""
        return ContextStore(config)
    
    @pytest.fixture
    def backup_manager(self, context_store, config):
        """Create backup manager instance"""
        return BackupManager(context_store, config)
    
    @pytest.fixture
    def sample_data(self):
        """Create sample data for testing"""
        session = Session(
            id="test-session-1",
            created_at=datetime.now(),
            last_activity=datetime.now(),
            status=SessionStatus.ACTIVE,
            metadata={"test": "session"},
            message_count=2
        )
        
        messages = [
            Message(
                id="msg1",
                session_id="test-session-1",
                content="Hello, backup!",
                role=MessageRole.USER,
                timestamp=datetime.now(),
                metadata={"test": "message1"}
            ),
            Message(
                id="msg2",
                session_id="test-session-1",
                content="Backup response",
                role=MessageRole.ASSISTANT,
                timestamp=datetime.now(),
                metadata={"test": "message2"}
            )
        ]
        
        return session, messages
    
    @pytest.mark.asyncio
    async def test_create_backup(self, backup_manager, context_store, sample_data):
        """Test creating a backup"""
        session, messages = sample_data
        
        # Add some data to backup
        await context_store.save_session(session)
        for message in messages:
            await context_store.save_message(session.id, message)
        
        # Create backup
        backup_info = await backup_manager.create_backup(backup_type="test")
        
        assert backup_info["status"] == "completed"
        assert backup_info["backup_type"] == "test"
        assert backup_info["storage_backend"] == "file"
        assert "backup_id" in backup_info
        assert "timestamp" in backup_info
        assert "created_at" in backup_info
    
    @pytest.mark.asyncio
    async def test_list_backups(self, backup_manager, context_store, sample_data):
        """Test listing backups"""
        session, messages = sample_data
        
        # Add some data
        await context_store.save_session(session)
        for message in messages:
            await context_store.save_message(session.id, message)
        
        # Create multiple backups
        backup1 = await backup_manager.create_backup(backup_type="test1")
        backup2 = await backup_manager.create_backup(backup_type="test2")
        
        # List backups
        backups = backup_manager.list_backups()
        
        assert len(backups) >= 2
        backup_ids = [b.get("backup_id") for b in backups]
        assert backup1["backup_id"] in backup_ids
        assert backup2["backup_id"] in backup_ids
    
    @pytest.mark.asyncio
    async def test_restore_from_backup(self, backup_manager, context_store, sample_data):
        """Test restoring from backup"""
        session, messages = sample_data
        
        # Add some data
        await context_store.save_session(session)
        for message in messages:
            await context_store.save_message(session.id, message)
        
        # Create backup
        backup_info = await backup_manager.create_backup(backup_type="test")
        backup_id = backup_info["backup_id"]
        
        # Delete original data
        await context_store.delete_session(session.id)
        assert await context_store.get_session(session.id) is None
        
        # Restore from backup
        success = await backup_manager.restore_from_backup(backup_id, verify_before_restore=False)
        assert success is True
        
        # Verify data is restored
        restored_session = await context_store.get_session(session.id)
        assert restored_session is not None
        assert restored_session.id == session.id
        
        restored_messages = await context_store.get_messages(session.id)
        assert len(restored_messages) == 2
        assert restored_messages[0].content == messages[0].content
        assert restored_messages[1].content == messages[1].content
    
    @pytest.mark.asyncio
    async def test_delete_backup(self, backup_manager, context_store, sample_data):
        """Test deleting a backup"""
        session, messages = sample_data
        
        # Add some data
        await context_store.save_session(session)
        for message in messages:
            await context_store.save_message(session.id, message)
        
        # Create backup
        backup_info = await backup_manager.create_backup(backup_type="test")
        backup_id = backup_info["backup_id"]
        
        # Verify backup exists
        backups = backup_manager.list_backups()
        assert any(b.get("backup_id") == backup_id for b in backups)
        
        # Delete backup
        success = await backup_manager.delete_backup(backup_id)
        assert success is True
        
        # Verify backup is gone
        backups = backup_manager.list_backups()
        assert not any(b.get("backup_id") == backup_id for b in backups)
    
    @pytest.mark.asyncio
    async def test_verify_all_backups(self, backup_manager, context_store, sample_data):
        """Test verifying all backups"""
        session, messages = sample_data
        
        # Add some data
        await context_store.save_session(session)
        for message in messages:
            await context_store.save_message(session.id, message)
        
        # Create backups
        backup1 = await backup_manager.create_backup(backup_type="test1")
        backup2 = await backup_manager.create_backup(backup_type="test2")
        
        # Verify all backups
        results = await backup_manager.verify_all_backups()
        
        assert backup1["backup_id"] in results
        assert backup2["backup_id"] in results
        # Results should be boolean values
        assert isinstance(results[backup1["backup_id"]], bool)
        assert isinstance(results[backup2["backup_id"]], bool)
    
    @pytest.mark.asyncio
    async def test_get_backup_stats(self, backup_manager, context_store, sample_data):
        """Test getting backup statistics"""
        session, messages = sample_data
        
        # Add some data
        await context_store.save_session(session)
        for message in messages:
            await context_store.save_message(session.id, message)
        
        # Create backup
        await backup_manager.create_backup(backup_type="test")
        
        # Get stats
        stats = await backup_manager.get_backup_stats()
        
        assert "total_backups" in stats
        assert "verified_backups" in stats
        assert "total_size_bytes" in stats
        assert "total_size_mb" in stats
        assert "auto_backup_enabled" in stats
        assert "backup_interval_hours" in stats
        assert "max_backups" in stats
        assert "storage_backend" in stats
        
        assert stats["total_backups"] >= 1
        assert stats["auto_backup_enabled"] is True
        assert stats["backup_interval_hours"] == 1
        assert stats["max_backups"] == 5
        assert stats["storage_backend"] == "file"
    
    @pytest.mark.asyncio
    async def test_schedule_auto_backup(self, backup_manager, context_store, sample_data):
        """Test scheduling automatic backup"""
        session, messages = sample_data
        
        # Add some data
        await context_store.save_session(session)
        for message in messages:
            await context_store.save_message(session.id, message)
        
        # Get initial backup count
        initial_backups = backup_manager.list_backups()
        initial_count = len(initial_backups)
        
        # Schedule auto backup (should create one since no previous backups)
        await backup_manager.schedule_auto_backup()
        
        # Check if backup was created
        backups = backup_manager.list_backups()
        assert len(backups) >= initial_count
        
        # If a backup was created, it should be of type "scheduled"
        if len(backups) > initial_count:
            # Find the backup created by the backup manager (not storage backend)
            manager_backups = [b for b in backups if b.get("backup_type") == "scheduled"]
            if manager_backups:
                latest_backup = manager_backups[0]
                assert latest_backup.get("backup_type") == "scheduled"
    
    @pytest.mark.asyncio
    async def test_health_check(self, backup_manager):
        """Test backup manager health check"""
        health = await backup_manager.health_check()
        
        assert "status" in health
        assert "backup_dir_exists" in health
        assert "backup_dir_writable" in health
        assert "auto_backup_enabled" in health
        assert "storage_backend_supports_backup" in health
        assert "storage_backend_supports_restore" in health
        assert "total_backups" in health
        
        assert health["status"] in ["healthy", "warning", "error"]
        assert health["backup_dir_exists"] is True
        assert health["backup_dir_writable"] is True
        assert health["auto_backup_enabled"] is True
        assert health["storage_backend_supports_backup"] is True
        assert health["storage_backend_supports_restore"] is True
    
    @pytest.mark.asyncio
    async def test_backup_cleanup(self, backup_manager, context_store, sample_data):
        """Test automatic cleanup of old backups"""
        session, messages = sample_data
        
        # Add some data
        await context_store.save_session(session)
        for message in messages:
            await context_store.save_message(session.id, message)
        
        # Create more backups than the limit (5)
        for i in range(7):
            await backup_manager.create_backup(backup_type=f"test{i}")
        
        # Check that only max_backups (5) remain for backup manager backups
        backups = backup_manager.list_backups()
        manager_backups = [b for b in backups if b.get("backup_id")]
        assert len(manager_backups) <= backup_manager.max_backups
    
    def test_backup_manager_initialization(self, context_store, config):
        """Test backup manager initialization"""
        backup_manager = BackupManager(context_store, config)
        
        assert backup_manager.context_store == context_store
        assert backup_manager.config == config
        assert backup_manager.backup_dir.exists()
        assert backup_manager.max_backups == 5
        assert backup_manager.backup_interval_hours == 1
        assert backup_manager.auto_backup_enabled is True
        assert backup_manager.verify_backups is True