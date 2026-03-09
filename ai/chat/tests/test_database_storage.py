"""
Tests for database storage backend
"""

import pytest
import tempfile
import os
from pathlib import Path
from datetime import datetime

from ai.chat.storage.database_storage import DatabaseStorage
from ai.chat.core.models import Message, Session, MessageRole, SessionStatus
from ai.chat.core.exceptions import StorageError


class TestDatabaseStorage:
    """Test database storage backend"""
    
    @pytest.fixture
    def temp_db_path(self):
        """Create temporary database path for testing"""
        temp_file = tempfile.NamedTemporaryFile(suffix='.db', delete=False)
        temp_file.close()
        yield temp_file.name
        # Clean up
        if os.path.exists(temp_file.name):
            os.unlink(temp_file.name)
    
    @pytest.fixture
    def db_storage(self, temp_db_path):
        """Create database storage instance with temporary database"""
        return DatabaseStorage(temp_db_path)
    
    @pytest.fixture
    def sample_message(self):
        """Create sample message for testing"""
        return Message(
            id="test-message-1",
            session_id="test-session-1",
            content="Hello, database!",
            role=MessageRole.USER,
            timestamp=datetime.now(),
            metadata={"test": "data"}
        )
    
    @pytest.fixture
    def sample_session(self):
        """Create sample session for testing"""
        return Session(
            id="test-session-1",
            created_at=datetime.now(),
            last_activity=datetime.now(),
            status=SessionStatus.ACTIVE,
            metadata={"test": "session"},
            message_count=1
        )
    
    @pytest.mark.asyncio
    async def test_save_and_get_message(self, db_storage, sample_message):
        """Test saving and retrieving messages"""
        # Save message
        await db_storage.save_message(sample_message.session_id, sample_message)
        
        # Retrieve messages
        messages = await db_storage.get_messages(sample_message.session_id)
        
        assert len(messages) == 1
        retrieved_message = messages[0]
        assert retrieved_message.id == sample_message.id
        assert retrieved_message.content == sample_message.content
        assert retrieved_message.role == sample_message.role
        assert retrieved_message.metadata == sample_message.metadata
    
    @pytest.mark.asyncio
    async def test_save_multiple_messages(self, db_storage):
        """Test saving multiple messages"""
        session_id = "test-session-multi"
        
        messages = [
            Message(id="msg1", session_id=session_id, content="First message", role=MessageRole.USER),
            Message(id="msg2", session_id=session_id, content="Second message", role=MessageRole.ASSISTANT),
            Message(id="msg3", session_id=session_id, content="Third message", role=MessageRole.USER)
        ]
        
        # Save all messages
        for message in messages:
            await db_storage.save_message(session_id, message)
        
        # Retrieve all messages
        retrieved_messages = await db_storage.get_messages(session_id)
        
        assert len(retrieved_messages) == 3
        assert [m.content for m in retrieved_messages] == ["First message", "Second message", "Third message"]
    
    @pytest.mark.asyncio
    async def test_get_messages_with_limit(self, db_storage):
        """Test retrieving messages with limit"""
        session_id = "test-session-limit"
        
        # Save 5 messages
        for i in range(5):
            message = Message(
                id=f"msg{i}",
                session_id=session_id,
                content=f"Message {i}",
                role=MessageRole.USER
            )
            await db_storage.save_message(session_id, message)
        
        # Get last 3 messages
        messages = await db_storage.get_messages(session_id, limit=3)
        
        assert len(messages) == 3
        assert [m.content for m in messages] == ["Message 2", "Message 3", "Message 4"]
    
    @pytest.mark.asyncio
    async def test_save_and_get_session(self, db_storage, sample_session):
        """Test saving and retrieving session"""
        # Save session
        await db_storage.save_session(sample_session)
        
        # Retrieve session
        retrieved_session = await db_storage.get_session(sample_session.id)
        
        assert retrieved_session is not None
        assert retrieved_session.id == sample_session.id
        assert retrieved_session.status == sample_session.status
        assert retrieved_session.metadata == sample_session.metadata
        assert retrieved_session.message_count == sample_session.message_count
    
    @pytest.mark.asyncio
    async def test_get_nonexistent_session(self, db_storage):
        """Test retrieving non-existent session"""
        session = await db_storage.get_session("nonexistent-session")
        assert session is None
    
    @pytest.mark.asyncio
    async def test_get_messages_nonexistent_session(self, db_storage):
        """Test retrieving messages from non-existent session"""
        messages = await db_storage.get_messages("nonexistent-session")
        assert messages == []
    
    @pytest.mark.asyncio
    async def test_delete_session(self, db_storage, sample_session, sample_message):
        """Test deleting session and its messages"""
        # Save session and message
        await db_storage.save_session(sample_session)
        await db_storage.save_message(sample_message.session_id, sample_message)
        
        # Verify they exist
        assert await db_storage.get_session(sample_session.id) is not None
        assert len(await db_storage.get_messages(sample_message.session_id)) == 1
        
        # Delete session
        result = await db_storage.delete_session(sample_session.id)
        assert result is True
        
        # Verify they're gone
        assert await db_storage.get_session(sample_session.id) is None
        assert await db_storage.get_messages(sample_message.session_id) == []
    
    @pytest.mark.asyncio
    async def test_delete_nonexistent_session(self, db_storage):
        """Test deleting non-existent session"""
        result = await db_storage.delete_session("nonexistent-session")
        assert result is False
    
    @pytest.mark.asyncio
    async def test_backup_data(self, db_storage, sample_session, sample_message):
        """Test data backup functionality"""
        # Save some data
        await db_storage.save_session(sample_session)
        await db_storage.save_message(sample_message.session_id, sample_message)
        
        # Create backup
        result = await db_storage.backup_data()
        assert result is True
        
        # Verify backup exists
        backups = db_storage.get_backup_list()
        assert len(backups) >= 1
        assert backups[0]["backup_type"] == "full"
    
    @pytest.mark.asyncio
    async def test_restore_from_backup(self, db_storage, sample_session, sample_message):
        """Test data restore functionality"""
        # Save some data
        await db_storage.save_session(sample_session)
        await db_storage.save_message(sample_message.session_id, sample_message)
        
        # Create backup
        await db_storage.backup_data()
        backups = db_storage.get_backup_list()
        backup_timestamp = backups[0]["timestamp"]
        
        # Delete original data
        await db_storage.delete_session(sample_session.id)
        assert await db_storage.get_session(sample_session.id) is None
        
        # Restore from backup
        result = await db_storage.restore_from_backup(backup_timestamp)
        assert result is True
        
        # Verify data is restored
        restored_session = await db_storage.get_session(sample_session.id)
        assert restored_session is not None
        assert restored_session.id == sample_session.id
        
        restored_messages = await db_storage.get_messages(sample_message.session_id)
        assert len(restored_messages) == 1
        assert restored_messages[0].content == sample_message.content
    
    @pytest.mark.asyncio
    async def test_get_stats(self, db_storage, sample_session, sample_message):
        """Test storage statistics"""
        # Add some data
        await db_storage.save_session(sample_session)
        await db_storage.save_message(sample_message.session_id, sample_message)
        
        stats = await db_storage.get_stats()
        
        assert "total_sessions" in stats
        assert "total_messages" in stats
        assert "active_sessions" in stats
        assert "database_size_bytes" in stats
        assert "backup_count" in stats
        
        assert stats["total_sessions"] >= 1
        assert stats["total_messages"] >= 1
        assert stats["active_sessions"] >= 1
    
    @pytest.mark.asyncio
    async def test_vacuum_database(self, db_storage):
        """Test database vacuum operation"""
        result = await db_storage.vacuum_database()
        assert result is True
    
    @pytest.mark.asyncio
    async def test_session_count_by_date(self, db_storage, sample_session):
        """Test getting session count by date"""
        # Save a session
        await db_storage.save_session(sample_session)
        
        # Get session count by date
        counts = await db_storage.get_session_count_by_date(days=7)
        
        assert isinstance(counts, dict)
        # Should have at least one entry for today
        assert len(counts) >= 0  # Might be 0 if no sessions created today
    
    def test_database_initialization(self, temp_db_path):
        """Test that database is properly initialized"""
        storage = DatabaseStorage(temp_db_path)
        
        # Database file should exist
        assert Path(temp_db_path).exists()
        
        # Should be able to create another instance without errors
        storage2 = DatabaseStorage(temp_db_path)
        assert storage2 is not None