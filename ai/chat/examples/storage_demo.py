#!/usr/bin/env python3
"""
Storage System Demo

This example demonstrates the AI Chat Module's storage system capabilities:
- File storage backend
- Database storage backend  
- Backup and recovery functionality
- Context store operations
"""

import asyncio
import tempfile
import shutil
from datetime import datetime
from pathlib import Path

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.models import ChatConfig, Message, Session, MessageRole, SessionStatus
from storage import ContextStore


async def demo_file_storage():
    """Demonstrate file storage backend"""
    print("=== File Storage Demo ===")
    
    # Create temporary directory for demo
    temp_dir = tempfile.mkdtemp()
    print(f"Using temporary directory: {temp_dir}")
    
    try:
        # Configure for file storage
        config = ChatConfig(
            storage_backend="file",
            data_dir=temp_dir,
            max_backups=3,
            auto_backup_enabled=True
        )
        
        # Create context store
        context_store = ContextStore(config)
        
        # Create a session
        session = Session(
            id="demo-session-1",
            status=SessionStatus.ACTIVE,
            metadata={"demo": "file_storage"}
        )
        
        await context_store.save_session(session)
        print(f"✓ Created session: {session.id}")
        
        # Add some messages
        messages = [
            Message(
                id="msg1",
                session_id=session.id,
                content="Hello, file storage!",
                role=MessageRole.USER,
                metadata={"demo": "message"}
            ),
            Message(
                id="msg2", 
                session_id=session.id,
                content="File storage is working great!",
                role=MessageRole.ASSISTANT,
                metadata={"demo": "response"}
            )
        ]
        
        for message in messages:
            await context_store.save_message(session.id, message)
        
        print(f"✓ Added {len(messages)} messages")
        
        # Retrieve messages
        retrieved_messages = await context_store.get_messages(session.id)
        print(f"✓ Retrieved {len(retrieved_messages)} messages")
        
        for msg in retrieved_messages:
            print(f"  - {msg.role.value}: {msg.content}")
        
        # Create backup
        backup_info = await context_store.create_backup("demo")
        print(f"✓ Created backup: {backup_info['backup_id']}")
        
        # List backups
        backups = context_store.get_backup_list()
        print(f"✓ Found {len(backups)} backups")
        
        # Health check
        health = await context_store.health_check()
        print(f"✓ Storage health: {health['status']}")
        
    finally:
        # Cleanup
        shutil.rmtree(temp_dir)
        print(f"✓ Cleaned up temporary directory")


async def demo_database_storage():
    """Demonstrate database storage backend"""
    print("\n=== Database Storage Demo ===")
    
    # Create temporary directory for demo
    temp_dir = tempfile.mkdtemp()
    print(f"Using temporary directory: {temp_dir}")
    
    try:
        # Configure for database storage
        config = ChatConfig(
            storage_backend="database",
            data_dir=temp_dir,
            max_backups=3,
            auto_backup_enabled=True
        )
        
        # Create context store
        context_store = ContextStore(config)
        
        # Create a session
        session = Session(
            id="demo-session-2",
            status=SessionStatus.ACTIVE,
            metadata={"demo": "database_storage"}
        )
        
        await context_store.save_session(session)
        print(f"✓ Created session: {session.id}")
        
        # Add some messages
        messages = [
            Message(
                id="msg1",
                session_id=session.id,
                content="Hello, database storage!",
                role=MessageRole.USER,
                metadata={"demo": "message"}
            ),
            Message(
                id="msg2",
                session_id=session.id,
                content="Database storage is efficient!",
                role=MessageRole.ASSISTANT,
                metadata={"demo": "response"}
            ),
            Message(
                id="msg3",
                session_id=session.id,
                content="Can you tell me more?",
                role=MessageRole.USER,
                metadata={"demo": "followup"}
            )
        ]
        
        for message in messages:
            await context_store.save_message(session.id, message)
        
        print(f"✓ Added {len(messages)} messages")
        
        # Test message limit
        limited_messages = await context_store.get_messages(session.id, limit=2)
        print(f"✓ Retrieved last {len(limited_messages)} messages (with limit)")
        
        # Create backup
        backup_info = await context_store.create_backup("database_demo")
        print(f"✓ Created backup: {backup_info['backup_id']}")
        
        # Delete session
        await context_store.delete_session(session.id)
        print(f"✓ Deleted session: {session.id}")
        
        # Verify deletion
        deleted_session = await context_store.get_session(session.id)
        print(f"✓ Session deleted: {deleted_session is None}")
        
        # Restore from backup
        success = await context_store.restore_backup(backup_info['backup_id'])
        print(f"✓ Restored from backup: {success}")
        
        # Verify restoration
        restored_session = await context_store.get_session(session.id)
        restored_messages = await context_store.get_messages(session.id)
        print(f"✓ Session restored: {restored_session is not None}")
        print(f"✓ Messages restored: {len(restored_messages)}")
        
    finally:
        # Cleanup
        shutil.rmtree(temp_dir)
        print(f"✓ Cleaned up temporary directory")


async def demo_backup_features():
    """Demonstrate backup and recovery features"""
    print("\n=== Backup & Recovery Demo ===")
    
    # Create temporary directory for demo
    temp_dir = tempfile.mkdtemp()
    print(f"Using temporary directory: {temp_dir}")
    
    try:
        # Configure with backup settings
        config = ChatConfig(
            storage_backend="file",
            data_dir=temp_dir,
            max_backups=2,  # Keep only 2 backups
            backup_interval_hours=0,  # Allow immediate backups
            auto_backup_enabled=True,
            verify_backups=True
        )
        
        context_store = ContextStore(config)
        
        # Create test data
        session = Session(id="backup-demo", status=SessionStatus.ACTIVE)
        await context_store.save_session(session)
        
        message = Message(
            session_id=session.id,
            content="This is a backup test message",
            role=MessageRole.USER
        )
        await context_store.save_message(session.id, message)
        
        print("✓ Created test data")
        
        # Create multiple backups to test cleanup
        backup_ids = []
        for i in range(4):
            backup_info = await context_store.create_backup(f"test_backup_{i}")
            backup_ids.append(backup_info['backup_id'])
            print(f"✓ Created backup {i+1}: {backup_info['backup_id']}")
        
        # Check backup cleanup (should keep only max_backups)
        backups = context_store.get_backup_list()
        manager_backups = [b for b in backups if b.get("backup_id")]
        print(f"✓ Backup cleanup: {len(manager_backups)} backups remaining (max: {config.max_backups})")
        
        # Get backup statistics
        backup_manager = context_store.backup_manager
        stats = await backup_manager.get_backup_stats()
        print(f"✓ Backup stats:")
        print(f"  - Total backups: {stats['total_backups']}")
        print(f"  - Total size: {stats['total_size_mb']} MB")
        print(f"  - Auto backup enabled: {stats['auto_backup_enabled']}")
        
        # Verify backup integrity
        verification_results = await backup_manager.verify_all_backups()
        verified_count = sum(1 for result in verification_results.values() if result)
        print(f"✓ Backup verification: {verified_count}/{len(verification_results)} backups verified")
        
        # Test scheduled backup
        await backup_manager.schedule_auto_backup()
        print("✓ Scheduled auto backup completed")
        
        # Health check
        health = await backup_manager.health_check()
        print(f"✓ Backup system health: {health['status']}")
        
    finally:
        # Cleanup
        shutil.rmtree(temp_dir)
        print(f"✓ Cleaned up temporary directory")


async def main():
    """Run all storage demos"""
    print("AI Chat Module - Storage System Demo")
    print("=" * 50)
    
    await demo_file_storage()
    await demo_database_storage()
    await demo_backup_features()
    
    print("\n" + "=" * 50)
    print("Storage system demo completed successfully!")
    print("\nKey features demonstrated:")
    print("• File storage backend with JSON serialization")
    print("• Database storage backend with SQLite")
    print("• Atomic write operations for data integrity")
    print("• Comprehensive backup and recovery system")
    print("• Automatic backup cleanup and scheduling")
    print("• Backup verification and integrity checking")
    print("• Health monitoring for storage components")


if __name__ == "__main__":
    asyncio.run(main())