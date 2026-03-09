"""
会话管理器的单元测试
"""

import pytest
import asyncio
from datetime import datetime, timedelta
from ..managers.session_manager import SessionManager
from ..core.models import ChatConfig, Session, SessionStatus
from ..core.exceptions import SessionError


class TestSessionManager:
    """测试SessionManager功能"""
    
    @pytest.fixture
    def config(self):
        """创建测试配置"""
        return ChatConfig(
            max_concurrent_sessions=3,  # 较小的限制用于测试
            session_timeout_hours=1
        )
    
    @pytest.fixture
    def session_manager(self, config):
        """创建SessionManager实例用于测试"""
        return SessionManager(config)
    
    def test_create_session(self, session_manager):
        """测试创建会话"""
        session_id = session_manager.create_session()
        
        assert session_id is not None
        assert isinstance(session_id, str)
        assert len(session_id) > 0
        assert session_id in session_manager.sessions
    
    def test_create_multiple_sessions(self, session_manager):
        """测试创建多个会话"""
        session_ids = []
        for i in range(3):  # 创建最大数量的会话
            session_id = session_manager.create_session()
            session_ids.append(session_id)
        
        # 所有会话ID应该是唯一的
        assert len(set(session_ids)) == 3
        assert len(session_manager.sessions) == 3
    
    def test_create_session_exceeds_limit(self, session_manager):
        """测试超过最大会话数限制"""
        # 创建最大数量的会话
        for i in range(3):
            session_manager.create_session()
        
        # 尝试创建第4个会话应该失败
        with pytest.raises(SessionError, match="Maximum number of concurrent sessions exceeded"):
            session_manager.create_session()
    
    @pytest.mark.asyncio
    async def test_get_session(self, session_manager):
        """测试获取会话"""
        session_id = session_manager.create_session()
        
        session = await session_manager.get_session(session_id)
        
        assert session is not None
        assert session.id == session_id
        assert session.status == SessionStatus.ACTIVE
    
    @pytest.mark.asyncio
    async def test_get_nonexistent_session(self, session_manager):
        """测试获取不存在的会话"""
        session = await session_manager.get_session("nonexistent-session")
        
        assert session is None
    
    @pytest.mark.asyncio
    async def test_update_session(self, session_manager):
        """测试更新会话"""
        session_id = session_manager.create_session()
        session = await session_manager.get_session(session_id)
        
        # 更新会话信息
        session.message_count = 5
        session.metadata["test"] = "value"
        await session_manager.update_session(session)
        
        # 验证更新
        updated_session = await session_manager.get_session(session_id)
        assert updated_session.message_count == 5
        assert updated_session.metadata["test"] == "value"
    
    @pytest.mark.asyncio
    async def test_end_session(self, session_manager):
        """测试结束会话"""
        session_id = session_manager.create_session()
        
        result = await session_manager.end_session(session_id)
        
        assert result is True
        session = await session_manager.get_session(session_id)
        assert session.status == SessionStatus.ENDED
    
    @pytest.mark.asyncio
    async def test_end_nonexistent_session(self, session_manager):
        """测试结束不存在的会话"""
        result = await session_manager.end_session("nonexistent-session")
        
        assert result is False
    
    @pytest.mark.asyncio
    async def test_delete_session(self, session_manager):
        """测试删除会话"""
        session_id = session_manager.create_session()
        
        result = await session_manager.delete_session(session_id)
        
        assert result is True
        assert session_id not in session_manager.sessions
    
    @pytest.mark.asyncio
    async def test_delete_nonexistent_session(self, session_manager):
        """测试删除不存在的会话"""
        result = await session_manager.delete_session("nonexistent-session")
        
        assert result is False
    
    @pytest.mark.asyncio
    async def test_list_active_sessions(self, session_manager):
        """测试列出活跃会话"""
        # 创建一些会话
        session_id1 = session_manager.create_session()
        session_id2 = session_manager.create_session()
        session_id3 = session_manager.create_session()
        
        # 结束一个会话
        await session_manager.end_session(session_id3)
        
        active_sessions = await session_manager.list_active_sessions()
        
        assert len(active_sessions) == 2
        assert session_id1 in active_sessions
        assert session_id2 in active_sessions
        assert session_id3 not in active_sessions
    
    @pytest.mark.asyncio
    async def test_cleanup_expired_sessions(self, session_manager):
        """测试清理过期会话"""
        # 创建会话
        session_id = session_manager.create_session()
        session = await session_manager.get_session(session_id)
        
        # 手动设置过期时间
        session.last_activity = datetime.now() - timedelta(hours=2)
        await session_manager.update_session(session)
        
        # 清理过期会话
        cleaned_count = await session_manager.cleanup_expired_sessions()
        
        assert cleaned_count == 1
        assert session_id not in session_manager.sessions
    
    def test_create_session_after_cleanup(self, session_manager):
        """测试清理后可以创建新会话"""
        # 填满会话池
        session_ids = []
        for i in range(3):
            session_id = session_manager.create_session()
            session_ids.append(session_id)
        
        # 手动设置所有会话为过期
        for session_id in session_ids:
            session = session_manager.sessions[session_id]
            session.last_activity = datetime.now() - timedelta(hours=2)
        
        # 现在应该可以创建新会话（会自动清理过期会话）
        new_session_id = session_manager.create_session()
        
        assert new_session_id is not None
        assert len(session_manager.sessions) == 1  # 只有新会话
    
    @pytest.mark.asyncio
    async def test_health_check(self, session_manager):
        """测试健康检查"""
        # 创建一些会话
        session_manager.create_session()
        session_manager.create_session()
        
        health_status = await session_manager.health_check()
        
        assert health_status["status"] == "healthy"
        assert health_status["active_sessions"] == 2
        assert health_status["total_sessions"] == 2
        assert health_status["max_sessions"] == 3
        assert health_status["utilization"] == 2/3
    
    def test_session_unique_ids(self, session_manager):
        """测试会话ID的唯一性"""
        session_ids = set()
        
        # 创建多个会话
        for i in range(10):
            session_id = session_manager.create_session()
            session_ids.add(session_id)
            # 删除会话以避免达到限制
            session_manager.sessions.pop(session_id)
        
        # 所有ID应该是唯一的
        assert len(session_ids) == 10
    
    @pytest.mark.asyncio
    async def test_session_activity_update(self, session_manager):
        """测试会话活动时间更新"""
        session_id = session_manager.create_session()
        session = await session_manager.get_session(session_id)
        
        original_activity = session.last_activity
        
        # 等待一小段时间
        await asyncio.sleep(0.01)
        
        # 更新活动时间
        session.update_activity()
        await session_manager.update_session(session)
        
        updated_session = await session_manager.get_session(session_id)
        assert updated_session.last_activity > original_activity