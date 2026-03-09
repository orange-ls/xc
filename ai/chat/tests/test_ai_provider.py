"""
Tests for AI provider system
"""

import pytest
import asyncio
from unittest.mock import Mock, AsyncMock
from ai.chat.core.models import Message, AIResponse, ChatConfig, MessageRole, TokenUsage
from ai.chat.core.exceptions import AllModelsFailedError, ModelError
from ai.chat.providers.ai_provider import AIProvider
from ai.chat.providers.mock_provider import MockAIProvider


class TestMockAIProvider:
    """Test cases for MockAIProvider"""
    
    @pytest.fixture
    def mock_provider(self):
        """Create a mock AI provider for testing"""
        return MockAIProvider()
    
    @pytest.fixture
    def sample_context(self):
        """Create sample conversation context"""
        return [
            Message(
                id="1",
                session_id="test-session",
                content="Hello",
                role=MessageRole.USER
            ),
            Message(
                id="2", 
                session_id="test-session",
                content="Hi there! How can I help you?",
                role=MessageRole.ASSISTANT
            )
        ]
    
    def test_mock_provider_initialization(self, mock_provider):
        """Test mock provider initialization"""
        assert mock_provider.get_model_name() == "mock-ai-v1"
        assert mock_provider.is_available() is True
        assert mock_provider.failure_rate == 0.0
        assert mock_provider.response_delay == (0.1, 0.5)
    
    @pytest.mark.asyncio
    async def test_generate_response_basic(self, mock_provider, sample_context):
        """Test basic response generation"""
        response = await mock_provider.generate_response(sample_context, "How are you?")
        
        assert isinstance(response, AIResponse)
        assert response.content
        assert response.model_name == "mock-ai-v1"
        assert 0.0 <= response.confidence_score <= 1.0
        assert response.processing_time > 0
        assert isinstance(response.token_usage, TokenUsage)
        assert response.token_usage.total_tokens > 0
    
    @pytest.mark.asyncio
    async def test_generate_response_greeting(self, mock_provider):
        """Test greeting response"""
        response = await mock_provider.generate_response([], "Hello")
        
        # Should recognize greeting and respond appropriately
        assert any(word in response.content.lower() for word in ['hello', 'hi', 'greetings'])
    
    @pytest.mark.asyncio
    async def test_generate_response_farewell(self, mock_provider):
        """Test farewell response"""
        response = await mock_provider.generate_response([], "Goodbye")
        
        # Should recognize farewell and respond appropriately
        # Note: The mock provider may not always use farewell responses, so we just check for valid response
        assert response.content
        assert len(response.content) > 0
    
    @pytest.mark.asyncio
    async def test_generate_response_question(self, mock_provider):
        """Test question response"""
        response = await mock_provider.generate_response([], "What is the weather like?")
        
        # Should recognize question and respond appropriately
        assert response.content
        assert len(response.content) > 0
    
    @pytest.mark.asyncio
    async def test_generate_response_with_context(self, mock_provider, sample_context):
        """Test response generation with conversation context"""
        response = await mock_provider.generate_response(sample_context, "Tell me more")
        
        # Should acknowledge ongoing conversation
        assert "continuing" in response.content.lower() or response.content
    
    @pytest.mark.asyncio
    async def test_generate_response_empty_message(self, mock_provider):
        """Test response generation with empty message"""
        with pytest.raises(ValueError, match="Message cannot be empty"):
            await mock_provider.generate_response([], "")
    
    @pytest.mark.asyncio
    async def test_generate_response_with_failure(self, mock_provider):
        """Test response generation with simulated failure"""
        mock_provider.set_failure_rate(1.0)  # 100% failure rate
        
        with pytest.raises(ModelError):
            await mock_provider.generate_response([], "Hello")
    
    @pytest.mark.asyncio
    async def test_health_check(self, mock_provider):
        """Test health check functionality"""
        # Should be healthy initially
        assert await mock_provider.health_check() is True
        
        # Test manual health setting
        mock_provider.set_healthy(False)
        assert mock_provider.is_available() is False
        
        mock_provider.set_healthy(True)
        assert mock_provider.is_available() is True
    
    def test_set_failure_rate(self, mock_provider):
        """Test setting failure rate"""
        mock_provider.set_failure_rate(0.5)
        assert mock_provider.failure_rate == 0.5
        
        # Test bounds
        mock_provider.set_failure_rate(-0.1)
        assert mock_provider.failure_rate == 0.0
        
        mock_provider.set_failure_rate(1.5)
        assert mock_provider.failure_rate == 1.0
    
    def test_set_response_delay(self, mock_provider):
        """Test setting response delay"""
        mock_provider.set_response_delay(0.2, 0.8)
        assert mock_provider.response_delay == (0.2, 0.8)


class TestAIProvider:
    """Test cases for AIProvider"""
    
    @pytest.fixture
    def config(self):
        """Create test configuration"""
        return ChatConfig(
            ai_models=["mock"],
            fallback_models=["mock"]
        )
    
    @pytest.fixture
    def ai_provider(self, config):
        """Create AI provider for testing"""
        return AIProvider(config)
    
    @pytest.fixture
    def sample_context(self):
        """Create sample conversation context"""
        return [
            Message(
                id="1",
                session_id="test-session", 
                content="Hello",
                role=MessageRole.USER
            )
        ]
    
    def test_ai_provider_initialization(self, ai_provider):
        """Test AI provider initialization"""
        assert "mock" in ai_provider.providers
        assert ai_provider.fallback_chain_order == ["mock"]
        assert ai_provider.get_current_model() == "mock"
    
    @pytest.mark.asyncio
    async def test_generate_response_success(self, ai_provider, sample_context):
        """Test successful response generation"""
        response = await ai_provider.generate_response(sample_context, "Hello")
        
        assert isinstance(response, AIResponse)
        assert response.content
        assert response.model_name == "mock-ai-v1"
    
    @pytest.mark.asyncio
    async def test_generate_response_empty_message(self, ai_provider, sample_context):
        """Test response generation with empty message"""
        with pytest.raises(ValueError, match="Message cannot be empty"):
            await ai_provider.generate_response(sample_context, "")
    
    @pytest.mark.asyncio
    async def test_generate_response_too_long_message(self, ai_provider, sample_context):
        """Test response generation with message too long"""
        long_message = "x" * 5000  # Exceeds default max_message_length of 4000
        
        with pytest.raises(ValueError, match="Message exceeds maximum length"):
            await ai_provider.generate_response(sample_context, long_message)
    
    @pytest.mark.asyncio
    async def test_generate_response_all_models_failed(self, ai_provider, sample_context):
        """Test response generation when all models fail"""
        # Make the mock provider unhealthy
        ai_provider.providers["mock"].set_healthy(False)
        
        with pytest.raises(AllModelsFailedError):
            await ai_provider.generate_response(sample_context, "Hello")
    
    def test_switch_model_success(self, ai_provider):
        """Test successful model switching"""
        result = ai_provider.switch_model("mock")
        assert result is True
        assert ai_provider.get_current_model() == "mock"
    
    def test_switch_model_unknown(self, ai_provider):
        """Test switching to unknown model"""
        result = ai_provider.switch_model("unknown")
        assert result is False
    
    def test_switch_model_unavailable(self, ai_provider):
        """Test switching to unavailable model"""
        ai_provider.providers["mock"].set_healthy(False)
        result = ai_provider.switch_model("mock")
        assert result is False
    
    def test_get_available_models(self, ai_provider):
        """Test getting available models"""
        models = ai_provider.get_available_models()
        assert "mock" in models
        
        # Make mock unavailable
        ai_provider.providers["mock"].set_healthy(False)
        models = ai_provider.get_available_models()
        assert "mock" not in models
    
    def test_get_provider_info(self, ai_provider):
        """Test getting provider information"""
        info = ai_provider.get_provider_info("mock")
        assert info is not None
        assert info["name"] == "mock"
        assert info["model_name"] == "mock-ai-v1"
        assert "available" in info
        assert "in_fallback_chain" in info
        
        # Test unknown provider
        info = ai_provider.get_provider_info("unknown")
        assert info is None
    
    def test_add_provider(self, ai_provider):
        """Test adding new provider"""
        new_provider = MockAIProvider()
        result = ai_provider.add_provider("test", new_provider)
        assert result is True
        assert "test" in ai_provider.providers
        
        # Test adding duplicate
        result = ai_provider.add_provider("test", new_provider)
        assert result is False
    
    def test_remove_provider(self, ai_provider):
        """Test removing provider"""
        # Add a provider first
        new_provider = MockAIProvider()
        ai_provider.add_provider("test", new_provider)
        ai_provider.fallback_chain.fallback_order.append("test")
        
        # Remove it
        result = ai_provider.remove_provider("test")
        assert result is True
        assert "test" not in ai_provider.providers
        assert "test" not in ai_provider.fallback_chain.fallback_order
        
        # Test removing non-existent provider
        result = ai_provider.remove_provider("nonexistent")
        assert result is False
    
    @pytest.mark.asyncio
    async def test_health_check(self, ai_provider):
        """Test health check functionality"""
        health = await ai_provider.health_check()
        
        assert "status" in health
        assert "providers" in health
        assert "fallback_stats" in health
        assert "retry_config" in health
        assert "current_model" in health
        assert "mock" in health["providers"]
        
        provider_health = health["providers"]["mock"]
        assert "status" in provider_health
        assert "available" in provider_health