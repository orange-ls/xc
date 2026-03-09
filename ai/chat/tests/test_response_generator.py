"""
Tests for ResponseGenerator
"""

import pytest
from unittest.mock import Mock, AsyncMock, patch
from ai.chat.core.models import Message, AIResponse, ChatResponse, ChatConfig, MessageRole, TokenUsage
from ai.chat.core.exceptions import AllModelsFailedError, ModelError
from ai.chat.generators.response_generator import ResponseGenerator
from ai.chat.providers.ai_provider import AIProvider


class TestResponseGenerator:
    """Test cases for ResponseGenerator"""
    
    @pytest.fixture
    def config(self):
        """Create test configuration"""
        return ChatConfig(
            max_message_length=4000,
            ai_models=["mock"],
            fallback_models=["mock"]
        )
    
    @pytest.fixture
    def mock_ai_provider(self):
        """Create mock AI provider"""
        provider = Mock(spec=AIProvider)
        provider.generate_response = AsyncMock()
        provider.switch_model = Mock(return_value=True)
        provider.get_available_models = Mock(return_value=["mock"])
        provider.get_current_model = Mock(return_value="mock")
        provider.health_check = AsyncMock(return_value={"status": "healthy"})
        return provider
    
    @pytest.fixture
    def response_generator(self, config, mock_ai_provider):
        """Create response generator for testing"""
        return ResponseGenerator(config, mock_ai_provider)
    
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
    
    @pytest.fixture
    def sample_ai_response(self):
        """Create sample AI response"""
        return AIResponse(
            content="This is a test response.",
            model_name="mock-ai-v1",
            confidence_score=0.9,
            processing_time=0.5,
            token_usage=TokenUsage(prompt_tokens=10, completion_tokens=5, total_tokens=15)
        )
    
    def test_response_generator_initialization(self, response_generator, config):
        """Test response generator initialization"""
        assert response_generator.config == config
        assert response_generator.max_response_length == 2000
        assert response_generator.min_confidence_threshold == 0.1
        assert len(response_generator.fallback_responses) > 0
    
    @pytest.mark.asyncio
    async def test_generate_response_success(self, response_generator, mock_ai_provider, sample_context, sample_ai_response):
        """Test successful response generation"""
        mock_ai_provider.generate_response.return_value = sample_ai_response
        
        response = await response_generator.generate_response(sample_context, "How are you?", "test-session")
        
        assert isinstance(response, ChatResponse)
        assert response.message == "This is a test response."
        assert response.session_id == "test-session"
        assert response.model_used == "mock-ai-v1"
        assert response.confidence == 0.9
        assert response.response_time > 0
        assert "token_usage" in response.metadata
        
        mock_ai_provider.generate_response.assert_called_once_with(sample_context, "How are you?")
    
    @pytest.mark.asyncio
    async def test_generate_response_empty_message(self, response_generator, sample_context):
        """Test response generation with empty message"""
        with pytest.raises(ValueError, match="Message cannot be empty"):
            await response_generator.generate_response(sample_context, "", "test-session")
    
    @pytest.mark.asyncio
    async def test_generate_response_all_models_failed(self, response_generator, mock_ai_provider, sample_context):
        """Test response generation when all models fail"""
        mock_ai_provider.generate_response.side_effect = AllModelsFailedError("All models failed")
        
        response = await response_generator.generate_response(sample_context, "Hello", "test-session")
        
        assert isinstance(response, ChatResponse)
        assert "unable to process" in response.message.lower()
        assert response.model_used == "error_handler"
        assert response.confidence == 0.0
        assert response.metadata["is_error_response"] is True
        assert response.metadata["error_type"] == "service_unavailable"
    
    @pytest.mark.asyncio
    async def test_generate_response_general_error(self, response_generator, mock_ai_provider, sample_context):
        """Test response generation with general error"""
        mock_ai_provider.generate_response.side_effect = Exception("Test error")
        
        response = await response_generator.generate_response(sample_context, "Hello", "test-session")
        
        assert isinstance(response, ChatResponse)
        assert "encountered an error" in response.message.lower()
        assert response.model_used == "error_handler"
        assert response.confidence == 0.0
        assert response.metadata["error_details"] == "Test error"
    
    def test_format_ai_response_normal(self, response_generator, sample_ai_response):
        """Test formatting normal AI response"""
        formatted = response_generator._format_ai_response(sample_ai_response)
        
        assert formatted.content == "This is a test response."
        assert formatted.confidence_score == 0.9
    
    def test_format_ai_response_empty_content(self, response_generator):
        """Test formatting AI response with empty content"""
        empty_response = AIResponse(
            content="",
            model_name="test",
            confidence_score=0.8
        )
        
        formatted = response_generator._format_ai_response(empty_response)
        
        assert formatted.content in response_generator.fallback_responses
        assert formatted.confidence_score == response_generator.min_confidence_threshold
    
    def test_format_ai_response_too_long(self, response_generator):
        """Test formatting AI response that's too long"""
        long_content = "x" * 3000  # Longer than max_response_length (2000)
        long_response = AIResponse(
            content=long_content,
            model_name="test",
            confidence_score=0.8
        )
        
        formatted = response_generator._format_ai_response(long_response)
        
        assert len(formatted.content) <= response_generator.max_response_length
        assert formatted.content.endswith("...")
    
    def test_format_ai_response_confidence_bounds(self, response_generator):
        """Test confidence score bounds checking"""
        # Test lower bound
        low_response = AIResponse(content="test", model_name="test", confidence_score=-0.5)
        formatted_low = response_generator._format_ai_response(low_response)
        assert formatted_low.confidence_score == 0.0
        
        # Test upper bound
        high_response = AIResponse(content="test", model_name="test", confidence_score=1.5)
        formatted_high = response_generator._format_ai_response(high_response)
        assert formatted_high.confidence_score == 1.0
    
    def test_normalize_text(self, response_generator):
        """Test text normalization"""
        # Test whitespace normalization
        assert response_generator._normalize_text("  hello   world  ") == "hello world."
        
        # Test sentence ending
        assert response_generator._normalize_text("hello world") == "hello world."
        assert response_generator._normalize_text("hello world!") == "hello world!"
        assert response_generator._normalize_text("hello world?") == "hello world?"
    
    def test_get_fallback_response(self, response_generator):
        """Test fallback response selection"""
        response = response_generator._get_fallback_response()
        assert response in response_generator.fallback_responses
    
    def test_create_error_response(self, response_generator):
        """Test error response creation"""
        error_response = response_generator._create_error_response(
            "test-session", 1.5, "service_unavailable", "Test error"
        )
        
        assert isinstance(error_response, ChatResponse)
        assert error_response.session_id == "test-session"
        assert error_response.response_time == 1.5
        assert error_response.model_used == "error_handler"
        assert error_response.confidence == 0.0
        assert error_response.metadata["error_type"] == "service_unavailable"
        assert error_response.metadata["error_details"] == "Test error"
        assert error_response.metadata["is_error_response"] is True
    
    @pytest.mark.asyncio
    async def test_generate_ai_response(self, response_generator, mock_ai_provider, sample_context, sample_ai_response):
        """Test raw AI response generation"""
        mock_ai_provider.generate_response.return_value = sample_ai_response
        
        response = await response_generator.generate_ai_response(sample_context, "Hello")
        
        assert response == sample_ai_response
        mock_ai_provider.generate_response.assert_called_once_with(sample_context, "Hello")
    
    def test_switch_model(self, response_generator, mock_ai_provider):
        """Test model switching"""
        result = response_generator.switch_model("new-model")
        
        assert result is True
        mock_ai_provider.switch_model.assert_called_once_with("new-model")
    
    def test_get_available_models(self, response_generator, mock_ai_provider):
        """Test getting available models"""
        models = response_generator.get_available_models()
        
        assert models == ["mock"]
        mock_ai_provider.get_available_models.assert_called_once()
    
    def test_get_current_model(self, response_generator, mock_ai_provider):
        """Test getting current model"""
        model = response_generator.get_current_model()
        
        assert model == "mock"
        mock_ai_provider.get_current_model.assert_called_once()
    
    def test_set_fallback_responses(self, response_generator):
        """Test setting custom fallback responses"""
        new_responses = ["Response 1", "Response 2", "Response 3"]
        response_generator.set_fallback_responses(new_responses)
        
        assert response_generator.fallback_responses == new_responses
        
        # Test invalid responses
        response_generator.set_fallback_responses(["", None, "Valid"])
        assert response_generator.fallback_responses == new_responses  # Should remain unchanged
    
    def test_update_config(self, response_generator):
        """Test configuration updates"""
        response_generator.update_config(
            max_response_length=1500,
            min_confidence_threshold=0.2
        )
        
        assert response_generator.max_response_length == 1500
        assert response_generator.min_confidence_threshold == 0.2
    
    @pytest.mark.asyncio
    async def test_health_check_healthy(self, response_generator, mock_ai_provider, sample_ai_response):
        """Test health check when everything is healthy"""
        mock_ai_provider.health_check.return_value = {"status": "healthy"}
        mock_ai_provider.generate_response.return_value = sample_ai_response
        
        health = await response_generator.health_check()
        
        assert health["status"] == "healthy"
        assert "ai_provider" in health
        assert health["functionality_test"] == "passed"
        assert "config" in health
        assert "current_model" in health
        assert "available_models" in health
    
    @pytest.mark.asyncio
    async def test_health_check_degraded(self, response_generator, mock_ai_provider):
        """Test health check when AI provider is degraded"""
        mock_ai_provider.health_check.return_value = {"status": "degraded"}
        mock_ai_provider.generate_response.side_effect = Exception("Test error")
        
        health = await response_generator.health_check()
        
        assert health["status"] == "degraded"
        assert "failed:" in health["functionality_test"]
    
    @pytest.mark.asyncio
    async def test_health_check_unhealthy(self, response_generator, mock_ai_provider):
        """Test health check when there's an exception"""
        mock_ai_provider.health_check.side_effect = Exception("Health check failed")
        
        health = await response_generator.health_check()
        
        assert health["status"] == "unhealthy"
        assert "error" in health
        assert health["error"] == "Health check failed"