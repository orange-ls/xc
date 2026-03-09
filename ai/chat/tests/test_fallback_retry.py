"""
Tests for fallback chain and retry strategy
"""

import pytest
import asyncio
from unittest.mock import Mock, AsyncMock
from ai.chat.core.models import Message, AIResponse, MessageRole, TokenUsage
from ai.chat.core.exceptions import ModelError, AllModelsFailedError, RetryableError
from ai.chat.providers.fallback_chain import FallbackChain
from ai.chat.providers.retry_strategy import RetryStrategy, CircuitBreaker
from ai.chat.providers.mock_provider import MockAIProvider


class TestFallbackChain:
    """Test cases for FallbackChain"""
    
    @pytest.fixture
    def mock_providers(self):
        """Create mock providers for testing"""
        provider1 = MockAIProvider()
        provider1.model_name = "provider1"
        
        provider2 = MockAIProvider()
        provider2.model_name = "provider2"
        
        provider3 = MockAIProvider()
        provider3.model_name = "provider3"
        
        return {
            "provider1": provider1,
            "provider2": provider2,
            "provider3": provider3
        }
    
    @pytest.fixture
    def fallback_chain(self, mock_providers):
        """Create fallback chain for testing"""
        return FallbackChain(mock_providers, ["provider1", "provider2", "provider3"])
    
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
    
    def test_fallback_chain_initialization(self, fallback_chain):
        """Test fallback chain initialization"""
        assert len(fallback_chain.providers) == 3
        assert fallback_chain.fallback_order == ["provider1", "provider2", "provider3"]
        assert all(count == 0 for count in fallback_chain.failure_counts.values())
        assert fallback_chain.last_successful is None
    
    @pytest.mark.asyncio
    async def test_try_next_model_success_first(self, fallback_chain, sample_context):
        """Test successful response from first provider"""
        response = await fallback_chain.try_next_model(sample_context, "Hello")
        
        assert isinstance(response, AIResponse)
        assert response.model_name == "provider1"
        assert fallback_chain.last_successful == "provider1"
        assert fallback_chain.failure_counts["provider1"] == 0
    
    @pytest.mark.asyncio
    async def test_try_next_model_fallback_success(self, fallback_chain, sample_context):
        """Test fallback to second provider when first fails"""
        # Make first provider unavailable
        fallback_chain.providers["provider1"].set_healthy(False)
        
        response = await fallback_chain.try_next_model(sample_context, "Hello")
        
        assert isinstance(response, AIResponse)
        assert response.model_name == "provider2"
        assert fallback_chain.last_successful == "provider2"
        assert fallback_chain.failure_counts["provider1"] == 1
        assert fallback_chain.failure_counts["provider2"] == 0
    
    @pytest.mark.asyncio
    async def test_try_next_model_all_fail(self, fallback_chain, sample_context):
        """Test when all providers fail"""
        # Make all providers fail
        for provider in fallback_chain.providers.values():
            provider.set_failure_rate(1.0)
        
        with pytest.raises(AllModelsFailedError):
            await fallback_chain.try_next_model(sample_context, "Hello")
        
        # Check that all providers have failure counts
        assert all(count > 0 for count in fallback_chain.failure_counts.values())
    
    def test_get_next_provider(self, fallback_chain):
        """Test getting next available provider"""
        # All healthy
        assert fallback_chain.get_next_provider() == "provider1"
        
        # First unhealthy
        fallback_chain.providers["provider1"].set_healthy(False)
        assert fallback_chain.get_next_provider() == "provider2"
        
        # All unhealthy
        for provider in fallback_chain.providers.values():
            provider.set_healthy(False)
        assert fallback_chain.get_next_provider() is None
    
    def test_reorder_by_success(self, fallback_chain):
        """Test reordering by success rates"""
        # Simulate failures
        fallback_chain.failure_counts["provider1"] = 5
        fallback_chain.failure_counts["provider2"] = 2
        fallback_chain.failure_counts["provider3"] = 0
        
        fallback_chain.reorder_by_success()
        
        # Should be ordered by failure count (ascending)
        assert fallback_chain.fallback_order == ["provider3", "provider2", "provider1"]
    
    def test_promote_provider(self, fallback_chain):
        """Test promoting a provider to front"""
        fallback_chain.promote_provider("provider3")
        assert fallback_chain.fallback_order[0] == "provider3"
        assert "provider3" in fallback_chain.fallback_order
        assert len(fallback_chain.fallback_order) == 3
    
    def test_get_failure_stats(self, fallback_chain):
        """Test getting failure statistics"""
        fallback_chain.failure_counts["provider1"] = 3
        fallback_chain.last_successful = "provider2"
        
        stats = fallback_chain.get_failure_stats()
        
        assert stats["failure_counts"]["provider1"] == 3
        assert stats["last_successful"] == "provider2"
        assert stats["total_failures"] == 3
        assert "fallback_order" in stats
    
    def test_reset_failure_counts(self, fallback_chain):
        """Test resetting failure counts"""
        fallback_chain.failure_counts["provider1"] = 5
        fallback_chain.failure_counts["provider2"] = 3
        
        fallback_chain.reset_failure_counts()
        
        assert all(count == 0 for count in fallback_chain.failure_counts.values())
    
    def test_update_providers(self, fallback_chain):
        """Test updating providers"""
        new_provider = MockAIProvider()
        new_providers = {"new_provider": new_provider}
        
        fallback_chain.update_providers(new_providers)
        
        assert fallback_chain.providers == new_providers


class TestRetryStrategy:
    """Test cases for RetryStrategy"""
    
    @pytest.fixture
    def retry_strategy(self):
        """Create retry strategy for testing"""
        return RetryStrategy(
            max_retries=3,
            base_delay=0.1,  # Short delay for testing
            backoff_factor=2.0,
            jitter=False  # Disable jitter for predictable testing
        )
    
    def test_retry_strategy_initialization(self, retry_strategy):
        """Test retry strategy initialization"""
        assert retry_strategy.max_retries == 3
        assert retry_strategy.base_delay == 0.1
        assert retry_strategy.backoff_factor == 2.0
        assert retry_strategy.jitter is False
    
    def test_calculate_delay(self, retry_strategy):
        """Test delay calculation"""
        assert retry_strategy.calculate_delay(0) == 0.1
        assert retry_strategy.calculate_delay(1) == 0.2
        assert retry_strategy.calculate_delay(2) == 0.4
    
    def test_calculate_delay_with_max(self):
        """Test delay calculation with maximum"""
        strategy = RetryStrategy(base_delay=10.0, max_delay=15.0, backoff_factor=2.0, jitter=False)
        
        assert strategy.calculate_delay(0) == 10.0
        assert strategy.calculate_delay(1) == 15.0  # Capped at max_delay
        assert strategy.calculate_delay(2) == 15.0  # Still capped
    
    def test_calculate_delay_with_jitter(self):
        """Test delay calculation with jitter"""
        strategy = RetryStrategy(base_delay=1.0, jitter=True)
        
        delay1 = strategy.calculate_delay(0)
        delay2 = strategy.calculate_delay(0)
        
        # With jitter, delays should be different
        assert delay1 != delay2
        assert 0.1 <= delay1 <= 2.0  # Should be within reasonable range
        assert 0.1 <= delay2 <= 2.0
    
    def test_is_retryable(self, retry_strategy):
        """Test retryable exception checking"""
        assert retry_strategy.is_retryable(RetryableError("test"))
        assert retry_strategy.is_retryable(ModelError("test"))
        assert not retry_strategy.is_retryable(ValueError("test"))
    
    @pytest.mark.asyncio
    async def test_execute_with_retry_success_first_attempt(self, retry_strategy):
        """Test successful execution on first attempt"""
        async def success_func():
            return "success"
        
        result = await retry_strategy.execute_with_retry(success_func)
        assert result == "success"
    
    @pytest.mark.asyncio
    async def test_execute_with_retry_success_after_retries(self, retry_strategy):
        """Test successful execution after retries"""
        call_count = 0
        
        async def retry_then_success():
            nonlocal call_count
            call_count += 1
            if call_count < 3:
                raise RetryableError("temporary failure")
            return "success"
        
        result = await retry_strategy.execute_with_retry(retry_then_success)
        assert result == "success"
        assert call_count == 3
    
    @pytest.mark.asyncio
    async def test_execute_with_retry_all_attempts_fail(self, retry_strategy):
        """Test when all retry attempts fail"""
        async def always_fail():
            raise RetryableError("persistent failure")
        
        with pytest.raises(RetryableError):
            await retry_strategy.execute_with_retry(always_fail)
    
    @pytest.mark.asyncio
    async def test_execute_with_retry_non_retryable_exception(self, retry_strategy):
        """Test with non-retryable exception"""
        async def non_retryable_fail():
            raise ValueError("non-retryable error")
        
        with pytest.raises(ValueError):
            await retry_strategy.execute_with_retry(non_retryable_fail)
    
    def test_get_config(self, retry_strategy):
        """Test getting configuration"""
        config = retry_strategy.get_config()
        
        assert config["max_retries"] == 3
        assert config["base_delay"] == 0.1
        assert config["backoff_factor"] == 2.0
        assert config["jitter"] is False
    
    def test_update_config(self, retry_strategy):
        """Test updating configuration"""
        retry_strategy.update_config(max_retries=5, base_delay=0.2)
        
        assert retry_strategy.max_retries == 5
        assert retry_strategy.base_delay == 0.2
        assert retry_strategy.backoff_factor == 2.0  # Unchanged


class TestCircuitBreaker:
    """Test cases for CircuitBreaker"""
    
    @pytest.fixture
    def circuit_breaker(self):
        """Create circuit breaker for testing"""
        return CircuitBreaker(
            failure_threshold=3,
            recovery_timeout=1.0,  # Short timeout for testing
            expected_exception=Exception
        )
    
    def test_circuit_breaker_initialization(self, circuit_breaker):
        """Test circuit breaker initialization"""
        assert circuit_breaker.failure_threshold == 3
        assert circuit_breaker.recovery_timeout == 1.0
        assert circuit_breaker.state == "closed"
        assert circuit_breaker.failure_count == 0
    
    @pytest.mark.asyncio
    async def test_circuit_breaker_success(self, circuit_breaker):
        """Test successful call through circuit breaker"""
        async def success_func():
            return "success"
        
        result = await circuit_breaker.call(success_func)
        assert result == "success"
        assert circuit_breaker.state == "closed"
        assert circuit_breaker.failure_count == 0
    
    @pytest.mark.asyncio
    async def test_circuit_breaker_opens_after_failures(self, circuit_breaker):
        """Test circuit breaker opens after threshold failures"""
        async def failing_func():
            raise Exception("failure")
        
        # Fail up to threshold
        for i in range(3):
            with pytest.raises(Exception):
                await circuit_breaker.call(failing_func)
        
        assert circuit_breaker.state == "open"
        assert circuit_breaker.failure_count == 3
        
        # Next call should fail immediately due to open circuit
        with pytest.raises(Exception, match="Circuit breaker is open"):
            await circuit_breaker.call(failing_func)
    
    @pytest.mark.asyncio
    async def test_circuit_breaker_half_open_recovery(self, circuit_breaker):
        """Test circuit breaker recovery through half-open state"""
        async def failing_func():
            raise Exception("failure")
        
        async def success_func():
            return "success"
        
        # Open the circuit
        for i in range(3):
            with pytest.raises(Exception):
                await circuit_breaker.call(failing_func)
        
        assert circuit_breaker.state == "open"
        
        # Wait for recovery timeout
        await asyncio.sleep(1.1)
        
        # Next call should move to half-open and succeed
        result = await circuit_breaker.call(success_func)
        assert result == "success"
        assert circuit_breaker.state == "closed"
        assert circuit_breaker.failure_count == 0
    
    def test_get_state(self, circuit_breaker):
        """Test getting circuit breaker state"""
        state = circuit_breaker.get_state()
        
        assert state["state"] == "closed"
        assert state["failure_count"] == 0
        assert state["failure_threshold"] == 3
        assert "last_failure_time" in state
        assert "recovery_timeout" in state