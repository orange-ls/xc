"""
AI provider for managing AI model integrations
"""

import asyncio
import logging
from typing import List, Dict, Any, Optional
from ..core.models import Message, AIResponse, TokenUsage, ChatConfig
from ..core.interfaces import AIModelProvider
from ..core.exceptions import ModelError, AllModelsFailedError, RetryableError
from .mock_provider import MockAIProvider
from .fallback_chain import FallbackChain
from .retry_strategy import RetryStrategy

logger = logging.getLogger(__name__)


class AIProvider:
    """
    Manages AI model providers and fallback chains
    
    This class provides a unified interface for accessing multiple AI models
    with automatic fallback support and health monitoring.
    """
    
    def __init__(self, config: ChatConfig):
        self.config = config
        self.providers: Dict[str, AIModelProvider] = {}
        self.fallback_chain_order = config.fallback_models or ["mock"]
        self.current_provider_index = 0
        self.logger = logging.getLogger(__name__)
        
        # Initialize providers
        self._initialize_providers()
        
        # Initialize fallback chain
        self.fallback_chain = FallbackChain(self.providers, self.fallback_chain_order)
        
        # Initialize retry strategy
        self.retry_strategy = RetryStrategy(
            max_retries=3,
            base_delay=1.0,
            backoff_factor=2.0,
            jitter=True
        )
    
    def _initialize_providers(self):
        """Initialize AI model providers based on configuration"""
        # Initialize mock provider (always available for testing)
        self.providers["mock"] = MockAIProvider()
        
        # Initialize other providers based on configuration
        for model_name in self.config.ai_models:
            if model_name == "mock":
                continue  # Already initialized
            
            # For now, we only support mock provider
            # Real providers (OpenAI, Anthropic, etc.) will be added in later tasks
            self.logger.warning(f"Provider '{model_name}' not yet implemented, using mock provider")
    
    async def generate_response(self, context: List[Message], message: str) -> AIResponse:
        """
        Generate AI response with fallback support and retry logic
        
        Args:
            context: List of previous messages for context
            message: The current user message
            
        Returns:
            AIResponse: The generated response
            
        Raises:
            AllModelsFailedError: If all models fail
            ValidationError: If input is invalid
        """
        if not message or not message.strip():
            raise ValueError("Message cannot be empty")
        
        if len(message) > self.config.max_message_length:
            raise ValueError(f"Message exceeds maximum length of {self.config.max_message_length} characters")
        
        # Use retry strategy with fallback chain
        async def _generate_with_fallback():
            return await self.fallback_chain.try_next_model(context, message)
        
        try:
            response = await self.retry_strategy.execute_with_retry(_generate_with_fallback)
            self.logger.info("Successfully generated response with fallback and retry")
            return response
        except Exception as e:
            self.logger.error(f"Failed to generate response after retries: {str(e)}")
            raise
    
    def switch_model(self, model_name: str) -> bool:
        """
        Switch to a specific AI model
        
        Args:
            model_name: Name of the model to switch to
            
        Returns:
            bool: True if switch was successful
        """
        if model_name not in self.providers:
            self.logger.warning(f"Cannot switch to unknown provider '{model_name}'")
            return False
        
        if not self.providers[model_name].is_available():
            self.logger.warning(f"Cannot switch to unavailable provider '{model_name}'")
            return False
        
        # Update fallback chain order
        self.fallback_chain.promote_provider(model_name)
        
        self.logger.info(f"Switched to AI model '{model_name}'")
        return True
    
    def get_available_models(self) -> List[str]:
        """
        Get list of available AI models
        
        Returns:
            List[str]: List of available model names
        """
        available = []
        for name, provider in self.providers.items():
            try:
                if provider.is_available():
                    available.append(name)
            except Exception as e:
                self.logger.error(f"Error checking availability of provider '{name}': {str(e)}")
        return available
    
    def get_current_model(self) -> str:
        """
        Get the name of the currently active model
        
        Returns:
            str: Current model name
        """
        next_provider = self.fallback_chain.get_next_provider()
        return next_provider if next_provider else "none"
    
    def get_provider_info(self, provider_name: str) -> Optional[Dict[str, Any]]:
        """
        Get information about a specific provider
        
        Args:
            provider_name: Name of the provider
            
        Returns:
            Dict[str, Any]: Provider information or None if not found
        """
        if provider_name not in self.providers:
            return None
        
        provider = self.providers[provider_name]
        return {
            "name": provider_name,
            "model_name": provider.get_model_name(),
            "available": provider.is_available(),
            "in_fallback_chain": provider_name in self.fallback_chain.fallback_order
        }
    
    def add_provider(self, name: str, provider: AIModelProvider) -> bool:
        """
        Add a new AI provider
        
        Args:
            name: Name for the provider
            provider: The provider instance
            
        Returns:
            bool: True if added successfully
        """
        if name in self.providers:
            self.logger.warning(f"Provider '{name}' already exists")
            return False
        
        self.providers[name] = provider
        self.fallback_chain.update_providers(self.providers)
        self.logger.info(f"Added AI provider '{name}'")
        return True
    
    def remove_provider(self, name: str) -> bool:
        """
        Remove an AI provider
        
        Args:
            name: Name of the provider to remove
            
        Returns:
            bool: True if removed successfully
        """
        if name not in self.providers:
            return False
        
        del self.providers[name]
        self.fallback_chain.update_providers(self.providers)
        self.logger.info(f"Removed AI provider '{name}'")
        return True
    
    def get_fallback_stats(self) -> Dict[str, Any]:
        """
        Get fallback chain statistics
        
        Returns:
            Dict[str, Any]: Fallback statistics
        """
        return self.fallback_chain.get_failure_stats()
    
    def get_retry_config(self) -> Dict[str, Any]:
        """
        Get retry strategy configuration
        
        Returns:
            Dict[str, Any]: Retry configuration
        """
        return self.retry_strategy.get_config()
    
    def update_retry_config(self, **kwargs):
        """
        Update retry strategy configuration
        
        Args:
            **kwargs: Configuration parameters to update
        """
        self.retry_strategy.update_config(**kwargs)
    
    def reorder_fallback_chain(self):
        """Reorder fallback chain based on success rates"""
        self.fallback_chain.reorder_by_success()
    
    def reset_failure_stats(self):
        """Reset failure statistics for all providers"""
        self.fallback_chain.reset_failure_counts()
    
    async def health_check(self) -> Dict[str, Any]:
        """
        Perform health check on all providers
        
        Returns:
            Dict[str, Any]: Health status information
        """
        provider_health = {}
        overall_healthy = False
        
        for name, provider in self.providers.items():
            try:
                is_healthy = await provider.health_check()
                provider_health[name] = {
                    "status": "healthy" if is_healthy else "unhealthy",
                    "available": provider.is_available()
                }
                if is_healthy:
                    overall_healthy = True
            except Exception as e:
                provider_health[name] = {
                    "status": "error",
                    "error": str(e),
                    "available": False
                }
        
        return {
            "status": "healthy" if overall_healthy else "unhealthy",
            "providers": provider_health,
            "fallback_stats": self.get_fallback_stats(),
            "retry_config": self.get_retry_config(),
            "current_model": self.get_current_model()
        }