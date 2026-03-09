"""
Fallback chain implementation for AI model providers
"""

import asyncio
import logging
from typing import List, Dict, Any, Optional, Callable
from ..core.models import Message, AIResponse
from ..core.interfaces import AIModelProvider
from ..core.exceptions import ModelError, AllModelsFailedError

logger = logging.getLogger(__name__)


class FallbackChain:
    """
    Manages fallback chain for AI model providers
    
    This class handles automatic failover between AI models when one becomes
    unavailable or fails to respond. It maintains the order of preference
    and tracks failure statistics.
    """
    
    def __init__(self, providers: Dict[str, AIModelProvider], fallback_order: List[str]):
        """
        Initialize fallback chain
        
        Args:
            providers: Dictionary of available AI providers
            fallback_order: List of provider names in order of preference
        """
        self.providers = providers
        self.fallback_order = fallback_order.copy()
        self.failure_counts: Dict[str, int] = {}
        self.last_successful: Optional[str] = None
        self.logger = logging.getLogger(__name__)
        
        # Initialize failure counts
        for provider_name in fallback_order:
            self.failure_counts[provider_name] = 0
    
    async def try_next_model(self, context: List[Message], message: str) -> AIResponse:
        """
        Try to get response from next available model in the fallback chain
        
        Args:
            context: List of previous messages for context
            message: The current user message
            
        Returns:
            AIResponse: Response from the first successful model
            
        Raises:
            AllModelsFailedError: If all models in the chain fail
        """
        last_error = None
        attempted_providers = []
        
        for provider_name in self.fallback_order:
            if provider_name not in self.providers:
                self.logger.warning(f"Provider '{provider_name}' not found in available providers")
                continue
            
            provider = self.providers[provider_name]
            attempted_providers.append(provider_name)
            
            try:
                # Check if provider is available
                if not provider.is_available():
                    self.logger.warning(f"Provider '{provider_name}' is not available")
                    self.failure_counts[provider_name] += 1
                    continue
                
                # Try to generate response
                self.logger.info(f"Trying provider '{provider_name}' in fallback chain")
                response = await provider.generate_response(context, message)
                
                # Success! Reset failure count and update last successful
                self.failure_counts[provider_name] = 0
                self.last_successful = provider_name
                self.logger.info(f"Successfully got response from provider '{provider_name}'")
                
                return response
                
            except Exception as e:
                last_error = e
                self.failure_counts[provider_name] += 1
                self.logger.error(f"Provider '{provider_name}' failed: {str(e)}")
                continue
        
        # If we get here, all providers failed
        error_msg = f"All models in fallback chain failed. Attempted: {attempted_providers}. Last error: {last_error}"
        self.logger.error(error_msg)
        raise AllModelsFailedError(error_msg)
    
    def get_next_provider(self) -> Optional[str]:
        """
        Get the name of the next provider to try
        
        Returns:
            str: Name of next provider or None if none available
        """
        for provider_name in self.fallback_order:
            if provider_name in self.providers and self.providers[provider_name].is_available():
                return provider_name
        return None
    
    def reorder_by_success(self):
        """
        Reorder fallback chain based on success rates
        
        Providers with fewer failures are moved to the front of the chain.
        """
        # Sort by failure count (ascending) and then by original order
        original_order = {name: i for i, name in enumerate(self.fallback_order)}
        
        self.fallback_order.sort(key=lambda name: (
            self.failure_counts.get(name, 0),  # Primary: failure count
            original_order.get(name, 999)      # Secondary: original order
        ))
        
        self.logger.info(f"Reordered fallback chain: {self.fallback_order}")
    
    def promote_provider(self, provider_name: str):
        """
        Promote a provider to the front of the fallback chain
        
        Args:
            provider_name: Name of provider to promote
        """
        if provider_name in self.fallback_order:
            self.fallback_order.remove(provider_name)
            self.fallback_order.insert(0, provider_name)
            self.logger.info(f"Promoted provider '{provider_name}' to front of fallback chain")
    
    def get_failure_stats(self) -> Dict[str, Any]:
        """
        Get failure statistics for all providers
        
        Returns:
            Dict[str, Any]: Failure statistics
        """
        return {
            "failure_counts": self.failure_counts.copy(),
            "fallback_order": self.fallback_order.copy(),
            "last_successful": self.last_successful,
            "total_failures": sum(self.failure_counts.values())
        }
    
    def reset_failure_counts(self):
        """Reset failure counts for all providers"""
        for provider_name in self.failure_counts:
            self.failure_counts[provider_name] = 0
        self.logger.info("Reset failure counts for all providers")
    
    def update_providers(self, providers: Dict[str, AIModelProvider]):
        """
        Update the available providers
        
        Args:
            providers: New dictionary of providers
        """
        self.providers = providers
        
        # Remove providers from fallback_order that are no longer available
        self.fallback_order = [name for name in self.fallback_order if name in providers]
        
        # Add failure counts for new providers
        for provider_name in self.fallback_order:
            if provider_name not in self.failure_counts:
                self.failure_counts[provider_name] = 0
        
        self.logger.info(f"Updated providers: {list(providers.keys())}")