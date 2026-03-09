"""
Mock AI provider for testing and development
"""

import asyncio
import random
import logging
from typing import List
from ..core.models import Message, AIResponse, TokenUsage
from ..core.interfaces import AIModelProvider
from ..core.exceptions import ModelError

logger = logging.getLogger(__name__)


class MockAIProvider(AIModelProvider):
    """
    Mock AI provider for testing purposes
    
    This provider simulates AI responses for testing and development.
    It provides realistic response patterns and can simulate various
    failure scenarios for testing error handling.
    """
    
    def __init__(self, failure_rate: float = 0.0, response_delay: tuple = (0.1, 0.5)):
        self.model_name = "mock-ai-v1"
        self.is_healthy = True
        self.failure_rate = failure_rate  # Probability of simulated failures
        self.response_delay = response_delay  # Min and max response delay
        self.logger = logging.getLogger(__name__)
        
        # Predefined responses for different scenarios
        self.responses = [
            "I understand your message. How can I help you further?",
            "That's an interesting point. Could you tell me more about it?",
            "I see what you mean. Let me think about that for a moment.",
            "Thank you for sharing that with me. What would you like to discuss next?",
            "I appreciate your question. Here's what I think about it...",
            "That's a great observation. I'd like to add that...",
            "I'm here to help you with whatever you need. What can I do for you?",
            "Your message is clear. Let me provide you with some thoughts on this topic.",
        ]
        
        self.greeting_responses = [
            "Hello! I'm here to help you. What would you like to talk about?",
            "Hi there! How can I assist you today?",
            "Greetings! I'm ready to help with any questions you might have.",
            "Hello! It's nice to meet you. What can I do for you?",
        ]
        
        self.farewell_responses = [
            "Goodbye! It was nice talking with you.",
            "Farewell! Feel free to come back anytime you need help.",
            "See you later! Have a great day!",
            "Goodbye! I hope our conversation was helpful.",
        ]
        
        self.question_responses = [
            "That's a great question! Let me think about it...",
            "I'd be happy to help you with that question.",
            "Interesting question! Here's what I think...",
            "Let me provide you with some information about that.",
        ]
    
    async def generate_response(self, context: List[Message], message: str) -> AIResponse:
        """
        Generate a mock AI response
        
        Args:
            context: List of previous messages for context
            message: The current user message
            
        Returns:
            AIResponse: Mock AI response
            
        Raises:
            ModelError: If simulated failure occurs
        """
        # Simulate random failures for testing
        if random.random() < self.failure_rate:
            self.logger.warning("Simulating AI provider failure")
            raise ModelError(f"Simulated failure in {self.model_name}")
        
        # Validate input
        if not message or not message.strip():
            raise ValueError("Message cannot be empty")
        
        # Simulate processing time
        delay = random.uniform(*self.response_delay)
        await asyncio.sleep(delay)
        
        # Simple intent-based response selection
        message_lower = message.lower().strip()
        
        if any(word in message_lower for word in ['hello', 'hi', 'hey', 'greetings']):
            response_text = random.choice(self.greeting_responses)
        elif any(word in message_lower for word in ['bye', 'goodbye', 'farewell', 'exit']):
            response_text = random.choice(self.farewell_responses)
        elif '?' in message:
            response_text = random.choice(self.question_responses)
        else:
            response_text = random.choice(self.responses)
        
        # Add context awareness for longer conversations
        if len(context) > 2:
            response_text = f"Continuing our conversation, {response_text.lower()}"
        elif len(context) > 5:
            response_text = f"As we've been discussing, {response_text.lower()}"
        
        # Simulate token usage calculation
        prompt_tokens = len(message.split()) + sum(len(msg.content.split()) for msg in context[-3:])
        completion_tokens = len(response_text.split())
        
        confidence = random.uniform(0.7, 0.95)
        
        self.logger.debug(f"Generated response with {completion_tokens} tokens, confidence: {confidence:.2f}")
        
        return AIResponse(
            content=response_text,
            model_name=self.model_name,
            confidence_score=confidence,
            processing_time=delay,
            token_usage=TokenUsage(
                prompt_tokens=prompt_tokens,
                completion_tokens=completion_tokens,
                total_tokens=prompt_tokens + completion_tokens
            )
        )
    
    def get_model_name(self) -> str:
        """Get the name of this AI model"""
        return self.model_name
    
    def is_available(self) -> bool:
        """Check if the model is currently available"""
        return self.is_healthy
    
    async def health_check(self) -> bool:
        """
        Perform health check on the model
        
        Returns:
            bool: True if healthy, False otherwise
        """
        # Simulate occasional health issues for testing
        if random.random() < 0.05:  # 5% chance of being unhealthy
            self.is_healthy = False
            self.logger.warning("Mock AI provider health check failed")
        else:
            self.is_healthy = True
            self.logger.debug("Mock AI provider health check passed")
        
        return self.is_healthy
    
    def set_failure_rate(self, rate: float):
        """
        Set the failure rate for testing purposes
        
        Args:
            rate: Failure rate between 0.0 and 1.0
        """
        self.failure_rate = max(0.0, min(1.0, rate))
        self.logger.info(f"Set failure rate to {self.failure_rate}")
    
    def set_healthy(self, healthy: bool):
        """
        Manually set health status for testing
        
        Args:
            healthy: Whether the provider should be healthy
        """
        self.is_healthy = healthy
        self.logger.info(f"Set health status to {healthy}")
    
    def set_response_delay(self, min_delay: float, max_delay: float):
        """
        Set response delay range for testing
        
        Args:
            min_delay: Minimum delay in seconds
            max_delay: Maximum delay in seconds
        """
        self.response_delay = (min_delay, max_delay)
        self.logger.info(f"Set response delay to {min_delay}-{max_delay} seconds")