"""
Response generator for creating AI responses
"""

import time
import logging
from typing import List, Dict, Any, Optional
from ..core.models import Message, AIResponse, ChatConfig, ChatResponse
from ..core.exceptions import ModelError, AllModelsFailedError
from ..providers.ai_provider import AIProvider

logger = logging.getLogger(__name__)


class ResponseGenerator:
    """
    Generates responses using AI models with formatting and validation
    
    This class integrates with AI providers to generate responses and handles
    response formatting, validation, and error recovery.
    """
    
    def __init__(self, config: ChatConfig, ai_provider: Optional[AIProvider] = None):
        """
        Initialize response generator
        
        Args:
            config: Chat configuration
            ai_provider: Optional AI provider instance (creates new if None)
        """
        self.config = config
        self.ai_provider = ai_provider or AIProvider(config)
        self.logger = logging.getLogger(__name__)
        
        # Response formatting settings
        self.max_response_length = getattr(config, 'max_response_length', 2000)
        self.min_confidence_threshold = getattr(config, 'min_confidence_threshold', 0.1)
        
        # Default fallback responses
        self.fallback_responses = [
            "I'm not sure how to respond to that. Could you please rephrase your question?",
            "I apologize, but I'm having trouble understanding your request. Could you provide more details?",
            "I'm having difficulty processing your message right now. Please try again.",
            "Could you please clarify what you're asking? I want to make sure I give you a helpful response."
        ]
    
    async def generate_response(self, context: List[Message], message: str, session_id: str = "") -> ChatResponse:
        """
        Generate AI response given context and message
        
        Args:
            context: List of previous messages for context
            message: The current user message
            session_id: Session identifier
            
        Returns:
            ChatResponse: The generated chat response
            
        Raises:
            ValueError: If input is invalid
        """
        if not message or not message.strip():
            raise ValueError("Message cannot be empty")
        
        start_time = time.time()
        
        try:
            self.logger.info(f"Generating response for session {session_id}")
            
            # Generate response using AI provider
            ai_response = await self.ai_provider.generate_response(context, message)
            
            # Validate and format response
            formatted_response = self._format_ai_response(ai_response)
            
            # Calculate total processing time
            processing_time = time.time() - start_time
            
            # Create chat response
            chat_response = ChatResponse(
                message=formatted_response.content,
                session_id=session_id,
                response_time=processing_time,
                model_used=formatted_response.model_name,
                confidence=formatted_response.confidence_score,
                metadata={
                    "token_usage": formatted_response.token_usage.__dict__ if formatted_response.token_usage else {},
                    "original_processing_time": formatted_response.processing_time,
                    "total_processing_time": processing_time,
                    "error": formatted_response.error
                }
            )
            
            self.logger.info(f"Successfully generated response in {processing_time:.2f}s using {formatted_response.model_name}")
            return chat_response
            
        except AllModelsFailedError as e:
            self.logger.error(f"All AI models failed: {str(e)}")
            return self._create_error_response(session_id, time.time() - start_time, "service_unavailable")
            
        except Exception as e:
            self.logger.error(f"Error generating response: {str(e)}")
            return self._create_error_response(session_id, time.time() - start_time, "generation_error", str(e))
    
    def _format_ai_response(self, response: AIResponse) -> AIResponse:
        """
        Format and validate AI response
        
        Args:
            response: The raw AI response
            
        Returns:
            AIResponse: The formatted response
        """
        # Ensure response content is not empty
        if not response.content or not response.content.strip():
            response.content = self._get_fallback_response()
            response.confidence_score = self.min_confidence_threshold
            self.logger.warning("AI response was empty, using fallback")
        
        # Trim excessive whitespace and normalize
        response.content = self._normalize_text(response.content)
        
        # Truncate if too long
        if len(response.content) > self.max_response_length:
            response.content = response.content[:self.max_response_length - 3] + "..."
            self.logger.warning(f"Response truncated to {self.max_response_length} characters")
        
        # Ensure confidence score is within valid range
        response.confidence_score = max(0.0, min(1.0, response.confidence_score))
        
        # Validate response quality
        if response.confidence_score < self.min_confidence_threshold:
            self.logger.warning(f"Low confidence response: {response.confidence_score}")
        
        return response
    
    def _normalize_text(self, text: str) -> str:
        """
        Normalize text content
        
        Args:
            text: Raw text content
            
        Returns:
            str: Normalized text
        """
        # Remove excessive whitespace
        text = ' '.join(text.split())
        
        # Ensure proper sentence ending
        if text and not text.endswith(('.', '!', '?', ':', ';')):
            text += '.'
        
        return text.strip()
    
    def _get_fallback_response(self) -> str:
        """
        Get a fallback response when AI fails
        
        Returns:
            str: Fallback response text
        """
        import random
        return random.choice(self.fallback_responses)
    
    def _create_error_response(
        self, 
        session_id: str, 
        processing_time: float, 
        error_type: str, 
        error_details: Optional[str] = None
    ) -> ChatResponse:
        """
        Create an error response
        
        Args:
            session_id: Session identifier
            processing_time: Time taken to process
            error_type: Type of error
            error_details: Optional error details
            
        Returns:
            ChatResponse: Error response
        """
        error_messages = {
            "service_unavailable": "I'm sorry, but I'm currently unable to process your request. Please try again in a moment.",
            "generation_error": "I apologize, but I encountered an error while generating a response. Please try rephrasing your question.",
            "timeout": "I'm taking too long to respond. Please try a simpler question or try again later.",
            "invalid_input": "I'm having trouble understanding your input. Could you please rephrase your question?"
        }
        
        message = error_messages.get(error_type, "I'm sorry, but something went wrong. Please try again.")
        
        return ChatResponse(
            message=message,
            session_id=session_id,
            response_time=processing_time,
            model_used="error_handler",
            confidence=0.0,
            metadata={
                "error_type": error_type,
                "error_details": error_details,
                "is_error_response": True
            }
        )
    
    async def generate_ai_response(self, context: List[Message], message: str) -> AIResponse:
        """
        Generate raw AI response (for internal use)
        
        Args:
            context: List of previous messages for context
            message: The current user message
            
        Returns:
            AIResponse: The raw AI response
        """
        return await self.ai_provider.generate_response(context, message)
    
    def switch_model(self, model_name: str) -> bool:
        """
        Switch to a different AI model
        
        Args:
            model_name: Name of the model to switch to
            
        Returns:
            bool: True if switch was successful
        """
        result = self.ai_provider.switch_model(model_name)
        if result:
            self.logger.info(f"Switched to AI model: {model_name}")
        else:
            self.logger.warning(f"Failed to switch to AI model: {model_name}")
        return result
    
    def get_available_models(self) -> List[str]:
        """
        Get list of available AI models
        
        Returns:
            List[str]: Available model names
        """
        return self.ai_provider.get_available_models()
    
    def get_current_model(self) -> str:
        """
        Get current AI model name
        
        Returns:
            str: Current model name
        """
        return self.ai_provider.get_current_model()
    
    def set_fallback_responses(self, responses: List[str]):
        """
        Set custom fallback responses
        
        Args:
            responses: List of fallback response texts
        """
        if responses and all(isinstance(r, str) and r.strip() for r in responses):
            self.fallback_responses = responses
            self.logger.info(f"Updated fallback responses: {len(responses)} responses")
        else:
            self.logger.warning("Invalid fallback responses provided")
    
    def update_config(self, **kwargs):
        """
        Update response generator configuration
        
        Args:
            **kwargs: Configuration parameters to update
        """
        if 'max_response_length' in kwargs:
            self.max_response_length = kwargs['max_response_length']
        if 'min_confidence_threshold' in kwargs:
            self.min_confidence_threshold = kwargs['min_confidence_threshold']
        
        self.logger.info(f"Updated response generator configuration: {kwargs}")
    
    async def health_check(self) -> Dict[str, Any]:
        """
        Perform health check on response generator
        
        Returns:
            Dict[str, Any]: Health status information
        """
        try:
            # Check AI provider health
            ai_health = await self.ai_provider.health_check()
            
            # Test basic functionality
            test_context = []
            test_message = "Hello"
            
            try:
                test_response = await self.generate_ai_response(test_context, test_message)
                functionality_test = "passed"
            except Exception as e:
                functionality_test = f"failed: {str(e)}"
            
            return {
                "status": "healthy" if ai_health["status"] == "healthy" and functionality_test == "passed" else "degraded",
                "ai_provider": ai_health,
                "functionality_test": functionality_test,
                "config": {
                    "max_response_length": self.max_response_length,
                    "min_confidence_threshold": self.min_confidence_threshold,
                    "fallback_responses_count": len(self.fallback_responses)
                },
                "current_model": self.get_current_model(),
                "available_models": self.get_available_models()
            }
        except Exception as e:
            return {
                "status": "unhealthy",
                "error": str(e)
            }