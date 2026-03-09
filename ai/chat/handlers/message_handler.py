"""
Message handler for processing and validating user messages
"""

import re
from typing import Optional
from ..core.models import ValidationResult, Intent, ChatConfig
from ..core.exceptions import ValidationError


class MessageHandler:
    """Handles message processing and validation"""
    
    def __init__(self, config: ChatConfig):
        self.config = config
        self.max_length = config.max_message_length
    
    def validate_message(self, message: str) -> ValidationResult:
        """
        Validate message format and content
        
        Args:
            message: The message to validate
            
        Returns:
            ValidationResult: Validation result with status and errors
        """
        if not isinstance(message, str):
            return ValidationResult(
                is_valid=False,
                error_message="Message must be a string"
            )
        
        # Check if message is empty or only whitespace
        if not message.strip():
            return ValidationResult(
                is_valid=False,
                error_message="Message cannot be empty"
            )
        
        # Check message length
        if len(message) > self.max_length:
            return ValidationResult(
                is_valid=False,
                error_message=f"Message exceeds maximum length of {self.max_length} characters"
            )
        
        # Check for potentially harmful content (basic check)
        warnings = []
        if self._contains_suspicious_content(message):
            warnings.append("Message contains potentially suspicious content")
        
        return ValidationResult(
            is_valid=True,
            warnings=warnings
        )
    
    def preprocess_message(self, message: str) -> str:
        """
        Preprocess message (clean and normalize)
        
        Args:
            message: The raw message
            
        Returns:
            str: Preprocessed message
        """
        # Strip leading/trailing whitespace
        processed = message.strip()
        
        # Normalize whitespace (replace multiple spaces with single space)
        processed = re.sub(r'\s+', ' ', processed)
        
        # Remove null characters and other control characters
        processed = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', processed)
        
        return processed
    
    def extract_intent(self, message: str) -> Intent:
        """
        Extract intent from message (basic implementation)
        
        Args:
            message: The message to analyze
            
        Returns:
            Intent: Extracted intent information
        """
        # This is a basic implementation - in a real system this would
        # use NLP models or rule-based systems
        
        message_lower = message.lower().strip()
        
        # Simple keyword-based intent detection with word boundaries
        greeting_words = ['hello', 'hi', 'hey', 'greetings']
        if any(re.search(r'\b' + word + r'\b', message_lower) for word in greeting_words):
            return Intent(name="greeting", confidence=0.8)
        
        farewell_words = ['bye', 'goodbye', 'farewell', 'exit']
        if any(re.search(r'\b' + word + r'\b', message_lower) for word in farewell_words):
            return Intent(name="farewell", confidence=0.8)
        
        if message_lower.endswith('?'):
            return Intent(name="question", confidence=0.7)
        
        help_words = ['help', 'assist', 'support']
        if any(re.search(r'\b' + word + r'\b', message_lower) for word in help_words):
            return Intent(name="help_request", confidence=0.9)
        
        # Default intent
        return Intent(name="general", confidence=0.5)
    
    def _contains_suspicious_content(self, message: str) -> bool:
        """
        Check for potentially suspicious content
        
        Args:
            message: The message to check
            
        Returns:
            bool: True if suspicious content is detected
        """
        # Basic checks for suspicious patterns
        suspicious_patterns = [
            r'<script.*?>',  # Script tags
            r'javascript:',  # JavaScript URLs
            r'data:.*base64',  # Base64 data URLs
            r'eval\s*\(',  # eval() calls
        ]
        
        message_lower = message.lower()
        
        for pattern in suspicious_patterns:
            if re.search(pattern, message_lower, re.IGNORECASE):
                return True
        
        return False