"""
Unit tests for MessageHandler
"""

import pytest
from ..handlers.message_handler import MessageHandler
from ..core.models import ChatConfig, ValidationResult, Intent


class TestMessageHandler:
    """Test MessageHandler functionality"""
    
    @pytest.fixture
    def handler(self):
        """Create MessageHandler instance for testing"""
        config = ChatConfig(max_message_length=100)  # Smaller limit for testing
        return MessageHandler(config)
    
    def test_validate_message_valid(self, handler):
        """Test validation of valid message"""
        result = handler.validate_message("Hello world")
        
        assert result.is_valid is True
        assert result.error_message is None
        assert isinstance(result.warnings, list)
    
    def test_validate_message_empty(self, handler):
        """Test validation of empty message"""
        result = handler.validate_message("")
        
        assert result.is_valid is False
        assert "cannot be empty" in result.error_message
    
    def test_validate_message_whitespace_only(self, handler):
        """Test validation of whitespace-only message"""
        result = handler.validate_message("   \n\t  ")
        
        assert result.is_valid is False
        assert "cannot be empty" in result.error_message
    
    def test_validate_message_too_long(self, handler):
        """Test validation of message exceeding length limit"""
        long_message = "x" * 101  # Exceeds the 100 char limit set in fixture
        result = handler.validate_message(long_message)
        
        assert result.is_valid is False
        assert "exceeds maximum length" in result.error_message
        assert "100 characters" in result.error_message
    
    def test_validate_message_not_string(self, handler):
        """Test validation of non-string input"""
        result = handler.validate_message(123)
        
        assert result.is_valid is False
        assert "must be a string" in result.error_message
    
    def test_validate_message_suspicious_content(self, handler):
        """Test validation with suspicious content"""
        result = handler.validate_message("Hello <script>alert('xss')</script>")
        
        assert result.is_valid is True  # Still valid but with warnings
        assert len(result.warnings) > 0
        assert "suspicious content" in result.warnings[0]
    
    def test_preprocess_message_basic(self, handler):
        """Test basic message preprocessing"""
        result = handler.preprocess_message("  Hello   world  \n")
        
        assert result == "Hello world"
    
    def test_preprocess_message_multiple_spaces(self, handler):
        """Test preprocessing with multiple spaces"""
        result = handler.preprocess_message("Hello     world    test")
        
        assert result == "Hello world test"
    
    def test_preprocess_message_control_characters(self, handler):
        """Test preprocessing removes control characters"""
        message_with_control = "Hello\x00\x01world\x7F"
        result = handler.preprocess_message(message_with_control)
        
        assert result == "Helloworld"  # Control chars removed but no space added
    
    def test_extract_intent_greeting(self, handler):
        """Test intent extraction for greeting"""
        intent = handler.extract_intent("Hello there!")
        
        assert intent.name == "greeting"
        assert intent.confidence > 0.5
    
    def test_extract_intent_farewell(self, handler):
        """Test intent extraction for farewell"""
        intent = handler.extract_intent("Goodbye!")
        
        assert intent.name == "farewell"
        assert intent.confidence > 0.5
    
    def test_extract_intent_question(self, handler):
        """Test intent extraction for question"""
        intent = handler.extract_intent("What is the weather like?")
        
        assert intent.name == "question"
        assert intent.confidence > 0.5
    
    def test_extract_intent_help_request(self, handler):
        """Test intent extraction for help request"""
        intent = handler.extract_intent("I need help with this")
        
        assert intent.name == "help_request"
        assert intent.confidence > 0.5
    
    def test_extract_intent_general(self, handler):
        """Test intent extraction for general message"""
        intent = handler.extract_intent("The weather is nice today.")
        
        assert intent.name == "general"
        assert intent.confidence > 0
    
    def test_suspicious_content_detection_script_tag(self, handler):
        """Test detection of script tags"""
        assert handler._contains_suspicious_content("<script>alert('test')</script>")
    
    def test_suspicious_content_detection_javascript_url(self, handler):
        """Test detection of javascript URLs"""
        assert handler._contains_suspicious_content("javascript:alert('test')")
    
    def test_suspicious_content_detection_base64_data(self, handler):
        """Test detection of base64 data URLs"""
        assert handler._contains_suspicious_content("data:text/html;base64,PHNjcmlwdD4=")
    
    def test_suspicious_content_detection_eval(self, handler):
        """Test detection of eval calls"""
        assert handler._contains_suspicious_content("eval('malicious code')")
    
    def test_no_suspicious_content(self, handler):
        """Test normal content is not flagged as suspicious"""
        assert not handler._contains_suspicious_content("Hello, how are you today?")