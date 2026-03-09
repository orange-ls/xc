"""
Configuration management for AI Chat Module
"""

import os
import json
from typing import Optional, Dict, Any
from .core.models import ChatConfig
from .core.exceptions import ConfigurationError


class ConfigManager:
    """Manages configuration loading and validation"""
    
    def __init__(self, config_path: Optional[str] = None):
        self.config_path = config_path or "ai/chat/config.json"
        self.default_config = ChatConfig()
    
    def load_config(self) -> ChatConfig:
        """
        Load configuration from file or use defaults
        
        Returns:
            ChatConfig: Loaded configuration
            
        Raises:
            ConfigurationError: If configuration is invalid
        """
        if os.path.exists(self.config_path):
            try:
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config_data = json.load(f)
                
                # Validate and create config
                return self._create_config_from_dict(config_data)
                
            except (json.JSONDecodeError, IOError) as e:
                raise ConfigurationError(f"Failed to load configuration: {e}")
        else:
            # Return default configuration
            return self.default_config
    
    def save_config(self, config: ChatConfig) -> None:
        """
        Save configuration to file
        
        Args:
            config: Configuration to save
            
        Raises:
            ConfigurationError: If save fails
        """
        try:
            # Create directory if it doesn't exist
            os.makedirs(os.path.dirname(self.config_path), exist_ok=True)
            
            # Convert config to dictionary
            config_dict = self._config_to_dict(config)
            
            # Save to file
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(config_dict, f, indent=2)
                
        except IOError as e:
            raise ConfigurationError(f"Failed to save configuration: {e}")
    
    def validate_config(self, config: ChatConfig) -> bool:
        """
        Validate configuration parameters
        
        Args:
            config: Configuration to validate
            
        Returns:
            bool: True if configuration is valid
            
        Raises:
            ConfigurationError: If configuration is invalid
        """
        if config.max_message_length <= 0:
            raise ConfigurationError("max_message_length must be positive")
        
        if config.context_window_size <= 0:
            raise ConfigurationError("context_window_size must be positive")
        
        if config.response_timeout <= 0:
            raise ConfigurationError("response_timeout must be positive")
        
        if config.session_timeout_hours <= 0:
            raise ConfigurationError("session_timeout_hours must be positive")
        
        if config.max_concurrent_sessions <= 0:
            raise ConfigurationError("max_concurrent_sessions must be positive")
        
        if not config.ai_models:
            raise ConfigurationError("ai_models cannot be empty")
        
        if config.storage_backend not in ["memory", "file", "database"]:
            raise ConfigurationError("storage_backend must be 'memory', 'file', or 'database'")
        
        if config.log_level not in ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]:
            raise ConfigurationError("log_level must be a valid logging level")
        
        return True
    
    def _create_config_from_dict(self, config_data: Dict[str, Any]) -> ChatConfig:
        """Create ChatConfig from dictionary"""
        try:
            config = ChatConfig(
                max_message_length=config_data.get("max_message_length", self.default_config.max_message_length),
                context_window_size=config_data.get("context_window_size", self.default_config.context_window_size),
                response_timeout=config_data.get("response_timeout", self.default_config.response_timeout),
                storage_backend=config_data.get("storage_backend", self.default_config.storage_backend),
                ai_models=config_data.get("ai_models", self.default_config.ai_models),
                fallback_models=config_data.get("fallback_models", self.default_config.fallback_models),
                log_level=config_data.get("log_level", self.default_config.log_level),
                session_timeout_hours=config_data.get("session_timeout_hours", self.default_config.session_timeout_hours),
                max_concurrent_sessions=config_data.get("max_concurrent_sessions", self.default_config.max_concurrent_sessions)
            )
            
            # Validate the configuration
            self.validate_config(config)
            
            return config
            
        except (TypeError, ValueError) as e:
            raise ConfigurationError(f"Invalid configuration format: {e}")
    
    def _config_to_dict(self, config: ChatConfig) -> Dict[str, Any]:
        """Convert ChatConfig to dictionary"""
        return {
            "max_message_length": config.max_message_length,
            "context_window_size": config.context_window_size,
            "response_timeout": config.response_timeout,
            "storage_backend": config.storage_backend,
            "ai_models": config.ai_models,
            "fallback_models": config.fallback_models,
            "log_level": config.log_level,
            "session_timeout_hours": config.session_timeout_hours,
            "max_concurrent_sessions": config.max_concurrent_sessions
        }