# AI Chat Module

A modular conversational AI system for handling user interactions, maintaining conversation context, and providing intelligent responses.

## Features

- **Modular Architecture**: Clean separation of concerns with well-defined interfaces
- **Multiple AI Model Support**: Pluggable AI providers with fallback chains
- **Context Management**: Intelligent conversation context and history management
- **Session Management**: Support for multiple concurrent conversation sessions
- **Flexible Storage**: Multiple storage backends (memory, file, database)
- **Error Handling**: Robust error handling with graceful degradation
- **Configuration**: Flexible configuration system with hot-reload support
- **Testing**: Comprehensive test suite with property-based testing

## Quick Start

```python
from ai.chat import ChatModule, ChatConfig

# Create configuration
config = ChatConfig(
    max_message_length=4000,
    context_window_size=10,
    ai_models=["mock"]
)

# Initialize chat module
chat = ChatModule(config)

# Create a session
session_id = chat.create_session()

# Process messages
response = await chat.process_message(session_id, "Hello, how are you?")
print(response.message)
```

## Project Structure

```
ai/chat/
├── core/                   # Core models and interfaces
│   ├── models.py          # Data models
│   ├── interfaces.py      # Abstract interfaces
│   ├── exceptions.py      # Custom exceptions
│   └── chat_module.py     # Main ChatModule class
├── handlers/              # Message and request handlers
│   └── message_handler.py # Message processing and validation
├── managers/              # Management components
│   ├── session_manager.py # Session lifecycle management
│   └── conversation_manager.py # Conversation context management
├── generators/            # Response generation
│   └── response_generator.py # AI response generation
├── providers/             # AI model providers
│   ├── ai_provider.py     # Main AI provider coordinator
│   └── mock_provider.py   # Mock provider for testing
├── storage/               # Storage backends
│   ├── context_store.py   # Context storage coordinator
│   └── memory_storage.py  # In-memory storage implementation
├── tests/                 # Test suite
│   ├── conftest.py        # Test configuration
│   ├── test_models.py     # Model tests
│   └── test_chat_module.py # Core functionality tests
├── config.py              # Configuration management
├── requirements.txt       # Dependencies
├── pytest.ini            # Test configuration
└── README.md             # This file
```

## Configuration

The chat module can be configured through a `ChatConfig` object or a JSON configuration file:

```python
config = ChatConfig(
    max_message_length=4000,      # Maximum message length
    context_window_size=10,       # Number of messages to keep in context
    response_timeout=5,           # Response timeout in seconds
    storage_backend="memory",     # Storage backend type
    ai_models=["mock"],          # Available AI models
    fallback_models=["mock"],    # Fallback model chain
    log_level="INFO",            # Logging level
    session_timeout_hours=24,    # Session expiry time
    max_concurrent_sessions=100  # Maximum concurrent sessions
)
```

## Testing

Run the test suite:

```bash
cd ai/chat
pip install -r requirements.txt
pytest
```

Run specific test categories:

```bash
pytest -m unit          # Unit tests only
pytest -m integration   # Integration tests only
pytest -m property      # Property-based tests only
```

## Development Status

This is the initial implementation with core interfaces and basic functionality. The following components are implemented:

- ✅ Core data models and interfaces
- ✅ Basic ChatModule functionality
- ✅ Message handling and validation
- ✅ Session management
- ✅ Memory storage backend
- ✅ Mock AI provider
- ✅ Basic test framework

## Next Steps

The following features will be implemented in subsequent tasks:

- File and database storage backends
- Real AI model integrations (OpenAI, Anthropic, etc.)
- Advanced error handling and retry mechanisms
- Configuration hot-reload
- Comprehensive logging and monitoring
- Plugin architecture
- Performance optimizations

## License

This project is part of the larger AI system and follows the same licensing terms.