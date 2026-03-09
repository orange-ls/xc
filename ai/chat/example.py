"""
Simple example demonstrating AI Chat Module usage
"""

import asyncio
import sys
import os

# Add the parent directory to the path so we can import the ai.chat module
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ai.chat import ChatModule, ChatConfig


async def main():
    """Main example function"""
    print("=== AI Chat Module Example ===\n")
    
    # Create configuration
    config = ChatConfig(
        max_message_length=1000,
        context_window_size=5,
        ai_models=["mock"],
        log_level="INFO"
    )
    
    # Initialize chat module
    chat = ChatModule(config)
    print("✅ Chat module initialized")
    
    # Create a session
    session_id = chat.create_session()
    print(f"✅ Created session: {session_id}")
    
    # Test messages
    test_messages = [
        "Hello, how are you?",
        "What's the weather like today?",
        "Can you help me with a programming question?",
        "Tell me about artificial intelligence",
        "Thank you for your help!"
    ]
    
    print("\n=== Conversation ===")
    
    # Process each message
    for i, message in enumerate(test_messages, 1):
        print(f"\n👤 User: {message}")
        
        try:
            response = await chat.process_message(session_id, message)
            print(f"🤖 Assistant: {response.message}")
            print(f"   (Response time: {response.response_time:.3f}s, Model: {response.model_used})")
            
        except Exception as e:
            print(f"❌ Error: {e}")
    
    # Get session info
    session_info = await chat.get_session_info(session_id)
    if session_info:
        print(f"\n📊 Session Info:")
        print(f"   Messages: {session_info.message_count}")
        print(f"   Created: {session_info.created_at}")
        print(f"   Status: {session_info.status.value}")
    
    # Health check
    health = await chat.health_check()
    print(f"\n🏥 System Health: {health['status']}")
    
    # List active sessions
    active_sessions = await chat.list_active_sessions()
    print(f"📋 Active sessions: {len(active_sessions)}")
    
    # End session
    ended = await chat.end_session(session_id)
    print(f"🔚 Session ended: {ended}")
    
    print("\n=== Example completed successfully! ===")


if __name__ == "__main__":
    asyncio.run(main())