"""
Basic functionality test for AI Chat Module
"""

import asyncio
import sys
import os

# Add current directory to path
sys.path.insert(0, '.')

from ai.chat.core.chat_module import ChatModule
from ai.chat.core.models import ChatConfig


async def test_basic_functionality():
    """Test basic chat module functionality"""
    print("=== Testing AI Chat Module Basic Functionality ===\n")
    
    # Create configuration
    config = ChatConfig(
        max_message_length=1000,
        context_window_size=5,
        ai_models=["mock"],
        log_level="INFO"
    )
    
    # Initialize chat module
    chat = ChatModule(config)
    print("✅ Chat module initialized successfully")
    
    # Create a session
    session_id = chat.create_session()
    print(f"✅ Created session: {session_id}")
    
    # Test message processing
    test_message = "Hello, how are you?"
    print(f"\n👤 User: {test_message}")
    
    try:
        response = await chat.process_message(session_id, test_message)
        print(f"🤖 Assistant: {response.message}")
        print(f"   Response time: {response.response_time:.3f}s")
        print(f"   Model used: {response.model_used}")
        print("✅ Message processing successful")
        
    except Exception as e:
        print(f"❌ Error processing message: {e}")
        return False
    
    # Test session info
    session_info = await chat.get_session_info(session_id)
    if session_info:
        print(f"\n📊 Session Info:")
        print(f"   ID: {session_info.id}")
        print(f"   Messages: {session_info.message_count}")
        print(f"   Status: {session_info.status.value}")
        print("✅ Session info retrieval successful")
    
    # Test health check
    health = await chat.health_check()
    print(f"\n🏥 System Health: {health['status']}")
    print("✅ Health check successful")
    
    # End session
    ended = await chat.end_session(session_id)
    print(f"\n🔚 Session ended: {ended}")
    print("✅ Session termination successful")
    
    print("\n=== All basic functionality tests passed! ===")
    return True


if __name__ == "__main__":
    success = asyncio.run(test_basic_functionality())
    if success:
        print("\n🎉 AI Chat Module is working correctly!")
    else:
        print("\n❌ Some tests failed.")
        sys.exit(1)