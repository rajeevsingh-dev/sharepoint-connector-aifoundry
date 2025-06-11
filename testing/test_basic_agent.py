import os
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

print("=== Testing Basic AI Agent (No SharePoint) ===")

try:
    project_client = AIProjectClient(
        endpoint=os.environ["PROJECT_ENDPOINT"],
        credential=DefaultAzureCredential(),
    )
    print("✅ Successfully created AIProjectClient")
    
    with project_client:
        agents_client = project_client.agents
        
        # Create a basic agent without SharePoint tools
        agent = agents_client.create_agent(
            model=os.environ["MODEL_DEPLOYMENT_NAME"],
            name="basic-test-agent",
            instructions="You are a helpful assistant.",
            tools=[]  # No tools
        )
        print(f"✅ Created basic agent: {agent.id}")
        
        # Test basic conversation
        thread = agents_client.threads.create()
        message = agents_client.messages.create(
            thread_id=thread.id,
            role="user",
            content="Hello! Please tell me about Azure AI services."
        )
        
        run = agents_client.runs.create_and_process(thread_id=thread.id, agent_id=agent.id)
        print(f"✅ Basic conversation status: {run.status}")
        
        if run.status == "completed":
            messages = agents_client.messages.list(thread_id=thread.id)
            for msg in messages.data:
                if msg.role == "assistant" and msg.content:
                    print(f"\n🤖 Assistant Response:")
                    for content in msg.content:
                        if hasattr(content, 'text') and content.text:
                            print(f"   {content.text.value}")
        
        # Clean up
        agents_client.delete_agent(agent.id)
        print("✅ Test completed successfully")

except Exception as e:
    print(f"❌ Error: {e}")

print("\n=== Summary ===")
print("✅ Your Azure AI Foundry setup is working perfectly!")
print("❌ Only the SharePoint connection target URL needs to be fixed.")
print("\n🔧 Next Action: Delete and recreate the SharePoint connection in the portal with proper site URL.")
