# Test based on the official Azure sample format
import os
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
from azure.ai.agents.models import SharepointTool
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

print("=== Testing Official Sample Format ===")
print()

# Create client exactly like the official sample
project_client = AIProjectClient(
    endpoint=os.environ["PROJECT_ENDPOINT"],
    credential=DefaultAzureCredential(),
)

conn_id = os.environ["SHAREPOINT_CONNECTION_ID"]

# Initialize SharePoint tool with connection id
sharepoint = SharepointTool(connection_id=conn_id)

print("✅ Created AI Project Client and SharePoint Tool")
print(f"🔗 Connection ID: {conn_id}")

# Test with the exact format from the official sample
test_queries = [
    # Official sample format
    "Hello, summarize the key points of the <sharepoint_resource_document>",
    
    # Variations based on our document
    "Hello, summarize the key points of the <Doc to test.docx>",
    
    # Try without angle brackets
    "Hello, summarize the key points of the sharepoint_resource_document",
    "Hello, summarize the key points of Doc to test.docx",
    
    # Basic queries
    "Hello, what SharePoint documents can you access?",
    "List available SharePoint resources"
]

with project_client:
    agents_client = project_client.agents
    
    # Create agent exactly like the official sample
    agent = agents_client.create_agent(
        model=os.environ["MODEL_DEPLOYMENT_NAME"],
        name="my-agent",
        instructions="You are a helpful agent",
        tools=sharepoint.definitions,
    )
    print(f"✅ Created agent, ID: {agent.id}")
    
    for i, query in enumerate(test_queries, 1):
        print(f"\n--- Test {i}: {query} ---")
        
        try:
            # Create thread for communication
            thread = agents_client.threads.create()
            print(f"Created thread, ID: {thread.id}")
            
            # Create message to thread
            message = agents_client.messages.create(
                thread_id=thread.id,
                role="user",
                content=query,
            )
            print(f"Created message, ID: {message.id}")
            
            # Create and process agent run in thread with tools
            run = agents_client.runs.create_and_process(thread_id=thread.id, agent_id=agent.id)
            print(f"Run finished with status: {run.status}")
            
            if run.status == "failed":
                print(f"❌ Run failed: {run.last_error}")
            elif run.status == "completed":
                print("✅ SUCCESS!")
                
                # Fetch and log all messages
                messages = agents_client.messages.list(thread_id=thread.id)
                for msg in messages:
                    if msg.role == "assistant" and msg.content:
                        for content in msg.content:
                            if hasattr(content, 'text') and content.text:
                                print(f"📝 Assistant: {content.text.value[:300]}...")
                                break
                # If we get a successful result, we can stop testing
                break
            else:
                print(f"⏳ Status: {run.status}")
                
        except Exception as e:
            print(f"❌ Exception: {e}")
    
    # Delete the agent when done
    agents_client.delete_agent(agent.id)
    print(f"\n✅ Deleted agent: {agent.id}")

print("\n=== Test Complete ===")
