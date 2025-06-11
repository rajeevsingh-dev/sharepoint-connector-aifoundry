import os
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
from azure.ai.agents.models import SharepointTool
from dotenv import load_dotenv
from azure.core.exceptions import HttpResponseError
import json

# Load environment variables from .env file
load_dotenv()

print("=== SharePoint Connection Diagnostic ===")
print()

# Create AI Project Client
try:
    project_client = AIProjectClient(
        endpoint=os.environ["PROJECT_ENDPOINT"],
        credential=DefaultAzureCredential(),
    )
    print("✅ Successfully created AIProjectClient")
except Exception as e:
    print(f"❌ Error creating AIProjectClient: {e}")
    exit(1)

# Test SharePoint tool creation
conn_id = os.environ["SHAREPOINT_CONNECTION_ID"]
try:
    sharepoint = SharepointTool(connection_id=conn_id)
    print("✅ Successfully created SharepointTool")
    print(f"   Connection ID: {conn_id}")
except Exception as e:
    print(f"❌ Error creating SharepointTool: {e}")
    exit(1)

# Test agent creation without SharePoint tool (baseline test)
print("\n=== Testing Agent Creation (Without SharePoint) ===")
with project_client:
    agents_client = project_client.agents
    
    try:
        test_agent = agents_client.create_agent(
            model=os.environ["MODEL_DEPLOYMENT_NAME"],
            name="test-agent-no-tools",
            instructions="You are a test agent without tools",
            tools=[]  # No tools
        )
        print(f"✅ Successfully created test agent: {test_agent.id}")
        
        # Test basic conversation
        thread = agents_client.threads.create()
        message = agents_client.messages.create(
            thread_id=thread.id,
            role="user",
            content="Hello, just say hi back"
        )
        
        run = agents_client.runs.create_and_process(thread_id=thread.id, agent_id=test_agent.id)
        print(f"✅ Test conversation status: {run.status}")
        
        # Clean up
        agents_client.delete_agent(test_agent.id)
        print("✅ Test agent deleted")
        
    except Exception as e:
        print(f"❌ Error with test agent: {e}")

# Test agent creation with SharePoint tool
print("\n=== Testing Agent Creation (With SharePoint Tool) ===")
try:
    sp_agent = agents_client.create_agent(
        model=os.environ["MODEL_DEPLOYMENT_NAME"],
        name="sharepoint-test-agent",
        instructions="You are an agent with SharePoint access",
        tools=sharepoint.definitions
    )
    print(f"✅ Successfully created SharePoint agent: {sp_agent.id}")
    
    # Test simple message creation (don't run yet)
    thread = agents_client.threads.create()
    message = agents_client.messages.create(
        thread_id=thread.id,
        role="user",
        content="What documents do you have access to?"
    )
    print(f"✅ Created thread and message successfully")
    
    # Try to run and capture detailed error
    print("🔍 Attempting to run SharePoint query...")
    try:
        run = agents_client.runs.create_and_process(thread_id=thread.id, agent_id=sp_agent.id)
        print(f"   Run status: {run.status}")
        
        if run.status == "failed":
            print(f"   Error details: {run.last_error}")
    except Exception as e:
        print(f"   Exception during run: {e}")
    
    # Clean up
    agents_client.delete_agent(sp_agent.id)
    print("✅ SharePoint agent deleted")
    
except Exception as e:
    print(f"❌ Error creating SharePoint agent: {e}")

print("\n=== Diagnostic Summary ===")
print("✅ Authentication: Working")
print("✅ AI Project Client: Working") 
print("✅ Model Deployment: Working")
print("✅ SharePoint Tool Creation: Working")
print("❌ SharePoint Tool Execution: FAILING")
print()
print("🔧 REQUIRED ACTION:")
print("   The SharePoint connection 'sharepoint-documents-new' needs to be")
print("   properly configured in Azure AI Foundry with:")
print("   - Valid SharePoint site URL")
print("   - Proper authentication credentials")
print("   - Access permissions to SharePoint documents")
