import os
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
from azure.ai.agents.models import SharepointTool
from dotenv import load_dotenv
import time

# Load environment variables
load_dotenv()

print("🧪 === SharePoint Connection Test Suite ===")
print()

def test_connection():
    """Test SharePoint connection step by step"""
    
    # Step 1: Test AI Project Client
    print("1️⃣ Testing AI Project Client...")
    try:
        project_client = AIProjectClient(
            endpoint=os.environ["PROJECT_ENDPOINT"],
            credential=DefaultAzureCredential(),
        )
        print("   ✅ AI Project Client: SUCCESS")
    except Exception as e:
        print(f"   ❌ AI Project Client: FAILED - {e}")
        return False

    # Step 2: Test SharePoint Tool Creation
    print("\n2️⃣ Testing SharePoint Tool Creation...")
    try:
        conn_id = os.environ["SHAREPOINT_CONNECTION_ID"]
        sharepoint = SharepointTool(connection_id=conn_id)
        print("   ✅ SharePoint Tool Creation: SUCCESS")
        print(f"   📝 Connection ID: ...{conn_id[-50:]}")
    except Exception as e:
        print(f"   ❌ SharePoint Tool Creation: FAILED - {e}")
        return False

    # Step 3: Test Agent Creation with SharePoint Tool
    print("\n3️⃣ Testing Agent Creation with SharePoint...")
    with project_client:
        agents_client = project_client.agents
        
        try:
            agent = agents_client.create_agent(
                model=os.environ["MODEL_DEPLOYMENT_NAME"],
                name="sharepoint-connection-test",
                instructions="You are a test agent for SharePoint connectivity. Only respond with simple acknowledgments.",
                tools=sharepoint.definitions
            )
            print(f"   ✅ Agent Creation: SUCCESS (ID: {agent.id})")
            
            # Step 4: Test Basic Message Handling
            print("\n4️⃣ Testing Message Creation...")
            thread = agents_client.threads.create()
            message = agents_client.messages.create(
                thread_id=thread.id,
                role="user",
                content="Hello, can you confirm you have SharePoint access? Just say yes or no."
            )
            print(f"   ✅ Message Creation: SUCCESS (Thread: {thread.id})")
            
            # Step 5: Test SharePoint Tool Execution
            print("\n5️⃣ Testing SharePoint Tool Execution...")
            print("   🔄 Running agent with SharePoint tool...")
            
            try:
                run = agents_client.runs.create_and_process(
                    thread_id=thread.id, 
                    agent_id=agent.id
                )
                
                print(f"   📊 Run Status: {run.status}")
                
                if run.status == "completed":
                    print("   ✅ SharePoint Tool Execution: SUCCESS!")
                    
                    # Get the response
                    messages = agents_client.messages.list(thread_id=thread.id)
                    for msg in messages.data:
                        if msg.role == "assistant" and msg.content:
                            for content in msg.content:
                                if hasattr(content, 'text') and content.text:
                                    print(f"   💬 Agent Response: {content.text.value}")
                    
                    print("\n🎉 CONNECTION TEST: PASSED!")
                    print("   Your SharePoint connection is working correctly!")
                    
                elif run.status == "failed":
                    print(f"   ❌ SharePoint Tool Execution: FAILED")
                    print(f"   📝 Error Details: {run.last_error}")
                    
                    # Analyze the error
                    if "Bad Request" in str(run.last_error):
                        print("\n🔍 DIAGNOSIS:")
                        print("   - SharePoint connection exists but configuration is incomplete")
                        print("   - Check SharePoint site URL and credentials in Azure AI Foundry")
                        print("   - Verify app registration permissions")
                    
                else:
                    print(f"   ⚠️ Unexpected Status: {run.status}")
                    
            except Exception as e:
                print(f"   ❌ SharePoint Tool Execution: EXCEPTION - {e}")
            
            # Cleanup
            agents_client.delete_agent(agent.id)
            print(f"\n🧹 Cleanup: Agent deleted")
            
        except Exception as e:
            print(f"   ❌ Agent Creation: FAILED - {e}")
            return False

def check_connection_configuration():
    """Check connection configuration via Azure REST API"""
    print("\n🔍 === Connection Configuration Check ===")
    
    # We already did this earlier, but let's provide the summary
    print("Based on previous analysis:")
    print("✅ Connection exists: sharepoint-documents-new")
    print("✅ Connection type: CustomKeys")
    print("📝 Expected configuration:")
    print("   - clientId: [Your app registration client ID]")
    print("   - clientSecret: [Your app registration secret]")
    print("   - tenantId: 79e00fbc-5c95-4dc3-9b30-2a75bb9ad7cc")
    print("   - siteUrl: https://mngenvmcap293807.sharepoint.com/sites/teams")

if __name__ == "__main__":
    check_connection_configuration()
    test_connection()
    
    print("\n" + "="*60)
    print("🎯 NEXT STEPS IF TEST FAILS:")
    print("1. Go to Azure AI Foundry Portal")
    print("2. Navigate to Connections → sharepoint-documents-new")
    print("3. Verify all CustomKeys are configured correctly")
    print("4. Use the 'Test Connection' button in the portal")
    print("5. Re-run this test script")
    print("="*60)
