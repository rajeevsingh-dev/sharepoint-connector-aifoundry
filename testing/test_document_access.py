"""
Test specific document access patterns for SharePoint connector
"""
import os
from dotenv import load_dotenv
from azure.ai.projects import AIProjectClient
from azure.ai.projects.models import SharepointTool
from azure.ai.projects.models import CodeInterpreterTool, FunctionTool, BingGroundingTool
from azure.identity import DefaultAzureCredential

# Load environment variables
load_dotenv()

def test_document_access():
    """Test different ways to access SharePoint documents"""
    
    # Environment variables
    project_endpoint = os.getenv("PROJECT_ENDPOINT")
    sharepoint_connection_id = os.getenv("SHAREPOINT_CONNECTION_ID")
    model_deployment_name = os.getenv("MODEL_DEPLOYMENT_NAME")
    document_name = os.getenv("document_name", "Doc to test.docx")
    
    print("=== SharePoint Document Access Testing ===")
    print(f"Project Endpoint: {project_endpoint}")
    print(f"Connection ID: {sharepoint_connection_id}")
    print(f"Document: {document_name}")
    print()
    
    try:
        # Create AI Project Client
        ai_project_client = AIProjectClient.from_connection_string(
            conn_str=project_endpoint, 
            credential=DefaultAzureCredential()
        )
        print("✅ AI Project Client created successfully")
        
        # Create SharePoint tool
        sharepoint_tool = SharepointTool(sharepoint_connection_id=sharepoint_connection_id)
        print("✅ SharePoint Tool created successfully")
        
        # Test different query patterns
        test_queries = [
            # Direct document reference
            f"Can you find and summarize the document named '{document_name}'?",
            
            # Library-specific queries  
            "Please list all documents in the Shared Documents library.",
            "What files are available in the Documents library?",
            
            # Path-based queries
            "Show me documents in /Shared Documents/",
            "List files in the root folder.",
            
            # Search-based queries
            "Search for documents containing 'test' in the name.",
            "Find any Word documents (.docx files).",
            
            # General queries
            "What SharePoint content can you access?",
            "Show me the available document libraries.",
        ]
        
        for i, query in enumerate(test_queries, 1):
            print(f"\n--- Test {i}: {query} ---")
            
            try:
                # Create agent with SharePoint tool
                agent = ai_project_client.agents.create_agent(
                    model=model_deployment_name,
                    name=f"SharePoint Test Agent {i}",
                    instructions="You are a helpful assistant that can access SharePoint documents. Be specific about what you can and cannot access.",
                    tools=[sharepoint_tool]
                )
                
                # Create thread and run
                thread = ai_project_client.agents.create_thread()
                message = ai_project_client.agents.create_message(
                    thread_id=thread.id,
                    role="user",
                    content=query
                )
                
                run = ai_project_client.agents.create_and_process_run(
                    thread_id=thread.id,
                    assistant_id=agent.id
                )
                
                print(f"Status: {run.status}")
                
                if run.status == "completed":
                    messages = ai_project_client.agents.list_messages(thread_id=thread.id)
                    if hasattr(messages, 'data') and messages.data:
                        for msg in messages.data:
                            if msg.role == "assistant":
                                for content in msg.content:
                                    if content.type == "text":
                                        print(f"✅ Response: {content.text.value[:200]}...")
                                        break
                                break
                    else:
                        print("❌ No response data available")
                elif run.status == "failed":
                    print(f"❌ Failed: {run.last_error}")
                else:
                    print(f"⚠️ Incomplete status: {run.status}")
                
                # Cleanup
                ai_project_client.agents.delete_agent(agent.id)
                
            except Exception as e:
                print(f"❌ Exception: {str(e)}")
                
    except Exception as e:
        print(f"❌ Setup failed: {str(e)}")
    
    print("\n✅ Document access testing completed")

if __name__ == "__main__":
    test_document_access()
