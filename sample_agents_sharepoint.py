# ------------------------------------
# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.
# ------------------------------------

"""
SharePoint AI Agent Demo

This script demonstrates how to create an AI agent that can access and analyze
SharePoint documents using Azure AI Foundry and the SharePoint Connector tool.

Prerequisites:
- Azure AI Foundry project with SharePoint connection configured
- Environment variables set in .env file
- SharePoint Online site with documents

Usage:
    python sample_agents_sharepoint.py

Note: SharePoint tool requires preview access. Contact azureagents-preview@microsoft.com
"""

import os
from azure.ai.projects import AIProjectClient
from azure.identity import DefaultAzureCredential
from azure.ai.agents.models import SharepointTool
from dotenv import load_dotenv
from azure.core.exceptions import HttpResponseError
import json

# Load environment variables from .env file
load_dotenv()

print("=== SharePoint AI Agent Demo ===")
print("Testing SharePoint document access with Azure AI Foundry")
print()

# Verify environment variables are loaded
required_vars = ["PROJECT_ENDPOINT", "MODEL_DEPLOYMENT_NAME", "SHAREPOINT_CONNECTION_ID"]
missing_vars = [var for var in required_vars if not os.environ.get(var)]

if missing_vars:
    print(f"❌ Missing required environment variables: {missing_vars}")
    print("Please check your .env file and ensure all required variables are set.")
    exit(1)

print("✅ Environment variables loaded successfully")
print(f"📁 Testing with document: {os.environ.get('DOCUMENT_NAME', 'default-document.docx')}")
print()

# Create an Azure AI Client from a connection string, copied from your AI Studio project.
# At the moment, it should be in the format "<HostName>;<AzureSubscriptionId>;<ResourceGroup>;<HubName>"
# Customer needs to login to Azure subscription via Azure CLI and set the environment variables

try:
    project_client = AIProjectClient(
        endpoint=os.environ["PROJECT_ENDPOINT"],
        credential=DefaultAzureCredential(),
    )
    print("✅ Successfully created AI Project Client")
except Exception as e:
    print(f"❌ Error creating AI Project Client: {e}")
    exit(1)

# Initialize SharePoint tool with connection id
try:
    sharepoint = SharepointTool(connection_id=os.environ["SHAREPOINT_CONNECTION_ID"])
    print("✅ Successfully created SharePoint Tool")
except Exception as e:
    print(f"❌ Error creating SharePoint Tool: {e}")
    exit(1)

# Create agent with Sharepoint tool and process agent run
with project_client:
    agents_client = project_client.agents
    
    try:
        agent = agents_client.create_agent(
            model=os.environ["MODEL_DEPLOYMENT_NAME"],
            name="sharepoint-ai-agent",
            instructions="You are a helpful SharePoint assistant that can access and analyze documents.",
            tools=sharepoint.definitions,
        )
        print(f"✅ Created agent: {agent.id}")
    except Exception as e:
        print(f"❌ Error creating agent: {e}")
        exit(1)

    # Create thread for communication
    thread = agents_client.threads.create()
    print(f"✅ Created thread: {thread.id}")

    # Create message to thread - Test with the document from env
    document_name = os.environ.get('DOCUMENT_NAME', 'your-document.docx')
    query_message = f"Please analyze and summarize the SharePoint document named '{document_name}'. If you can't access it, please list what documents are available in the SharePoint site."
    
    message = agents_client.messages.create(
        thread_id=thread.id,
        role="user",
        content=query_message,
    )
    print(f"✅ Created message: {message.id}")
    print(f"📝 Query: {query_message}")
    print()

    # Create and process agent run in thread with tools
    print("🔄 Processing agent run...")
    run = agents_client.runs.create_and_process(thread_id=thread.id, agent_id=agent.id)
    print(f"📊 Run finished with status: {run.status}")

    if run.status == "failed":
        print(f"❌ Run failed: {run.last_error}")
        print("\n💡 This is likely due to SharePoint tool requiring preview access.")
        print("📧 Contact azureagents-preview@microsoft.com for whitelist access.")
    elif run.status == "completed":
        print("✅ Run completed successfully!")
    else:
        print(f"⏳ Run status: {run.status}")

    # Fetch and log all messages
    print("\n=== Conversation Messages ===")
    messages = agents_client.messages.list(thread_id=thread.id)
    for msg in messages:
        role_icon = "👤" if msg.role == "user" else "🤖"
        if msg.content:
            for content in msg.content:
                if hasattr(content, 'text') and content.text:
                    print(f"{role_icon} {msg.role.title()}: {content.text.value}")
                    break

    # Clean up
    agents_client.delete_agent(agent.id)
    print(f"\n🧹 Cleaned up agent: {agent.id}")

print("\n=== Demo Complete ===")
print("For troubleshooting, run: python testing/diagnostic_sharepoint.py")