# Portal Testing Guide

## Overview
Test SharePoint connection directly in Azure AI Foundry portal without code.

## Step 1: Create Agent in Portal

1. **Navigate to AI Foundry Portal**
   - Go to your AI Foundry project
   - Click "Agents" section

2. **Create New Agent**
   - Click "Create agent"
   - Name: "SharePoint Test Agent"
   - Instructions: "You are a helpful SharePoint assistant"
   - Model: Select your deployed model (e.g., gpt-4o)

3. **Add SharePoint Tool**
   - In "Tools" section, click "Add tool"
   - Select "SharePoint"
   - Choose your SharePoint connection
   - Save agent

## Step 2: Test in Chat Interface

1. **Start Chat Session**
   - Click "Test" or "Chat" button
   - Begin conversation with agent

2. **Test Queries**
   Try these queries:
   ```
   Hello, what SharePoint capabilities do you have?
   Can you list documents in our SharePoint site?
   Please summarize the content of [your-document-name]
   ```

## Expected Results

### If Whitelisted ✅
- Agent responds with document listings
- Can access and summarize document content
- SharePoint tool functions normally

### If Not Whitelisted ❌
- Agent responds generically about SharePoint
- Cannot access specific documents
- May get "Bad Request" errors for document operations

## Troubleshooting

**Connection Issues:**
- Verify SharePoint connection is "Connected" status
- Check app registration permissions
- Validate SharePoint site URL accessibility

**Preview Access:**
- Contact `azureagents-preview@microsoft.com` for whitelist access
- Include subscription ID and business justification

Next: [Python SDK Implementation](python-sdk-guide.md)
