# SharePoint Connection Setup in AI Foundry

## Overview
Create SharePoint connection in Azure AI Foundry with proper authentication.

## Step 1: Navigate to Connections

1. **Open AI Foundry Project**
   - Go to your AI Foundry project
   - Navigate to "Connections" section

2. **Create New Connection**
   - Click "New connection"
   - Select "SharePoint" from connection types

## Step 2: Configure Connection

1. **Connection Details**
   - Connection name: `sharepoint-documents-connection`
   - Description: "SharePoint document access for AI agents"

2. **Authentication Settings**
   - Authentication type: "Custom Keys"
   - Client ID: [From app registration]
   - Client Secret: [From app registration]
   - Tenant ID: [From app registration]

3. **SharePoint Configuration**
   - Site URL: `https://your-tenant.sharepoint.com/sites/your-site`
   - Validate the URL is accessible

## Step 3: Test Connection

1. **Validate Setup**
   - Click "Test connection"
   - Verify success message

2. **Get Connection ID**
   - After creation, copy the full connection ID
   - Format: `/subscriptions/{sub-id}/resourceGroups/{rg}/providers/Microsoft.CognitiveServices/accounts/{account}/projects/{project}/connections/{connection-name}`

## Required for .env File

After setup, you'll have:
- `SHAREPOINT_CONNECTION_ID`: Full connection resource ID

Next: [Environment Setup](environment-setup.md)
