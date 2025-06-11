# Environment Setup Guide

## Getting Required Environment Variables

This guide helps you find all the values needed for your `.env` file.

## Step 1: Create .env File

First, copy the example file:
```bash
cp .env.example .env
```

Then edit the `.env` file with your specific values.

## PROJECT_ENDPOINT

**Where to find:**
1. Open Azure AI Foundry portal
2. Navigate to your project
3. Go to "Overview" page
4. Copy the "Project endpoint" URL

**Format:**
```
PROJECT_ENDPOINT=https://your-project.services.ai.azure.com/api/projects/your-project-name
```

## MODEL_DEPLOYMENT_NAME

**Where to find:**
1. In AI Foundry project, go to "Models + endpoints"
2. Find your deployed model (e.g., gpt-4o)
3. Copy the deployment name from "Name" column

**Format:**
```
MODEL_DEPLOYMENT_NAME=gpt-4o
```

## SHAREPOINT_CONNECTION_ID

**Where to find:**
1. In AI Foundry project, go to "Connections"
2. Find your SharePoint connection
3. Click on connection name
4. Copy the full Resource ID

**Format:**
```
SHAREPOINT_CONNECTION_ID=/subscriptions/{subscription-id}/resourceGroups/{resource-group}/providers/Microsoft.CognitiveServices/accounts/{account}/projects/{project}/connections/{connection-name}
```

## DOCUMENT_NAME

**Where to find:**
1. Go to your SharePoint site
2. Navigate to document library
3. Note the exact filename including extension

**Format:**
```
DOCUMENT_NAME=your-document.docx
```

## Complete .env File Example

```env
# Azure AI Foundry Configuration
PROJECT_ENDPOINT=https://my-project.services.ai.azure.com/api/projects/SharePoint-AI-Project
MODEL_DEPLOYMENT_NAME=gpt-4o
SHAREPOINT_CONNECTION_ID=/subscriptions/12345678-1234-1234-1234-123456789012/resourceGroups/myResourceGroup/providers/Microsoft.CognitiveServices/accounts/myAIAccount/projects/SharePoint-AI-Project/connections/sharepoint-documents-connection
DOCUMENT_NAME=test-document.docx
```

## Validation

After creating your `.env` file, run the diagnostic script:
```bash
python testing/diagnostic_sharepoint.py
```

This will validate all your environment variables and connections.
