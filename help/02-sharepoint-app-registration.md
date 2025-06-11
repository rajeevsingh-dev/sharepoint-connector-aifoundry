# SharePoint App Registration Setup

## Overview
Configure SharePoint app registration with proper Microsoft Graph API permissions.

## Step 1: Create App Registration

1. **Navigate to Azure Portal**
   - Go to [https://portal.azure.com](https://portal.azure.com)
   - Search for "App registrations"

2. **Create New Registration**
   - Click "New registration"
   - Name: "SharePoint-AI-Foundry-Access"
   - Account types: "Accounts in this organizational directory only"
   - Redirect URI: Leave blank
   - Click "Register"

3. **Note Application Details**
   - Copy Application (client) ID
   - Copy Directory (tenant) ID

## Step 2: Configure API Permissions

1. **Add Microsoft Graph Permissions**
   - Navigate to "API permissions"
   - Click "Add a permission"
   - Select "Microsoft Graph"
   - Choose "Application permissions"

2. **Required Permissions**
   Add these permissions:
   - `Sites.Read.All` - Read items in all site collections
   - `Files.Read.All` - Read files in all site collections
   - `User.Read.All` - Read all users' full profiles

3. **Grant Admin Consent**
   - Click "Grant admin consent for [tenant]"
   - Confirm the action

## Step 3: Create Client Secret

1. **Navigate to Certificates & secrets**
2. **Create New Secret**
   - Click "New client secret"
   - Description: "SharePoint AI Access"
   - Expires: 24 months
   - Click "Add"
3. **Copy Secret Value** (save immediately - won't be shown again)

## Required Values for SharePoint Connection

Collect these values:
- Application (client) ID
- Directory (tenant) ID  
- Client secret value
- SharePoint site URL

Next: [SharePoint Connection Setup](03-sharepoint-connection-setup.md)
