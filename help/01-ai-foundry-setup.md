# AI Foundry Project Setup Guide

## Overview
This guide walks you through creating an Azure AI Foundry project with screenshots and detailed steps.

## Prerequisites
- Azure subscription with AI Foundry access
- Azure CLI installed and authenticated

## Step 1: Create Azure AI Foundry Project

1. **Navigate to Azure AI Foundry Portal**
   - Go to [https://ai.azure.com](https://ai.azure.com)
   - Sign in with your Azure credentials

2. **Create New Project**
   - Click "Create new project"
   - Choose your subscription and resource group
   - Enter project name (e.g., "SharePoint-AI-Project")
   - Select region

3. **Deploy GPT-4o Model**
   - Navigate to "Models + endpoints"
   - Click "Deploy model"
   - Select "gpt-4o"
   - Configure deployment settings
   - Note the deployment name for your `.env` file

4. **Get Project Endpoint**
   - Go to project Overview page
   - Copy the project endpoint URL
   - Format: `https://your-project.services.ai.azure.com/api/projects/your-project`

## Step 2: Note Required Values

After setup, collect these values for your `.env` file:
- `PROJECT_ENDPOINT`: From project Overview page
- `MODEL_DEPLOYMENT_NAME`: From Models + endpoints page

Next: [SharePoint App Registration](02-sharepoint-app-registration.md)
