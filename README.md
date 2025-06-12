# SharePoint Connector with Azure AI Foundry

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Azure AI](https://img.shields.io/badge/Azure-AI%20Foundry-blue.svg)](https://azure.microsoft.com/en-us/products/ai-foundry/)

## 📋 Overview

This project demonstrates how to integrate SharePoint documents with Azure AI Foundry using the new **SharePoint Connector** tool for AI-powered document analysis, announced in Microsoft's latest AI Foundry updates.

**Recent Announcement**: Microsoft has introduced SharePoint Connector as a preview feature in Azure AI Foundry. Learn more: [SharePoint Tools Documentation](https://learn.microsoft.com/en-us/azure/ai-services/agents/how-to/tools/sharepoint)

## 🔥 Why Use SharePoint Connector vs Microsoft Graph API?

| Feature | SharePoint Connector | Microsoft Graph API |
|---------|---------------------|---------------------|
| **AI Integration** | ✅ Native AI agent integration | ❌ Manual integration required |
| **Natural Language** | ✅ Ask questions in plain English | ❌ Complex query building |
| **Document Analysis** | ✅ Built-in content understanding | ❌ Raw file access only |
| **Complexity** | ✅ Simple agent setup | ❌ Complex authentication & parsing |
| **Maintenance** | ✅ Microsoft-managed | ❌ Custom code maintenance |

## 📋 Prerequisites

- Azure subscription with AI Foundry access
- SharePoint Online site with documents  
- Python 3.8+ installed
- Azure CLI installed
- **Microsoft 365 Copilot license** (Required for SharePoint tool preview)

## 🔧 Detailed Setup Prerequisites

For complete step-by-step setup with screenshots and detailed instructions:

📖 **[AI Foundry Project Setup Guide](help/01-ai-foundry-setup.md)** - Create Azure AI Foundry project with screenshots

📋 **[SharePoint App Registration Guide](help/02-sharepoint-app-registration.md)** - Configure SharePoint app with permissions

🔗 **[SharePoint Connection Setup](help/03-sharepoint-connection-setup.md)** - Create SharePoint connection in AI Foundry

## 🚀 Quick Start

### Option 1: Validate SharePoint Connection in AI Foundry Portal

Test your SharePoint connection directly in the Azure AI Foundry web interface without writing code. This option provides a visual way to verify your setup and test document access through the portal's chat interface.

**📖 [Portal Testing Guide with Screenshots](help/portal-testing-guide.md)**

### Option 2: Using Python SDK Code

Build custom applications using the Azure AI Python SDK to programmatically access SharePoint documents through AI agents. This option provides full control and integration capabilities for enterprise applications.

## ⚡ Installation & Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/sharepoint-ai-foundry-demo.git
   cd sharepoint-ai-foundry-demo
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure environment variables**:
   Copy the example file and add your specific values:
   ```bash
   cp .env.example .env
   ```
   
   Edit `.env` with your Azure AI Foundry project details:
   ```env
   PROJECT_ENDPOINT=https://your-project.services.ai.azure.com/api/projects/your-project
   SHAREPOINT_CONNECTION_ID=/subscriptions/{subscription-id}/resourceGroups/{resource-group}/providers/Microsoft.CognitiveServices/accounts/{account}/projects/{project}/connections/{connection-name}
   MODEL_DEPLOYMENT_NAME=gpt-4o
   DOCUMENT_NAME=your-test-document.docx
   ```

   **📖 [How to Get Environment Variables](help/environment-setup.md)** - Detailed guide to find all required values

4. **Authenticate with Azure**:
   ```bash
   az login
   az account set --subscription "your-subscription-id"
   ```

## 🧪 Code Examples

### Code 1: Basic SharePoint AI Agent

Run the main application to test SharePoint document access:

```bash
python sample_agents_sharepoint.py
```

**Expected Output:**
```
=== SharePoint AI Agent Demo ===
✅ Created AI Project Client
✅ Created SharePoint Tool
✅ Created agent: asst_abc123
✅ Created thread: thread_xyz789
📝 Processing query: "Summarize the key points of the document"

Agent Response: "Based on the SharePoint document analysis..."
✅ Test completed successfully
```

### Code 2: Connection Validation & Troubleshooting

Run diagnostic tests to validate all connections:

```bash
python testing/diagnostic_sharepoint.py
```

**Expected Output:**
```
=== SharePoint Connection Diagnostic ===
✅ Azure Authentication: Working
✅ AI Project Client: Working  
✅ SharePoint Tool Creation: Working
✅ SharePoint Connection: Valid
✅ Document Library Access: Working
🎉 All systems operational!
```

## 🧪 Testing

Run the diagnostic script to validate your setup:

```bash
python testing/diagnostic_sharepoint.py
```

This will test all connections and show the status of each component.

## 📁 Project Structure

```
sharepoint-ai-foundry-demo/
├── 📄 README.md                    # This file
├── 📄 requirements.txt             # Python dependencies
├── 📄 sample_agents_sharepoint.py  # Main application
├── 📄 .env.example                 # Environment template (copy to .env)
├── 📄 LICENSE                      # MIT License
├── 📁 help/                        # Detailed documentation
│   ├── 01-ai-foundry-setup.md
│   ├── 02-sharepoint-app-registration.md
│   ├── 03-sharepoint-connection-setup.md
│   ├── portal-testing-guide.md
│   └── environment-setup.md
└── 📁 testing/                     # Test scripts
    └── diagnostic_sharepoint.py    # Connection diagnostics
```

For complete step-by-step setup with screenshots and detailed instructions:

1. **[Azure AI Foundry Project Setup](help/01-ai-foundry-setup.md)** - Create AI Foundry project, deploy models, configure endpoints
2. **[SharePoint App Registration](help/02-sharepoint-app-registration.md)** - Register app in Azure AD, configure Graph API permissions
3. **[SharePoint Connection Setup](help/03-sharepoint-connection.md)** - Configure SharePoint connector in AI Foundry with authentication





## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.


**Built with ❤️ for the Azure AI community**
