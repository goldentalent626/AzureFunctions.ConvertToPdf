# AzureFunctions.ConvertToPdf

## Introduction
AzureFunctions.ConvertToPdf is a .NET 8 Azure Function that automates the conversion of Word documents stored in SharePoint Online document libraries to PDF format using Microsoft Graph API. The converted PDFs are then uploaded to a specified target library or folder within SharePoint. This solution is ideal for organizations looking to streamline document workflows and leverage serverless cloud automation.

## Getting Started

### Installation
1. Clone the repository to your local machine.
2. Ensure you have the .NET 8 SDK installed.
3. Configure your Azure environment with the necessary resources:
   - Azure Function App
   - Azure Key Vault (for certificate storage)
   - Azure Storage Account

### Software Dependencies
- .NET 8
- Microsoft.Azure.Functions.Worker
- PnP.Core SDK
- Microsoft.Identity.Web
- Microsoft.Extensions.Azure
- Application Insights (for telemetry)

### Configuration
Update the `AzureFunctionSettings` section in your configuration (e.g., `local.settings.json` or environment variables) with:
- `ClientId`, `TenantId`, `SiteUrl`, `TenantDomain`
- Certificate details: `CertificateThumbprint`, `KeyVaultUrl`, `KeyVaultCertName`
- `AzureWebJobsStorage` connection string

### API Reference

#### ConvertToPdf Function
**Endpoint:** `POST /api/ConvertToPdf`

**Request Body:**