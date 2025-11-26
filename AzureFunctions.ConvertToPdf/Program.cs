using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using PnP.Core.Auth.Services.Builder.Configuration;
using PnP.Core.Services.Builder.Configuration;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Azure.Functions.Worker.Builder;
using Microsoft.Identity.Web;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Azure;
using AzureFunctions.ConvertToPdf.Models;

var builder = FunctionsApplication.CreateBuilder(args);

builder.ConfigureFunctionsWebApplication();

// Bind configuration to POCO
var azureFunctionSettings = new AzureFunctionSettings();
builder.Configuration.Bind(azureFunctionSettings);

// Register config
builder.Services.AddSingleton(azureFunctionSettings);
builder.Services.AddSingleton<IConfiguration>(builder.Configuration);


// Add PnP Core SDK base
builder.Services.AddPnPCore(options =>
{
    options.PnPContext.GraphFirst = true;
    options.Sites.Add("Default", new PnPCoreSiteOptions
    {
        SiteUrl = azureFunctionSettings.SiteUrl
    });
});


// Add PnP Core SDK auth via certificate
builder.Services.AddPnPCoreAuthentication(options =>
{
    var cert = LoadCertificate(azureFunctionSettings);

    options.Credentials.Configurations.Add("CertAuth", new PnPCoreAuthenticationCredentialConfigurationOptions
    {
        ClientId = azureFunctionSettings.ClientId,
        TenantId = azureFunctionSettings.TenantId,
        X509Certificate = new PnPCoreAuthenticationX509CertificateOptions
        {
            Certificate = cert
        }
    });
    options.Sites.Add("Default", new PnPCoreAuthenticationSiteOptions
    {
        AuthenticationProviderName = "CertAuth"
    });
});


builder.Services.AddAzureClients(factory =>
{
    factory.AddBlobServiceClient(azureFunctionSettings.AzureWebJobsStorage);
    factory.AddQueueServiceClient(azureFunctionSettings.AzureWebJobsStorage);
    factory.AddTableServiceClient(azureFunctionSettings.AzureWebJobsStorage);
});


builder.Services.AddApplicationInsightsTelemetryWorkerService(); // telemetry pipeline
builder.Logging.AddApplicationInsights();


// demo token store (swap for TableChangeTokenStore from earlier when ready)
builder.Services.AddSingleton<IChangeTokenStore, InMemoryChangeTokenStore>();

builder.Services.AddHttpClient();

var app = builder.Build();
app.Run();


// Helper to load cert from Key Vault using DefaultCertificateLoader
static X509Certificate2 LoadCertificate(AzureFunctionSettings azureFunctionSettings)
{
    var kvUrl = azureFunctionSettings.KeyVaultUrl;
    var kvName = azureFunctionSettings.KeyVaultCertName;

    if (!string.IsNullOrEmpty(kvUrl) && !string.IsNullOrEmpty(kvName))
    {
        var desc = new CertificateDescription
        {
            SourceType = CertificateSource.KeyVault,
            KeyVaultUrl = kvUrl,
            KeyVaultCertificateName = kvName
        };
        new DefaultCertificateLoader().LoadIfNeeded(desc); // uses the Function App's Managed Identity
        X509Certificate2 cert = desc.Certificate!; // includes private key, no password needed
        return cert;
    }
    else // if we are running on localhost, load from cert store
    {
        var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
        var certificateCollection = store.Certificates.Find(
            X509FindType.FindByThumbprint,
            azureFunctionSettings.CertificateThumbprint,
            false
        );
        store.Close();
        return certificateCollection.First();
    }
}


// ===== Demo token store =====
public interface IChangeTokenStore
{
    Task<string?> GetAsync(string scopeKey, CancellationToken ct = default);
    Task SetAsync(string scopeKey, string token, CancellationToken ct = default);
}

public sealed class InMemoryChangeTokenStore : IChangeTokenStore
{
    private readonly Dictionary<string, string> _map = new();
    public Task<string?> GetAsync(string k, CancellationToken _ = default)
        => Task.FromResult(_map.TryGetValue(k, out var v) ? v : null);
    public Task SetAsync(string k, string t, CancellationToken _ = default)
    { _map[k] = t; return Task.CompletedTask; }
}


