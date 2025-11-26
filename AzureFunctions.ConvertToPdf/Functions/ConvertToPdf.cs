using System.Net;
using System.Text.Json;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using System.Text.Json.Serialization;
using System.Net.Http.Headers;
using System.Net.Http.Json;

namespace AzureFunctions.ConvertToPdf
{
    public class ConvertToPdf
    {
        private readonly IPnPContextFactory _pnpContextFactory;
        private readonly ILogger<ConvertToPdf> _logger;
        private readonly IHttpClientFactory _httpClientFactory;

        public ConvertToPdf(IPnPContextFactory pnpContextFactory, ILogger<ConvertToPdf> logger, IHttpClientFactory httpClientFactory)
        {
            _pnpContextFactory = pnpContextFactory;
            _logger = logger;
            _httpClientFactory = httpClientFactory;
        }

        // Response DTO
        public record ConvertResponse(
            bool Success,
            string? PdfItemWebUrl,
            string? Message
        );

        [Function("ConvertToPdf")]
        public async Task<HttpResponseData> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)]
            HttpRequestData req)
        {
            try
            {
                // Parse your payload (site config name, library titles, item id, etc.)
                var request = await JsonSerializer.DeserializeAsync<ConvertRequest>(req.Body, new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                });

                if (request == null)
                {
                    return await CreateResponseAsync(
                        req,
                        HttpStatusCode.BadRequest,
                        new ConvertResponse(false, null, "Invalid request body.")
                    );
                }

                _logger.LogInformation($"ConvertToPdf called. SourceSiteUrl={request.SourceSiteUrl}, " +
                    $"SourceLibrary={request.SourceLibraryTitle}, ItemId={request.ItemId}, " +
                    $"TargetLibrary={request.TargetLibraryTitle}, TargetFolder={request.TargetFolderPathInLibrary}",
                    request.SourceSiteUrl,
                    request.SourceLibraryTitle,
                    request.ItemId,
                    request.TargetLibraryTitle,
                    request.TargetFolderPathInLibrary);


                using var defaultCtx = await _pnpContextFactory.CreateAsync("Default");
                using var context = await _pnpContextFactory.CreateAsync(new Uri(request.SourceSiteUrl), defaultCtx.AuthenticationProvider);

                // Get source list + item to build a nice PDF file name
                IList sourceList = await context.Web.Lists
                    .GetByTitleAsync(request.SourceLibraryTitle, l => l.Id);

                var sourceItem = sourceList.Items.GetById(request.ItemId);

                await sourceItem.LoadAsync(i => i.All);

                var sourceName = request.FileName; // sourceItem["FileLeafRef"]?.ToString() ?? "document.docx";
                var pdfFileName = Path.GetFileNameWithoutExtension(sourceName) + ".pdf";

                _logger.LogInformation("Source file name resolved to {SourceName}, will create PDF as {PdfFileName}", sourceName, pdfFileName);

                // Convert DOCX to PDF bytes via Graph
                var pdfBytes = await ConvertWordItemToPdfAsync(
                    context,
                    _httpClientFactory,
                    request.SourceLibraryTitle,
                    request.ItemId);

                _logger.LogInformation("PDF conversion done, size={Size} bytes", pdfBytes.Length);

                // Upload PDF bytes to target library (and optional folder) via Graph
                var pdfWebUrl = await UploadPdfToTargetLibraryAsync(
                    context,
                    _httpClientFactory,
                    request.TargetLibraryTitle,
                    request.TargetFolderPathInLibrary,
                    pdfFileName,
                    pdfBytes);

                if (string.IsNullOrEmpty(pdfWebUrl))
                {
                    return await CreateResponseAsync(
                        req,
                        HttpStatusCode.OK,
                        new ConvertResponse(true, null, "PDF created but webUrl could not be resolved.")
                    );
                }

                _logger.LogInformation("PDF uploaded successfully. WebUrl={WebUrl}", pdfWebUrl);

                return await CreateResponseAsync(
                    req,
                    HttpStatusCode.OK,
                    new ConvertResponse(true, pdfWebUrl, "Conversion completed successfully.")
                );
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error in ConvertToPdf.");
                return await CreateResponseAsync(
                    req,
                    HttpStatusCode.InternalServerError,
                    new ConvertResponse(false, null, $"Error: {ex.Message}")
                );
            }
        }

        private static async Task<HttpResponseData> CreateResponseAsync(
            HttpRequestData req,
            HttpStatusCode statusCode,
            ConvertResponse payload)
        {
            var response = req.CreateResponse(statusCode);
            response.Headers.Add("Content-Type", "application/json; charset=utf-8");
            await response.WriteStringAsync(JsonSerializer.Serialize(payload));
            return response;
        }

        /// <summary>
        /// Creates a HttpClient for Graph with a Bearer token taken from PnP.Core's AuthenticationProvider
        /// </summary>
        private static async Task<HttpClient> CreateGraphClientAsync(PnPContext context, IHttpClientFactory httpClientFactory)
        {
            var client = httpClientFactory.CreateClient();
            client.BaseAddress = new Uri("https://graph.microsoft.com/v1.0/");

            // Ask PnP.Core to authenticate the request (it will add the Bearer token)
            var dummyRequest = new HttpRequestMessage(HttpMethod.Get, "https://graph.microsoft.com/v1.0/$metadata");
            await context.AuthenticationProvider.AuthenticateRequestAsync(dummyRequest.RequestUri, dummyRequest);

            // Copy the Authorization header to our client so we can reuse it
            if (dummyRequest.Headers.Authorization is AuthenticationHeaderValue auth)
            {
                client.DefaultRequestHeaders.Authorization = auth;
            }

            return client;
        }

        /// <summary>
        /// Resolves the Graph siteId (GUID-like string) for the current PnPContext site
        /// </summary>
        private static async Task<string> GetGraphSiteIdAsync(PnPContext context, HttpClient graphClient)
        {
            var uri = context.Uri; // e.g. https://tenant.sharepoint.com/sites/MySite
            var host = uri.Host;   // tenant.sharepoint.com
            var path = uri.AbsolutePath; // /sites/MySite

            var response = await graphClient.GetAsync($"sites/{host}:{path}?$select=id");
            response.EnsureSuccessStatusCode();

            var site = await response.Content.ReadFromJsonAsync<GraphSite>();
            if (site == null || string.IsNullOrEmpty(site.Id))
                throw new InvalidOperationException("Could not resolve Graph site id.");

            return site.Id;
        }

        /// <summary>
        /// Uses Graph to convert a list item (Word) to PDF bytes via /drives/{driveId}/items/{itemId}/content?format=pdf
        /// </summary>
        private static async Task<byte[]> ConvertWordItemToPdfAsync(
            PnPContext context,
            IHttpClientFactory httpClientFactory,
            string sourceLibraryTitle,
            int listItemId)
        {
            // 1) Get list Id (GUID) via PnP.Core (no Drive involved)
            IList sourceList = await context.Web.Lists
                .GetByTitleAsync(sourceLibraryTitle, l => l.Id);

            var listId = sourceList.Id; // Guid

            // 2) Get a Graph client with auth from PnP.Core
            var graphClient = await CreateGraphClientAsync(context, httpClientFactory);

            // 3) Resolve the Graph siteId for this site
            var siteId = await GetGraphSiteIdAsync(context, graphClient);

            // 4) Get the DriveItem (and driveId) for the given list item
            var driveItemResponse = await graphClient.GetAsync(
                $"sites/{siteId}/lists/{listId}/items/{listItemId}/driveItem?$select=id,parentReference");

            driveItemResponse.EnsureSuccessStatusCode();

            var driveItem = await driveItemResponse.Content.ReadFromJsonAsync<GraphDriveItemInfo>();
            if (driveItem == null || driveItem.ParentReference == null)
                throw new InvalidOperationException("Could not resolve DriveItem for list item.");

            var driveId = driveItem.ParentReference.DriveId;
            var driveItemId = driveItem.Id;

            // 5) Call the Graph "convert to PDF" endpoint
            var pdfResponse = await graphClient.GetAsync(
                $"drives/{driveId}/items/{driveItemId}/content?format=pdf");

            pdfResponse.EnsureSuccessStatusCode();
            return await pdfResponse.Content.ReadAsByteArrayAsync();
        }

        /// <summary>
        /// Uploads the given PDF bytes to the specified target library (and optional folder) using Graph.
        /// Returns the webUrl of the created PDF item.
        /// </summary>
        private static async Task<string?> UploadPdfToTargetLibraryAsync(
            PnPContext context,
            IHttpClientFactory httpClientFactory,
            string targetLibraryTitle,
            string? targetFolderPathInLibrary,
            string pdfFileName,
            byte[] pdfBytes)
        {
            // 1) Get target list (library) to know its GUID
            IList targetList = await context.Web.Lists
                .GetByTitleAsync(targetLibraryTitle, l => l.Id);

            var listId = targetList.Id; // Guid

            // 2) Build an authenticated Graph client
            var graphClient = await CreateGraphClientAsync(context, httpClientFactory);

            // 3) Resolve Graph siteId for this site
            var siteId = await GetGraphSiteIdAsync(context, graphClient);

            // 4) Get the drive associated with this list
            //    GET /sites/{siteId}/lists/{listId}/drive?$select=id
            var driveResponse = await graphClient.GetAsync(
                $"sites/{siteId}/lists/{listId}/drive?$select=id");

            driveResponse.EnsureSuccessStatusCode();

            var drive = await driveResponse.Content.ReadFromJsonAsync<GraphDrive>();
            if (drive == null || string.IsNullOrEmpty(drive.Id))
                throw new InvalidOperationException("Could not resolve drive for target library.");

            var driveId = drive.Id;

            // 5) Build the path inside the drive
            var relativePath = string.IsNullOrWhiteSpace(targetFolderPathInLibrary)
                ? pdfFileName
                : $"{targetFolderPathInLibrary.TrimEnd('/')}/{pdfFileName}";

            var escapedPath = Uri.EscapeDataString(relativePath);

            // 6) Upload to: PUT /drives/{driveId}/root:/{path}:/content
            var uploadUrl = $"drives/{driveId}/root:/{escapedPath}:/content";

            using var contentBytes = new ByteArrayContent(pdfBytes);
            contentBytes.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");

            var uploadResponse = await graphClient.PutAsync(uploadUrl, contentBytes);
            uploadResponse.EnsureSuccessStatusCode();

            var json = await uploadResponse.Content.ReadAsStringAsync();

            // 7) Extract webUrl from the response
            string? webUrl = null;
            try
            {
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("webUrl", out var webUrlProp))
                {
                    webUrl = webUrlProp.GetString();
                }
            }
            catch
            {
                // ignore, webUrl stays null
            }

            return webUrl;
        }
    }

    // Helper POCOs for Graph responses

    public sealed class GraphSite
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;
    }

    public sealed class GraphDriveItemInfo
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;

        [JsonPropertyName("parentReference")]
        public GraphParentReference ParentReference { get; set; } = new();
    }

    public sealed class GraphParentReference
    {
        [JsonPropertyName("driveId")]
        public string DriveId { get; set; } = string.Empty;
    }

    public sealed class GraphDrive
    {
        [JsonPropertyName("id")]
        public string Id { get; set; } = string.Empty;
    }

    public sealed class ConvertRequest
    {
        public string SourceSiteUrl { get; set; } = string.Empty;
        public string SourceLibraryTitle { get; set; } = string.Empty;
        public int ItemId { get; set; }
        public string FileName { get; set; } = string.Empty;

        // New: where to store the PDF
        public string TargetLibraryTitle { get; set; } = string.Empty;
        public string? TargetFolderPathInLibrary { get; set; }
    }
}
