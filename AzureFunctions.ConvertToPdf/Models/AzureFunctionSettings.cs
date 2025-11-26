using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AzureFunctions.ConvertToPdf.Models
{
    public class AzureFunctionSettings
    {
        public string ClientId { get; set; }
        public string TenantId { get; set; }
        public string SiteUrl { get; set; }

        public string TenantDomain { get; set; }

        public string CertificateThumbprint { get; set; }

        public string BaseUrl { get; set; }
        public string AppCertificate { get; set; }
        public string AppCertificatePassword { get; set; }

        public string AzureWebJobsStorage { get; set; }

        public string KeyVaultUrl { get; set; }
        public string KeyVaultCertName { get; set; }

    }
}
