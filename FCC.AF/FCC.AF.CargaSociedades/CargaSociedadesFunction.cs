using System;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
 

namespace FCC.AF.CargaSociedades
{
    public class CargaSociedadesFunction
    {
        [FunctionName("CargaSociedadesFunction")]
        public async Task RunAsync([TimerTrigger("%ScheduleTriggerTime%"
        #if DEBUG
            , RunOnStartup=true
        #endif
            )]TimerInfo myTimer, ILogger log)
        {
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");

            X509Certificate2 certificate = await Shared.KeyVault.KeyVaultHelper.GetCertificateAsync(
                Environment.GetEnvironmentVariable("_CertURL"),
                Environment.GetEnvironmentVariable("_CertName"),
                Environment.GetEnvironmentVariable("_TenantId"),
                Environment.GetEnvironmentVariable("_ClientId"),
                Environment.GetEnvironmentVariable("_ClientSecret"),
                log);

            Application.CargaSociedades.DoProcess(certificate, log);

            log.LogInformation($"C# Timer trigger function finished at: {DateTime.Now}");
        }

    }
}
