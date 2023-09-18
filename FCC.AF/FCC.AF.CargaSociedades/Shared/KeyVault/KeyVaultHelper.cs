using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using Microsoft.Extensions.Logging;
using System;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace FCC.AF.CargaSociedades.Shared.KeyVault
{
    public static class KeyVaultHelper
    {
        /// <summary>
        /// Funcion que obtiene certificado del keyvault
        /// </summary>
        /// <param name="URI_Vault"></param>
        /// <param name="NombreCertificado"></param>
        /// <returns></returns>
        public static async Task<X509Certificate2> GetCertificateAsync(string URI_Vault, string NombreCertificado, string TenantId, string ClientId, string ClientSecret, ILogger log)
        {
            X509Certificate2 certificate = null;
            try
            {
                var kv = new SecretClient(new Uri(URI_Vault), new ClientSecretCredential(TenantId, ClientId, ClientSecret));

                log.LogInformation("Se va a obtener el certificado");
                log.LogInformation("KV Url: " + URI_Vault);
                log.LogInformation("Certificado: " + NombreCertificado);
                KeyVaultSecret secret = await kv.GetSecretAsync(NombreCertificado);
                log.LogInformation("Secret Id: " + secret.Id);
                log.LogInformation("Secret Value: " + secret.Value.Substring(0,30) + "...");
                log.LogInformation("Certificado cargado correctamente");
                var privateKeyBytes = Convert.FromBase64String(secret.Value);

                certificate = new X509Certificate2(privateKeyBytes, (string)null, X509KeyStorageFlags.MachineKeySet | X509KeyStorageFlags.PersistKeySet | X509KeyStorageFlags.Exportable);

                log.LogInformation("Certificado correcto");
            }
            catch (Exception e)
            {
                log.LogError("Error en el certificado: " + e.Message);
            }
            return certificate;
        }

        //public static async Task<string> GetToken(string authority, string resource, string scope)

        //{
        //    ClientCredential credential = new ClientCredential(Environment.GetEnvironmentVariable("_ClientId"), Environment.GetEnvironmentVariable("_ClientSecret"));
        //    var context = new AuthenticationContext(authority, TokenCache.DefaultShared);
        //    var result = await context.AcquireTokenAsync(resource, credential);
        //    return result.AccessToken;
        //}
    }
}
