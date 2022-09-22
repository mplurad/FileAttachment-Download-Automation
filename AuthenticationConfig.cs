using Microsoft.Extensions.Configuration;
using Microsoft.Identity.Client;
using System.IO;
using System.Reflection;

namespace High_Radius_Invoice_Download_Automation
{
    /// <summary>
    /// Description of the configuration of an AzureAD public client application (desktop/mobile application). This should
    /// match the application registration done in the Azure portal
    /// </summary>
    static class AuthenticationConfig
    {
        /// <summary>
        /// Authentication options
        /// </summary>
        public static PublicClientApplicationOptions PublicClientApplicationOptions { get; set; }

        /// <summary>
        /// Base URL for Microsoft Graph (it varies depending on whether the application is ran
        /// in Microsoft Azure public clouds or national / sovereign clouds
        /// </summary>
        public static string MicrosoftGraphBaseEndpoint { get; set; }

        /// <summary>
        /// Reads the configuration from a json file
        /// </summary>
        /// <param name="path">Path to the configuration json file</param>
        /// <returns>AuthenticationConfig read from the json file</returns>
        public static void ReadFromJsonFile(string path)
        {
            // .NET configuration
            IConfigurationRoot Configuration;

            var builder = new ConfigurationBuilder().SetBasePath(Path.GetDirectoryName(Assembly.GetEntryAssembly().Location)).AddJsonFile(path);

            Configuration = builder.Build();

            // Read the auth and graph endpoint config
            PublicClientApplicationOptions = new PublicClientApplicationOptions();
            Configuration.Bind("Authentication", PublicClientApplicationOptions);
            MicrosoftGraphBaseEndpoint = Configuration.GetValue<string>("WebAPI:MicrosoftGraphBaseEndpoint");
        }
    }
}
