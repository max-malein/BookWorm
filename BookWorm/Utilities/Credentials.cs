using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Util.Store;
using Grasshopper.Kernel;
using System;
using System.IO;
using System.Threading;

namespace BookWorm.Utilities
{
    /// <summary>
    /// 
    /// </summary>
    public class Credentials
    {
        /// <summary>
        /// 
        /// </summary>
        public static SheetsService Service => SetSheetService();

        private static UserCredential Credential => SetCredentials();

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        private static string[] scopes = { SheetsService.Scope.Spreadsheets };
        private static string applicationName = "Bookworm";

        private static UserCredential SetCredentials()
        {
            GH_AssemblyInfo info = Grasshopper.Instances.ComponentServer.FindAssembly(new Guid("56dfe1a3-4e7b-425f-b169-965c0d1f7977"));
            string assemblyLocation = Path.GetDirectoryName(info.Location);

            UserCredential credential;

            string credPath = Path.Combine(assemblyLocation, "token.json");

            var sec = new ClientSecrets()
            {
                ClientId = "631922160569-4klgvfht5ngpihjr7l2ovs4enp8ved8i.apps.googleusercontent.com",
                ClientSecret = "Iu8Y5q82pLMXvN4y_6qPK9TP",
            };

            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                sec,
                scopes,
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true)).Result;

            return credential;
        }

        private static SheetsService SetSheetService()
        {
            // Create Google Sheets API service.
            return new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = Credential,
                ApplicationName = applicationName,
            });
        }
    }
}
