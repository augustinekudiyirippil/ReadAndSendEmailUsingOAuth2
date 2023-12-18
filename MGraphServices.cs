using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace Simplisys_OAuth2_Service
{
    class MGraphServices
    {
        [Obsolete]
        public async Task<GraphServiceClient> connectToMailBox(string clientID, string tenantID, string emailAddress, string emailPassword)
        {

            GraphServiceClient graphServiceClient = await CreateGraphClientService(new PublicClientApplicationOptions
            {
                ClientId = clientID,
                TenantId = tenantID
            }, emailAddress, emailPassword);



            return graphServiceClient;

        }

        [Obsolete]
        private  async  Task<GraphServiceClient> CreateGraphClientService(PublicClientApplicationOptions _PublicClientApplicationOptions, string _EmailId, string _Password)
        {
            try
            {
                string[] scopes = new string[] { "user.read" };

                var pca = PublicClientApplicationBuilder.CreateWithApplicationOptions(_PublicClientApplicationOptions).WithAuthority(AzureCloudInstance.AzurePublic, _PublicClientApplicationOptions.TenantId).Build();



                var authResult = await pca.AcquireTokenByUsernamePassword(new string[] { "https://graph.microsoft.com/.default" }, _EmailId, new NetworkCredential("", _Password).SecurePassword).ExecuteAsync();

                 

                return  new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) => { requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer",  authResult.AccessToken); }));

                //return new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) => { requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authResult.AccessToken); return Task.CompletedTask; }));
            }
            catch (Exception ex)
            {
                string exc = ex.Message.ToString();
                throw;
            }
        }


    }
}
