using Azure.Identity;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions.Authentication;

namespace OrganizationStructure.Generator
{
    public static class GraphServiceClientFactory
    {
        public static GraphServiceClient CreateClientFromClientSecretCredential(string tenantId, string clientId, string clientSecret)
        {
            var clientSecretCredentials = new ClientSecretCredential(tenantId, clientId, clientSecret);
            return new GraphServiceClient(clientSecretCredentials);
        }

        public static GraphServiceClient CreateClientFromToken(string token)
        {
            var authProvider = new BaseBearerTokenAuthenticationProvider(new TokenProvider(token));
            return new GraphServiceClient(authProvider);
        }
    }
}