using Microsoft.Kiota.Abstractions.Authentication;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OrganizationStructure.Generator
{
    public class TokenProvider : IAccessTokenProvider
    {
        private readonly string _token;

        public TokenProvider(string token)
        {
            _token = token;
        }

        public Task<string> GetAuthorizationTokenAsync(Uri uri, Dictionary<string, object> additionalAuthenticationContext = default,
            CancellationToken cancellationToken = default)
        {
            // get the token and return it in your own way
            return Task.FromResult(_token);
        }

        public AllowedHostsValidator AllowedHostsValidator { get; }
    }
}