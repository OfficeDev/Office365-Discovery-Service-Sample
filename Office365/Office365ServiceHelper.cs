﻿//----------------------------------------------------------------------------------------------
//    Copyright 2014 Microsoft Corporation
//
//    Licensed under the Apache License, Version 2.0 (the "License");
//    you may not use this file except in compliance with the License.
//    You may obtain a copy of the License at
//
//      http://www.apache.org/licenses/LICENSE-2.0
//
//    Unless required by applicable law or agreed to in writing, software
//    distributed under the License is distributed on an "AS IS" BASIS,
//    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//    See the License for the specific language governing permissions and
//    limitations under the License.
//----------------------------------------------------------------------------------------------

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Windows.UI.Popups;

namespace Win8ServiceDiscovery
{
    public enum ServiceCapabilities
    {
        Mail,
        Calendar,
        Contacts,
        MyFiles
    }

    public static class Office365ServiceHelper
    {        
        static AuthenticationContext authContext = null;

        public static readonly string Authority = "https://login.windows.net/Common/";

        public static readonly string ClientId = App.Current.Resources["ida:ClientID"].ToString();

        public static readonly Uri DiscoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");

        public static readonly string DiscoveryServiceResourceId = "https://api.office.com/discovery/";
 
        public static async Task<AuthenticationResult> GetAccessTokenAsync(string serviceResourceId)
        {
            if (authContext == null)
            {
                authContext = new AuthenticationContext(Authority);

                var tokenCache = authContext.TokenCache.ReadItems().FirstOrDefault();

                if (tokenCache != null)
                {
                    authContext = new AuthenticationContext(tokenCache.Authority);
                }

            }

            authContext.UseCorporateNetwork = true;

            //
            // Every Windows Store application has a unique URI.
            // Windows ensures that only this application will receive messages sent to this URI.
            // ADAL uses this URI as the application's redirect URI to receive OAuth responses.
            // 
            // To determine this application's redirect URI, which is necessary when registering the app
            //      in AAD, set a breakpoint on the next line, run the app, and copy the string value of the URI.
            //      This is the only purposes of this line of code, it has no functional purpose in the application.
            //
            var redirectUri = WebAuthenticationBroker.GetCurrentApplicationCallbackUri();

            var authResult = await authContext.AcquireTokenAsync(serviceResourceId, ClientId, redirectUri);

            if (authResult.Status != AuthenticationStatus.Success)
            {
                if (authResult.Error == "authentication_canceled")
                {
                    // The user cancelled the sign-in, no need to display a message.
                }
                else
                {
                    MessageDialog dialog = new MessageDialog(string.Format("If the error continues, please contact your administrator.\n\nError: {0}\n\n Error Description:\n\n{1}", authResult.Error, authResult.ErrorDescription), "Sorry, an error occurred while signing you in.");
                    await dialog.ShowAsync();
                }
            }

            return authResult;
        }

        public static async Task<string> GetAccessTokenForResourceAsync(CapabilityDiscoveryResult discoveryCapabilityResult)
        {
            var authResult = await authContext.AcquireTokenSilentAsync(discoveryCapabilityResult.ServiceResourceId, ClientId);

            return authResult.AccessToken;
        }

        public static async Task<DiscoveryServiceCache> CreateAndSaveDiscoveryServiceCacheAsync()
        {
            DiscoveryServiceCache discoveryCache = null;

            string userId = null;

            var discoveryClient = new DiscoveryClient(DiscoveryServiceEndpointUri,
                        async () =>
                        {
                            var authResult = await GetAccessTokenAsync(DiscoveryServiceResourceId);

                            userId = authResult.UserInfo.UniqueId;

                            return authResult.AccessToken;
                        });

            var discoveryCapabilityResult = await discoveryClient.DiscoverCapabilitiesAsync();

            discoveryCache = await DiscoveryServiceCache.CreateAndSave(userId, discoveryCapabilityResult);

            return discoveryCache;
        }

        public static async Task<CapabilityDiscoveryResult> GetCapabilityDiscoveryResultAsync(ServiceCapabilities serviceCapability)
        {
            var discoveryServiceInfo = await GetAllCapabilityDiscoveryResultAsync();

            DiscoveryServiceCache discoveryCache = null;

            var cacheResult = await DiscoveryServiceCache.Load(serviceCapability);

            if (cacheResult == null)
            {
                discoveryCache = await CreateAndSaveDiscoveryServiceCacheAsync();

                CapabilityDiscoveryResult capabilityDiscoveryResult = null;

                discoveryCache.DiscoveryInfoForServices.TryGetValue(serviceCapability.ToString(), out capabilityDiscoveryResult);

                return capabilityDiscoveryResult;
            }
            else
            {
                return cacheResult.CapabilityResult;
            }
        }

        public static async Task<IDictionary<string, CapabilityDiscoveryResult>> GetAllCapabilityDiscoveryResultAsync()
        {
            DiscoveryServiceCache discoveryCache = null;

            var cacheResult = await DiscoveryServiceCache.Load();

            if (cacheResult == null)
            {
                discoveryCache = await CreateAndSaveDiscoveryServiceCacheAsync();

                return discoveryCache.DiscoveryInfoForServices;
            }

            return cacheResult.DiscoveryInfoForServices;
        }

        public static async Task<CapabilityDiscoveryResult> GetDiscoveryCapabilityResultAsync(string capability)
        {
            var cacheResult = await DiscoveryServiceCache.Load();

            CapabilityDiscoveryResult discoveryCapabilityResult = null;

            if (cacheResult != null && cacheResult.DiscoveryInfoForServices.ContainsKey(capability))
            {
                discoveryCapabilityResult = cacheResult.DiscoveryInfoForServices[capability];

                var initialAuthResult = await GetAccessTokenAsync(discoveryCapabilityResult.ServiceResourceId);

                if (initialAuthResult.UserInfo.UniqueId != cacheResult.UserId)
                {
                    // cache is for another user
                    cacheResult = null;
                }
            }

            if (cacheResult == null)
            {
                cacheResult = await CreateAndSaveDiscoveryServiceCacheAsync();
                discoveryCapabilityResult = cacheResult.DiscoveryInfoForServices[capability];
            }

            return discoveryCapabilityResult;
        }
    }
}
