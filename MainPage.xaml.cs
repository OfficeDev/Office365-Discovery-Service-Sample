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
using System.Linq;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;


namespace Win8ServiceDiscovery
{
    /// <summary>
    /// Discovery Service Page
    /// </summary>
    /// 
    public sealed partial class MainPage : Page
    {
        #region Private Fields and Constants

        //
        // The Client ID is used by the application to uniquely identify itself to Azure AD.
        // The Authority is the sign-in URL of the tenant.
        //
        static readonly string authority = "https://login.windows.net/Common/";
        static readonly string clientId = App.Current.Resources["ida:ClientID"].ToString();

        static readonly Uri discoveryServiceEndpointUri = new Uri("https://api.office.com/discovery/v1.0/me/");
        static readonly string discoveryServiceResourceId = "https://api.office.com/discovery/";

        static readonly string capabilityMail = "Mail";
        static readonly string capabilityCalendar = "Calendar";
        static readonly string capabilityContacts = "Contacts";
        static readonly string capabilityMyFiles = "MyFiles";

        static DiscoveryClient discoveryClient;
        static AuthenticationContext authContext = null;
        private Uri redirectURI = null;       

        #endregion

        #region Constructors
        public MainPage()
        {
            this.InitializeComponent();

            //
            // Every Windows Store application has a unique URI.
            // Windows ensures that only this application will receive messages sent to this URI.
            // ADAL uses this URI as the application's redirect URI to receive OAuth responses.
            // 
            // To determine this application's redirect URI, which is necessary when registering the app
            //      in AAD, set a breakpoint on the next line, run the app, and copy the string value of the URI.
            //      This is the only purposes of this line of code, it has no functional purpose in the application.
            //
            redirectURI = WebAuthenticationBroker.GetCurrentApplicationCallbackUri();           
        }
        #endregion

        private async Task<string> GetAccessToken(string serviceResourceId)
        {
            string accessToken = null;

            if(authContext == null)
            {
                authContext = new AuthenticationContext(authority);

                if (authContext.TokenCache.ReadItems().Count() > 0)
                {
                    var tokenCache = authContext.TokenCache.ReadItems().First();
                    authContext = new AuthenticationContext(tokenCache.Authority);
                }
            }                     

            var authResult = await authContext.AcquireTokenAsync(serviceResourceId, clientId, redirectURI);

            if(authResult.Status != AuthenticationStatus.Success)
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

            accessToken = authResult.AccessToken;

            return accessToken;
        }

        private DiscoveryClient GetDiscoveryClient()
        {
            if(discoveryClient == null)
            {
                discoveryClient = new DiscoveryClient(discoveryServiceEndpointUri,
                    async () =>
                    {
                        return await GetAccessToken(discoveryServiceResourceId);
                    });
            }

            return discoveryClient;

        }
        private async void btnGetAllCapabilities_Click(object sender, RoutedEventArgs e)
        {
            txtBoxStatus.Text = "";

            String discoveryResultText = "Capability: {0} \nEndpoint Uri: {1} \nResource Id: {2}\n\n";

            discoveryClient = GetDiscoveryClient();

            var discoveryCapabilityResult = await discoveryClient.DiscoverCapabilitiesAsync();

            foreach(var capability in discoveryCapabilityResult)
            {
                txtBoxStatus.Text += String.Format(discoveryResultText,
                                                   capability.Key,
                                                   capability.Value.ServiceEndpointUri.ToString(),
                                                   capability.Value.ServiceResourceId).Replace("\n", Environment.NewLine);

            }

        }

        private async void btnDiscoverContacts_Click(object sender, RoutedEventArgs e)
        {
            String discoveryResultText = "Capability: {0} \nEndpoint Uri: {1} \nResource Id: {2}\n\n";

            discoveryClient = GetDiscoveryClient();

            var discoveryCapabilityResult = await discoveryClient.DiscoverCapabilityAsync(capabilityContacts);

            txtBoxStatus.Text = String.Format(discoveryResultText,
                                                   capabilityContacts,
                                                   discoveryCapabilityResult.ServiceEndpointUri.ToString(),
                                                   discoveryCapabilityResult.ServiceResourceId).Replace("\n", Environment.NewLine);
        }

        private async void btnDiscoverCalendar_Click(object sender, RoutedEventArgs e)
        {
            String discoveryResultText = "Capability: {0} \nEndpoint Uri: {1} \nResource Id: {2}\n\n";

            discoveryClient = GetDiscoveryClient();

            var discoveryCapabilityResult = await discoveryClient.DiscoverCapabilityAsync(capabilityCalendar);

            txtBoxStatus.Text = String.Format(discoveryResultText,
                                                   capabilityCalendar,
                                                   discoveryCapabilityResult.ServiceEndpointUri.ToString(),
                                                   discoveryCapabilityResult.ServiceResourceId).Replace("\n", Environment.NewLine);
        }

        private async void btnDiscoverMail_Click(object sender, RoutedEventArgs e)
        {
            String discoveryResultText = "Capability: {0} \nEndpoint Uri: {1} \nResource Id: {2}\n\n";

            discoveryClient = GetDiscoveryClient();

            var discoveryCapabilityResult = await discoveryClient.DiscoverCapabilityAsync(capabilityMail);

            txtBoxStatus.Text = String.Format(discoveryResultText,
                                                   capabilityMail,
                                                   discoveryCapabilityResult.ServiceEndpointUri.ToString(),
                                                   discoveryCapabilityResult.ServiceResourceId).Replace("\n", Environment.NewLine);
        }

        private async void btnDiscoverMyFiles_Click(object sender, RoutedEventArgs e)
        {
            String discoveryResultText = "Capability: {0} \nEndpoint Uri: {1} \nResource Id: {2}\n\n";

            discoveryClient = GetDiscoveryClient();

            var discoveryCapabilityResult = await discoveryClient.DiscoverCapabilityAsync(capabilityMyFiles);

            txtBoxStatus.Text = String.Format(discoveryResultText,
                                                   capabilityMyFiles,
                                                   discoveryCapabilityResult.ServiceEndpointUri.ToString(),
                                                   discoveryCapabilityResult.ServiceResourceId).Replace("\n", Environment.NewLine);
        }

    }
}
