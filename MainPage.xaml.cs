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
using Microsoft.Office365.OutlookServices;
using System;
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

        static AuthenticationContext authContext = null;

        #endregion

        #region Constructors

        public MainPage()
        {
            this.InitializeComponent();
        }

        #endregion

        private async void btnGetAllCapabilities_Click(object sender, RoutedEventArgs e)
        {
            txtBoxStatus.Text = "";

            String discoveryResultText = "Capability: {0} \nEndpoint Uri: {1} \nResource Id: {2}\n\n";

            //var capabilitiesResult = await GetAllCapabilityDiscoveryResult();
            var capabilitiesResult = await Office365ServiceHelper.GetAllCapabilityDiscoveryResultAsync();

            foreach (var capability in capabilitiesResult)
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

            var capabilityContacts = ServiceCapabilities.Contacts.ToString();

            //CapabilityDiscoveryResult discoveryCapabilityResult = await GetDiscoveryCapabilityResultAsync(capability);
            CapabilityDiscoveryResult discoveryCapabilityResult = await Office365ServiceHelper.GetDiscoveryCapabilityResultAsync(capabilityContacts);

            txtBoxStatus.Text = String.Format(discoveryResultText,
                                                   capabilityContacts,
                                                   discoveryCapabilityResult.ServiceEndpointUri.ToString(),
                                                   discoveryCapabilityResult.ServiceResourceId).Replace("\n", Environment.NewLine);

            OutlookServicesClient outlookContactsClient = new OutlookServicesClient(discoveryCapabilityResult.ServiceEndpointUri,
               async () =>
               {
                   return await Office365ServiceHelper.GetAccessTokenForResourceAsync(discoveryCapabilityResult);

               });

            var contactsResults = await outlookContactsClient.Me.Contacts.ExecuteAsync();

            do
            {
                var contacts = contactsResults.CurrentPage;
                foreach (var contact in contacts)
                {
                    txtBoxStatus.Text += String.Format(discoveryResultText,
                                                    capabilityContacts,
                                                    contact.DisplayName,
                                                    contact.JobTitle).Replace("\n", Environment.NewLine);
                }

                contactsResults = await contactsResults.GetNextPageAsync();

            } while (contactsResults != null);
        }

        private async void btnDiscoverCalendar_Click(object sender, RoutedEventArgs e)
        {
            String discoveryResultText = "Capability: {0} \nEndpoint Uri: {1} \nResource Id: {2}\n\n";

            var capabilityCalendar = ServiceCapabilities.Calendar.ToString();

            CapabilityDiscoveryResult discoveryCapabilityResult = await Office365ServiceHelper.GetDiscoveryCapabilityResultAsync(capabilityCalendar);

            txtBoxStatus.Text = String.Format(discoveryResultText,
                                                   capabilityCalendar,
                                                   discoveryCapabilityResult.ServiceEndpointUri.ToString(),
                                                   discoveryCapabilityResult.ServiceResourceId).Replace("\n", Environment.NewLine);
        }

        private async void btnDiscoverMail_Click(object sender, RoutedEventArgs e)
        {
            String discoveryResultText = "Capability: {0} \nEndpoint Uri: {1} \nResource Id: {2}\n\n";

            var capabilityMail = ServiceCapabilities.Mail.ToString();

            CapabilityDiscoveryResult discoveryCapabilityResult = await Office365ServiceHelper.GetDiscoveryCapabilityResultAsync(capabilityMail);

            txtBoxStatus.Text = String.Format(discoveryResultText,
                                                   capabilityMail,
                                                   discoveryCapabilityResult.ServiceEndpointUri.ToString(),
                                                   discoveryCapabilityResult.ServiceResourceId).Replace("\n", Environment.NewLine);
        }

        private async void btnDiscoverMyFiles_Click(object sender, RoutedEventArgs e)
        {
            String discoveryResultText = "Capability: {0} \nEndpoint Uri: {1} \nResource Id: {2}\n\n";

            var capabilityMyFiles = ServiceCapabilities.MyFiles.ToString();

            CapabilityDiscoveryResult discoveryCapabilityResult = await Office365ServiceHelper.GetDiscoveryCapabilityResultAsync(capabilityMyFiles);

            txtBoxStatus.Text = String.Format(discoveryResultText,
                                                   capabilityMyFiles,
                                                   discoveryCapabilityResult.ServiceEndpointUri.ToString(),
                                                   discoveryCapabilityResult.ServiceResourceId).Replace("\n", Environment.NewLine);
        }

    }
}
