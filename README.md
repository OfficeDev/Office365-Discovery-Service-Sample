Office 365 Discovery Service Sample
==================================

This sample lists all the Office 365 capabilities available for the signed in Office 365 user such as Mail, Calendar, Contacts, My Files. It will discover the service endpoint Uri and resource Id for each capability which are required later for you to authenticate and get access tokens to those resources.

Note that, in order for your Windows Store application to authenticate to Office 365 discovery service, your Azure AD application should have permissions to at least one Office 365 API service. 

**Update 12/15/2014**
Sample now caches the information from the discovery service so the app doesn't need to query the discovery service everytime to create the respective service client. Since this is a Windows Store app, we use the local storage and write the information in a file called DiscoveryInfo.txt.

You can look at the following class files to understand how the caching works:
- [Office365\DiscoveryServiceCache.cs](https://github.com/OfficeDev/Office365-Discovery-Service-Sample/blob/master/Office365/DiscoveryServiceCache.cs)
- [Office365\Office365ServiceHelper.cs](https://github.com/OfficeDev/Office365-Discovery-Service-Sample/blob/master/Office365/Office365ServiceHelper.cs)

## How to Run this Sample
To run this sample, you need:

1. Visual Studio 2013
2. [Office 365 API Tools for Visual Studio 2013](http://aka.ms/OfficeDevToolsForVS2013)
3. [Office 365 Developer Subscription](https://aka.ms/devprogramsignup)

## Step 1: Clone or download this repository
From your Git Shell or command line:

`git clone https://github.com/OfficeDev/Office365-Discovery-Service-Sample.git`

## Step 2: Build the Project
1. Open the project in Visual Studio 2013.
2. Simply Build the project to restore NuGet packages.

## Step 3: Configure the sample

### Register Azure AD application to consume Office 365 APIs
Office 365 applications use Azure Active Directory (Azure AD) to authenticate and authorize users and applications respectively. All users, application registrations, permissions are stored in Azure AD.

Using the Office 365 API Tool for Visual Studio you can configure your Windows Store app to consume Office 365 APIs. 

1. In the Solution Explorer window, **right click your project -> Add -> Connected Service**.
2. A Services Manager dialog box will appear. Choose **Office 365 -> Office 365 API** and click **Register your app**.
3. On the sign-in dialog box, enter the username and password for your Office 365 tenant. 
4. After you're signed in, you will see a list of all the services. 
5. Initially, no permissions will be selected, as the app is not registered to consume any services yet.
6. Select **Mail** and then click **Permissions**
7. In the **Mail Permissions** dialog, select **Read users' mail** profiles' and click **Apply**
8. Select **Contacts** and then click **Permissions**
9. In the **Contacts Permissions** dialog, select **Read users' contacts** and click **Apply**
10. Click **Ok**

After clicking OK in the Services Manager dialog box, Office 365 client libraries (in the form of NuGet packages) for connecting to Office 365 APIs will be added to your project. 

In this process, Office 365 API tool registered an Azure AD Application in the Office 365 tenant that you signed in the wizard and added the Azure AD application details to web.config. 

### Step 4: Build and Debug your web application
Now you are ready for a test run. Hit F5 to test the app.
