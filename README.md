# MicrosoftGraph-SendMail-Console
Sending a mail with Microsoft Graph AD V2.0 EndPoint with SDK API (No Rest API inside).
Great amount of the project code comes from the [Microsoft Graph UWP sample](https://github.com/microsoftgraph/uwp-csharp-connect-sample)  
# Register and configure the app
1. Sign into the [App Registration Portal](https://apps.dev.microsoft.com/) using either your personal or work or school account.
2. Select **Add an app**.
3. Enter a name for the app, and select **Create application**.
   The registration page displays, listing the properties of your app.
4. Under **Platforms**, select **Add platform**.
5. Select **Native Application**.
6. Copy both the Application Id and Redirect URI values to the clipboard. You'll need to enter these values into the sample app.
The app id is a unique identifier for your app. The redirect URI is a unique URI provided by Windows 10 for each application to ensure that messages sent to that URI are only sent to that application.
7. Select **Save**.
# Running the Application
Update the App.config file with your infos (Tenant, secret , ...)
  [ClientID] : Application ID
  [Dest1];[Dest2] : Email of destination users
  [AppSecret] : Not relevant for now
  [TenantID] : Not relevant for now

