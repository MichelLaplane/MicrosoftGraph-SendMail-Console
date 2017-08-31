using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Reflection;
using System.Resources;
using System.Threading;
using System.Threading.Tasks;

namespace MicrosoftGraph_SendMail_Console
  {

  class Program
    {
    private static GraphServiceClient graphClient = null;
    private static readonly string strClientId = ConfigurationManager.AppSettings["clientId"];
    private static readonly string strRecipients = ConfigurationManager.AppSettings["distributionList"];
    private static readonly string strScopes = ConfigurationManager.AppSettings["scopes"];
    private static readonly string strAppSecret = ConfigurationManager.AppSettings["appsecret"];
    private static readonly string strTenantId = ConfigurationManager.AppSettings["tenantid"];
    private static readonly string strRedirectUri = ConfigurationManager.AppSettings["redirecturi"];
    public static PublicClientApplication IdentityClientApp = new PublicClientApplication(strClientId);

    private const string authorityFormat = "https://login.microsoftonline.com/{0}/v2.0";
    private const string msGraphScope = "https://graph.microsoft.com/.default";
    private const string msGraphQuery = "https://graph.microsoft.com/v1.0/users";

    public static string[] arScopes;
    public static string strTokenForUser = null;
    public static DateTimeOffset Expiration;
    public static DriveItem photoFile;
    public static string strBodyContent;
    public static string strMailSubject;
    public static Stream photoStream;
    public static string strBodyContentWithSharingLink;

    static void Main(string[] args)
      {
      string assemblyName = Assembly.GetCallingAssembly().GetName().Name;
      ResourceManager rmLocal = new ResourceManager(assemblyName + "." +
        "Properties.ApplicationResource", Assembly.GetCallingAssembly());
      CultureInfo cultInfo = Thread.CurrentThread.CurrentUICulture;
      strBodyContent = rmLocal.GetString("MailContents", cultInfo);
      arScopes = strScopes.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
      strMailSubject = rmLocal.GetString("MailSubject", cultInfo);

      try
        {
        graphClient = GetAuthenticatedClient();
        Task.Run(async () =>
        {
          // Lecture de la photo de l'utilisateur authentifié
          photoStream = await GetCurrentUserPhotoStreamAsync();
          if (photoStream == null)
            {
            // Pas de photo on met une image
            photoStream = System.IO.File.OpenRead(@"../../BlankMe.jpg");
            }
          MemoryStream photoStreamMS = new MemoryStream();
          photoStream.CopyTo(photoStreamMS);
          // Sauvegarde dans OneDrive du compte authentifié avec le nom me.png
          photoFile = await UploadFileToOneDriveAsync(photoStreamMS.ToArray());
          MessageAttachmentsCollectionPage attachments = new MessageAttachmentsCollectionPage();
          attachments.Add(new FileAttachment
            {
            ODataType = "#microsoft.graph.fileAttachment",
            ContentBytes = photoStreamMS.ToArray(),
            ContentType = "image/png",
            Name = "me.png"
            });
          // Récupération du lien de partage du fichier et insertion dans le corps du message
          Permission sharingLink = await GetSharingLinkAsync(photoFile.Id);
          strBodyContentWithSharingLink = String.Format(strBodyContent, sharingLink.Link.WebUrl);
          // Préparation de la liste des destinataires
          var splitRecipientsString = strRecipients.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
          List<Recipient> recipientList = new List<Recipient>();

          foreach (string recipient in splitRecipientsString)
            {
            recipientList.Add(new Recipient { EmailAddress = new EmailAddress { Address = recipient.Trim() } });
            }
          try
            {
            var email = new Message
              {
              Body = new ItemBody
                {
                Content = strBodyContentWithSharingLink,
                ContentType = BodyType.Html,
                },
              Subject = strMailSubject,
              ToRecipients = recipientList,
              Attachments = attachments
              };
            try
              {
              await graphClient.Me.SendMail(email, true).Request().PostAsync();
              }
            catch (ServiceException exception)
              {
              Console.WriteLine("We could not send the message: " + exception.Error == null ? "No error message returned." : exception.Error.Message);
              Console.ReadLine();
              return;
              }
            }

          catch (Exception e)
            {
            Console.WriteLine("We could not send the message: " + e.Message);
            return;
            }
        }).GetAwaiter().GetResult();
        }
      catch (Exception e)
        {
        Console.WriteLine("We could not send the message: " + e.Message);
        return;
        }
      finally
        {
        Console.WriteLine("\nMail envoyé {0} \n Appuyer sur une touche pour terminer.", DateTime.Now.ToUniversalTime());
        Console.ReadKey();
        }
      }


    // Get an access token for the given context and resourceId. An attempt is first made to 
    // acquire the token silently. If that fails, then we try to acquire the token by prompting the user.
    public static GraphServiceClient GetAuthenticatedClient()
      {
      if (graphClient == null)
        {
        // Create Microsoft Graph client.
        try
          {
          graphClient = new GraphServiceClient(
              "https://graph.microsoft.com/v1.0",
              new DelegateAuthenticationProvider(
                  async (requestMessage) =>
                  {
                    var token = await GetTokenForUserAsync();
                    requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", token);
                    // Header d'identification de la solution dans Microsoft Graph Service
                    requestMessage.Headers.Add("RegPortalGraphSendMail", "Send a Mail");
                  }));
          return graphClient;
          }
        catch (Exception ex)
          {
          Console.WriteLine("Could not create a graph client: " + ex.Message);
          }
        }

      return graphClient;
      }

    /// <summary>
    /// Récupération du token pour l'utilisateur
    /// </summary>
    /// <returns>Token</returns>
    public static async Task<string> GetTokenForUserAsync()
      {
      AuthenticationResult authResult;
      try
        {
        // Lecture du premier utilisateur authentifié
        authResult = await IdentityClientApp.AcquireTokenSilentAsync(arScopes, IdentityClientApp.Users.First());
        strTokenForUser = authResult.AccessToken;
        }
      catch (Exception)
        {
        // Pas d'utilisateur authentifié ou plus authentifié on demande l'autorisation
        if (strTokenForUser == null || Expiration <= DateTimeOffset.UtcNow.AddMinutes(5))
          {
          authResult = await IdentityClientApp.AcquireTokenAsync(arScopes);

          strTokenForUser = authResult.AccessToken;
          Expiration = authResult.ExpiresOn;
          }
        }

      return strTokenForUser;
      }


    // Récupération d'un stream sur la photo de l'utilisateur authentifié
    public static async Task<Stream> GetCurrentUserPhotoStreamAsync()
      {
      Stream currentUserPhotoStream = null;

      try
        {
        var graphClient = GetAuthenticatedClient();
        currentUserPhotoStream = await graphClient.Me.Photo.Content.Request().GetAsync();

        }
      // If the user account is MSA (not work or school), the service will throw an exception.
      catch (ServiceException)
        {
        // utilisateur Management Service Account non "work or school" 
        return null;
        }
      return currentUserPhotoStream;
      }


    // Chargement du fichier du OneDrive à la racine de l'utilisateur authentifié
    public static async Task<DriveItem> UploadFileToOneDriveAsync(byte[] file)
      {
      DriveItem uploadedFile = null;

      try
        {
        var graphClient = GetAuthenticatedClient();
        MemoryStream fileStream = new MemoryStream(file);
        uploadedFile = await graphClient.Me.Drive.Root.ItemWithPath("me.png").Content.Request().PutAsync<DriveItem>(fileStream);

        }
      catch (ServiceException)
        {
        return null;
        }
      return uploadedFile;
      }

    public static async Task<Permission> GetSharingLinkAsync(string Id)
      {
      Permission permission = null;

      try
        {
        var graphClient = GetAuthenticatedClient();
        permission = await graphClient.Me.Drive.Items[Id].CreateLink("view").Request().PostAsync();
        }
      catch (ServiceException)
        {
        return null;
        }
      return permission;
      }

    }
  }
