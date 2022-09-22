using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Net;
using System.Net.Http;
using System.Security;
using System.Threading.Tasks;

namespace High_Radius_Invoice_Download_Automation
{
    static class Program
    {
        /// <summary>
        /// The properties set by arguments passed in through the command line.
        /// </summary>
        private static string Username { get; set; }
        private static SecureString Password { get; set; }
        private static string Mailbox { get; set; }
        private static string FileDestination { get; set; }
        private static string[] FileExtensions { get; set; }

        /// <summary>
        /// The HttpClient to last for the lifetime of the application.
        /// </summary>
        public static HttpClient HttpClient = new HttpClient();

        /// <summary>
        /// The application permissions granted.
        /// </summary>
        private static readonly string[] Scopes = { "Mail.Read", "Mail.Read.Shared" };

        static void Main(string[] args)
        {
            try
            {
                Print.PrintText($"Starting Program \"High Radius Invoice Download Automation\" - {DateTime.Now}\n", ConsoleColor.DarkGreen);

                if (args.Length == 5)
                {
                    InitializeCommandLineArguments(args[0], args[1], args[2], args[3], args[4]);
                    MailAutomationStartAsync().GetAwaiter().GetResult();
                }
                else
                {
                    Print.PrintText("Wrong number of arguments given at command line.", ConsoleColor.Red);
                }
            }
            catch (Exception ex)
            {
                Print.PrintText(ex.ToString(), ConsoleColor.Red);
            }
        }

        /// <summary>
        /// Initializes the arguments passed in through the command line into properties to be used by the class
        /// </summary>
        /// <param name="username"> The username of the Office 365 account to login. (E.g. "JohnDoe@companydomain.com") </param>
        /// <param name="password"> The password of the Office 365 account to login. </param> (E.g. "JoHnDo3sPaSsWoRd!")
        /// <param name="mailbox"> The mailbox holding the Office 365 messages to look at. The user must have access to this mailbox. </param> (E.g. "somemailbox@companydomain.com")
        /// <param name="fileDestination"> The destination which file attachments are downloaded to. (E.g. "C:/Users/John.Doe/Source/High Radius/" </param>
        /// <param name="fileExtensions"> The file extensions indicating which files should be downloaded. (E.g. ".xls .csv .xlsx") </param>
        private static void InitializeCommandLineArguments(string username, string password, string mailbox, string fileDestination, string fileExtensions)
        {
            Username = username;
            Password = new NetworkCredential("", password).SecurePassword;
            Mailbox = mailbox;
            FileDestination = fileDestination;
            FileExtensions = fileExtensions.Split();
        }

        /// <summary>
        /// Set up the automation by binding the values in the JSON file,
        /// getting the bearer token, initializing the mail object,
        /// and calling the method to process the mail.
        /// </summary>
        /// <returns> Returns an asynchronous operation. </returns>
        private static async Task MailAutomationStartAsync()
        {
            var app = CreateIPublicClientApplication();

            // Get the authentication result to access the bearer token
            var authenticationResult = await new PublicAppUsingUsernamePassword(app).GetTokenForWebApiUsingUsernamePasswordAsync(Scopes, Username, Password);

            // Set the base url for mail API calls
            var baseUrl = $"{AuthenticationConfig.MicrosoftGraphBaseEndpoint}/v1.0/users/{Mailbox}/mailFolders/inbox/messages/";

            // Create mail object and process it
            var mail = new Mail(Mailbox, baseUrl, authenticationResult);
            await MailAutomationProcessMail(mail);
        }

        /// <summary>
        /// Creates a public app by first creating an authentication config from a json file.
        /// </summary>
        /// <returns> Returns IPublicClientApplication. </returns>
        private static IPublicClientApplication CreateIPublicClientApplication()
        {
            AuthenticationConfig.ReadFromJsonFile("appsettings.json");
            var appConfig = AuthenticationConfig.PublicClientApplicationOptions;
            return PublicClientApplicationBuilder.CreateWithApplicationOptions(appConfig).Build();
        }

        /// <summary>
        /// Main method to process the mail by calling methods
        /// GetMessages to get all the messages,
        /// MailAutomationProcessMessage on each message,
        /// and finally DeleteMessage if an attachment was downloaded from that message.
        /// </summary>
        /// <param name="mail"> Mail object holding properties referencing an Office 365 mailbox with messages inside. </param>
        private static async Task MailAutomationProcessMail(Mail mail)
        {
            Print.PrintText($"Mailbox \"{mail.Mailbox}\"\n", ConsoleColor.Blue);

            var messages = await mail.GetMessages();
            if (messages != null)
            {
                foreach (var message in messages)
                {
                    var downloadedAnAttachment = await MailAutomationProcessMessage(mail, message);
                    if (downloadedAnAttachment)
                    {
                        mail.DeleteMessage(message.Id);
                        Print.PrintText($"Message \"{message.Subject}\" deleted", ConsoleColor.Magenta);
                    }
                    Console.WriteLine();
                }
            }
        }

        /// <summary>
        /// Helper method to process each message being iterated over in the calling method.
        /// If a message has an attachment, MailAutomationProcessAttachment will be called.
        /// </summary>
        /// <param name="mail"> Mail object holding properties referencing an Office 365 mailbox with messages inside. </param>
        /// <param name="message"> The message object being processed. </param>
        /// <returns> Returns true if an attachment was downloaded. Otherwise, returns false. </returns>
        private static async Task<bool> MailAutomationProcessMessage(Mail mail, Message message)
        {
            if (message == null)
                return false;

            Print.PrintText($"Message \"{message.Subject}\"", ConsoleColor.Cyan);

            if (!(bool)message.HasAttachments)
            {
                Console.WriteLine("No attachments available!");
                return false;
            }

            var attachments = await mail.GetAttachmentsByMessageId(message.Id);
            if (attachments == null)
                return false;

            var downloadedAnAttachment = false;
            foreach (var attachment in attachments)
            {
                if (downloadedAnAttachment)
                    await MailAutomationProcessAttachment(mail, message, attachment);
                else
                    downloadedAnAttachment = await MailAutomationProcessAttachment(mail, message, attachment);
            }
            return downloadedAnAttachment;
        }

        /// <summary>
        /// Helper method to process each attachment being iterated over in the calling method.
        /// If an attachment ends with a name contained in the string array FileExtensions, 
        /// method GetFileAttachment will get the file attachment from the associated attachment object
        /// and download that file attachment to the destination FileDestination.
        /// (FileExtensions and FileDestination are both arguments passed in through the command line.)
        /// </summary>
        /// <param name="mail"> Mail object holding properties referencing an Office 365 mailbox with messages inside. </param>
        /// <param name="message"> The message object being processed. </param>
        /// <param name="attachment"> The attachment object belonging to the message object being processed. </param>
        /// <returns> Returns true if an attachment was downloaded. Otherwise, returns false. </returns>
        private static async Task<bool> MailAutomationProcessAttachment(Mail mail, Message message, Attachment attachment)
        {
            if (attachment == null)
                return false;

            if (!Mail.EndsWithExtension(attachment.Name, FileExtensions))
            {
                Console.WriteLine($"File \"{attachment.Name}\" ignored");
                return false;
            }

            var fileAttachment = await mail.GetFileAttachment(message.Id, attachment.Id);
            if (fileAttachment == null)
                return false;

            var downloadedAnAttachment = Mail.DownloadFile(fileAttachment, FileDestination);
            return downloadedAnAttachment;
        }
    }
}
