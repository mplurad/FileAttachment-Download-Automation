using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace High_Radius_Invoice_Download_Automation
{
    public class Mail
    {
        /// <summary>
        /// Constructor to create a mail object. 
        /// These parameters set the properties for the Mail object and must have valid data.
        /// </summary>
        /// <param name="mailbox"> The mailbox holding the messages to look at. </param>
        /// <param name="baseUrl"> The base URL used for API calls to the MS Graph API. </param>
        /// <param name="authenticationResult"> The authentication result to acquire the bearer access token. </param>
        public Mail(string mailbox, string baseUrl, AuthenticationResult authenticationResult)
        {
            Mailbox = mailbox;
            BaseUrl = baseUrl;
            AuthenticationResult = authenticationResult;
            AccessToken = authenticationResult.AccessToken;
        }

        public string Mailbox { get; set; }
        public string BaseUrl { get; set; }
        public AuthenticationResult AuthenticationResult { get; set; }
        public string AccessToken { get; set; }

        /// <summary>
        /// Gets the first 1000 messages by retrieving them as a JObject and then converting them to Message objects.
        /// </summary>
        /// <returns> A collection of Message objects or null if there's an issue calling the API or converting the JObject. </returns>
        public async Task<ICollection<Message>> GetMessages()
        {
            var JObjectOfmessages = await ProtectedApiCallHelper.GetRequest($"{BaseUrl}?$top=1000", AccessToken); 
            var messages = ConvertJObjectToCollection<Message>(JObjectOfmessages);
            return messages;
        }

        /// <summary>
        /// Gets the attachments associated with a particular message by retrieving them as a JObject and then converting them to Attachment objects.
        /// </summary>
        /// <param name="messageId"> The string id of the message we want to grab attachments for. </param>
        /// <returns> A collection of Attachment objects or null if there's an issue calling the API or converting the JObject. </returns>
        public async Task<ICollection<Attachment>> GetAttachmentsByMessageId(string messageId)
        {
            var JObjectOfAttachments = await ProtectedApiCallHelper.GetRequest($"{BaseUrl}{messageId}/attachments", AccessToken);
            var attachments = ConvertJObjectToCollection<Attachment>(JObjectOfAttachments);
            return attachments;
        }

        /// <summary>
        /// Gets a file attachment from a message by retrieving it as a JObject and then converting it to a FileAttachment object.
        /// </summary>
        /// <param name="messageId"> The string id of the message we want to grab a file attachment from. </param>
        /// <param name="attachmentId"> the string id of the attachment we want to retrieve as a file attachment. </param>
        /// <returns> A FileAttachment or null if there's an issue calling the API or converting the JObject. </returns>
        public async Task<FileAttachment> GetFileAttachment(string messageId, string attachmentId)
        {
            var JObjectOfFileAttachment = await ProtectedApiCallHelper.GetRequest($"{BaseUrl}{messageId}/attachments/{attachmentId}", AccessToken);
            var fileAttachment = ConvertJObjectToObject<FileAttachment>(JObjectOfFileAttachment);
            return fileAttachment;
        }

        /// <summary>
        /// Deletes a message from the mailbox.
        /// </summary>
        /// <param name="messageId"> The string id of the message to be deleted. </param>
        public async void DeleteMessage(string messageId)
        {
            await ProtectedApiCallHelper.DeleteRequest($"{BaseUrl}{messageId}", AccessToken);
        }

        /// <summary>
        /// Checks
        /// </summary>
        /// <param name="attachmentName"></param>
        /// <param name="extensions"></param>
        /// <returns></returns>
        public static bool EndsWithExtension(string attachmentName, ICollection<string> extensions)
        {
            var extensionFound = false;
            foreach (var extension in extensions)
            {
                if (attachmentName.ToLower().EndsWith(extension.ToLower()))
                {
                    extensionFound = true;
                }
            }
            return extensionFound;
        }

        /// <summary>
        /// Download the file attachment to the specified file destination.
        /// </summary>
        /// <param name="fileAttachment"> The file to download. </param>
        /// <param name="fileDestination"> The destination the file will download to. </param>
        /// <returns></returns>
        public static bool DownloadFile(FileAttachment fileAttachment, string fileDestination)
        {
            try
            {
                System.IO.File.WriteAllBytes($"{fileDestination}{fileAttachment.Name}", fileAttachment.ContentBytes);
                Print.PrintText($"File \"{fileAttachment.Name}\" downloaded", ConsoleColor.Green);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Converts a JObject retrieved using the MS Graph API into a collection of T objects.
        /// </summary>
        /// <typeparam name="T"> A class that we'll convert the JObject to. </typeparam>
        /// <param name="result"> The JObject returned by a call to the MS Graph API. </param>
        /// <returns> A list of objects of type T. </returns>
        public ICollection<T> ConvertJObjectToCollection<T>(JObject result) where T : class
        {
            if (result == null)
                return null;

            ICollection<T> items = new List<T>();
            foreach (var item in result["value"])
            {
                items.Add(item.ToObject<T>());
            }
            return items;
        }

        /// <summary>
        /// Converts a JObject retrieved using the MS Graph API into a T object.
        /// </summary>
        /// <typeparam name="T"> A class that we'll convert the JObject to. </typeparam>
        /// <param name="result"> The JObject returned by a call to the MS Graph API. </param>
        /// <returns> A object of type T. </returns>
        public T ConvertJObjectToObject<T>(JObject result) where T : class
        {
            if (result == null)
                return null;

            return result.ToObject<T>();
        }
    }
}
