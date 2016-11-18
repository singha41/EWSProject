using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace Microsoft.Exchange.Samples.EWS
{
    class Program
    {
        private static readonly string Office365WebServicesURL= "https://outlook.office365.com/EWS/Exchange.asmx";

        static void Main(string[] args)
        {

            // Start tracing to console and a log file.
            Tracing.OpenLog("./GetStartedWithEWS.log");
            Tracing.WriteLine("EWS sample application started.");

            var isValidEmailAddress = false;
            Console.Write("Enter an email address: ");
            var emailAddress = Console.ReadLine();


            isValidEmailAddress = (emailAddress.Contains("@") && emailAddress.Contains("."));

            if (!isValidEmailAddress)
            {
                Tracing.WriteLine("Email address " + emailAddress + " is not a valid SMTP address. Closing program.");
                return;
            }

            SecureString password = GetPasswordFromConsole();
            if (password.Length == 0)
            {
                Tracing.WriteLine("Password empty, closing program.");
            }

            NetworkCredential userCredentials = new NetworkCredential(emailAddress, password);

            // These are the sample methods that demonstrate using EWS.
            // ShowNumberOfMessagesInInbox(userCredentials);
             SendTestEmail(userCredentials);

            Tracing.WriteLine("EWS sample application ends.");
            Tracing.CloseLog();

            Console.WriteLine("Press enter to exit: ");
            Console.ReadLine();
        }

        private static SecureString GetPasswordFromConsole()
        {
            SecureString password = new SecureString();
            bool readingPassword = true;

            Console.Write("Enter password: ");

            while (readingPassword)
            {
                ConsoleKeyInfo userInput = Console.ReadKey(true);

                switch (userInput.Key)
                {
                    case (ConsoleKey.Enter):
                        readingPassword = false;
                        break;
                    case (ConsoleKey.Escape):
                        password.Clear();
                        readingPassword = false;
                        break;
                    case (ConsoleKey.Backspace):
                        if (password.Length > 0)
                        {
                            password.RemoveAt(password.Length - 1);
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            Console.Write(" ");
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        }
                        break;
                    default:
                        if (userInput.KeyChar != 0)
                        {
                            password.AppendChar(userInput.KeyChar);
                            Console.Write("*");
                        }
                        break;
                }
            }
            Console.WriteLine();

            password.MakeReadOnly();
            return password;
        }

        // These method stubs will be filled in later.
        private static void ShowNumberOfMessagesInInbox(NetworkCredential userCredentials)


        {
            /// This is the XML request that is sent to the Exchange server.
            var getFolderSOAPRequest =
      "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" +
      "<soap:Envelope xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\"\n" +
      "   xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n" +
      "<soap:Header>\n" +
      "    <t:RequestServerVersion Version=\"Exchange2007_SP1\" />\n" +
      "  </soap:Header>\n" +
      "  <soap:Body>\n" +
      "    <GetFolder xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\"\n" +
      "               xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\">\n" +
      "      <FolderShape>\n" +
      "        <t:BaseShape>Default</t:BaseShape>\n" +
      "      </FolderShape>\n" +
      "      <FolderIds>\n" +
      "        <t:DistinguishedFolderId Id=\"inbox\"/>\n" +
      "      </FolderIds>\n" +
      "    </GetFolder>\n" +
      "  </soap:Body>\n" +
      "</soap:Envelope>\n";

            // Write the get folder operation request to the console and log file.
            Tracing.WriteLine("Get folder operation request:");
            Tracing.WriteLine(getFolderSOAPRequest);

          //  string Office365WebServicesURL = "https://outlook.office365.com/EWS/Exchange.asmx" ;
            var getFolderRequest = WebRequest.CreateHttp(Office365WebServicesURL);
            getFolderRequest.AllowAutoRedirect = false;
            getFolderRequest.Credentials = userCredentials;
            getFolderRequest.Method = "POST";
            getFolderRequest.ContentType = "text/xml";

            var requestWriter = new StreamWriter(getFolderRequest.GetRequestStream());
            requestWriter.Write(getFolderSOAPRequest);
            requestWriter.Close();

            try
            {
                var getFolderResponse = (HttpWebResponse)(getFolderRequest.GetResponse());
                if (getFolderResponse.StatusCode == HttpStatusCode.OK)
                {
                    var responseStream = getFolderResponse.GetResponseStream();
                    XElement responseEnvelope = XElement.Load(responseStream);
                    if (responseEnvelope != null)
                    {
                        // Write the response to the console and log file.
                        Tracing.WriteLine("Response:");
                        StringBuilder stringBuilder = new StringBuilder();
                        XmlWriterSettings settings = new XmlWriterSettings();
                        settings.Indent = true;
                        XmlWriter writer = XmlWriter.Create(stringBuilder, settings);
                        responseEnvelope.Save(writer);
                        writer.Close();
                        Tracing.WriteLine(stringBuilder.ToString());

                        // Check the response for error codes. If there is an error, throw an application exception.
                        IEnumerable<XElement> errorCodes = from errorCode in responseEnvelope.Descendants
                                                           ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                                                           select errorCode;
                        foreach (var errorCode in errorCodes)
                        {
                            if (errorCode.Value != "NoError")
                            {
                                switch (errorCode.Parent.Name.LocalName.ToString())
                                {
                                    case "Response":
                                        string responseError = "Response-level error getting inbox information:\n" + errorCode.Value;
                                        throw new ApplicationException(responseError);

                                    case "UserResponse":
                                        string userError = "User-level error getting inbox information:\n" + errorCode.Value;
                                        throw new ApplicationException(userError);
                                }
                            }
                        }

                        // Process the response.
                        IEnumerable<XElement> folders = from folderElement in
                                                          responseEnvelope.Descendants
                                                          ("{http://schemas.microsoft.com/exchange/services/2006/messages}Folders")
                                                        select folderElement;

                        foreach (var folder in folders)
                        {
                            Tracing.Write("Folder name:     ");
                            var folderName = from folderElement in
                                               folder.Descendants
                                               ("{http://schemas.microsoft.com/exchange/services/2006/types}DisplayName")
                                             select folderElement.Value;
                            Tracing.WriteLine(folderName.ElementAt(0));

                            Tracing.Write("Total messages:  ");
                            var totalCount = from folderElement in
                                               folder.Descendants
                                                 ("{http://schemas.microsoft.com/exchange/services/2006/types}TotalCount")
                                             select folderElement.Value;
                            Tracing.WriteLine(totalCount.ElementAt(0));

                            Tracing.Write("Unread messages: ");
                            var unreadCount = from folderElement in
                                               folder.Descendants
                                                 ("{http://schemas.microsoft.com/exchange/services/2006/types}UnreadCount")
                                              select folderElement.Value;
                            Tracing.WriteLine(unreadCount.ElementAt(0));
                        }
                    }
                }
            }
            catch (WebException ex)
            {
                Tracing.WriteLine("Caught Web exception:");
                Tracing.WriteLine(ex.Message);
            }
            catch (ApplicationException ex)
            {
                Tracing.WriteLine("Caught application exception:");
                Tracing.WriteLine(ex.Message);
            }

        }

        private static void SendTestEmail(NetworkCredential userCredentials)
        {

            var createItemSOAPRequest =
      "<?xml version=\"1.0\" encoding=\"utf-8\"?>\n" +
      "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" \n" +
      "               xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" \n" +
      "               xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" \n" +
      "               xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">\n" +
      "  <soap:Header>\n" +
      "    <t:RequestServerVersion Version=\"Exchange2007_SP1\" />\n" +
      "  </soap:Header>\n" +
      "  <soap:Body>\n" +
      "    <m:CreateItem MessageDisposition=\"SendAndSaveCopy\">\n" +
      "      <m:SavedItemFolderId>\n" +
      "        <t:DistinguishedFolderId Id=\"sentitems\" />\n" +
      "      </m:SavedItemFolderId>\n" +
      "      <m:Items>\n" +
      "        <t:Message>\n" +
      "          <t:Subject>Company Soccer Team</t:Subject>\n" +
      "          <t:Body BodyType=\"HTML\">Are you interested in joining?</t:Body>\n" +
      "          <t:ToRecipients>\n" +
      "            <t:Mailbox>\n" +
      "              <t:EmailAddress>singha@michigan.gov</t:EmailAddress>\n" +
      "              </t:Mailbox>\n" +
      "          </t:ToRecipients>\n" +
      "        </t:Message>\n" +
      "      </m:Items>\n" +
      "    </m:CreateItem>\n" +
      "  </soap:Body>\n" +
      "</soap:Envelope>\n";

            // Write the create item operation request to the console and log file.
            Tracing.WriteLine("Get folder operation request:");
            Tracing.WriteLine(createItemSOAPRequest);

            var getFolderRequest = WebRequest.CreateHttp(Office365WebServicesURL);
            getFolderRequest.AllowAutoRedirect = false;
            getFolderRequest.Credentials = userCredentials;
            getFolderRequest.Method = "POST";
            getFolderRequest.ContentType = "text/xml";

            var requestWriter = new StreamWriter(getFolderRequest.GetRequestStream());
            requestWriter.Write(createItemSOAPRequest);
            requestWriter.Close();

            try
            {
                var getFolderResponse = (HttpWebResponse)(getFolderRequest.GetResponse());
                if (getFolderResponse.StatusCode == HttpStatusCode.OK)
                {
                    var responseStream = getFolderResponse.GetResponseStream();
                    XElement responseEnvelope = XElement.Load(responseStream);
                    if (responseEnvelope != null)
                    {
                        // Write the response to the console and log file.
                        Tracing.WriteLine("Response:");
                        StringBuilder stringBuilder = new StringBuilder();
                        XmlWriterSettings settings = new XmlWriterSettings();
                        settings.Indent = true;
                        XmlWriter writer = XmlWriter.Create(stringBuilder, settings);
                        responseEnvelope.Save(writer);
                        writer.Close();
                        Tracing.WriteLine(stringBuilder.ToString());

                        // Check the response for error codes. If there is an error, throw an application exception.
                        IEnumerable<XElement> errorCodes = from errorCode in responseEnvelope.Descendants
                                                           ("{http://schemas.microsoft.com/exchange/services/2006/messages}ResponseCode")
                                                           select errorCode;
                        foreach (var errorCode in errorCodes)
                        {
                            if (errorCode.Value != "NoError")
                            {
                                switch (errorCode.Parent.Name.LocalName.ToString())
                                {
                                    case "Response":
                                        string responseError = "Response-level error getting inbox information:\n" + errorCode.Value;
                                        throw new ApplicationException(responseError);

                                    case "UserResponse":
                                        string userError = "User-level error getting inbox information:\n" + errorCode.Value;
                                        throw new ApplicationException(userError);
                                }
                            }
                        }

                        Tracing.WriteLine("Message sent successfully.");
                    }
                }
            }
            catch (WebException ex)
            {
                Tracing.WriteLine("Caught Web exception:");
                Tracing.WriteLine(ex.Message);
            }
            catch (ApplicationException ex)
            {
                Tracing.WriteLine("Caught application exception:");
                Tracing.WriteLine(ex.Message);
            }
        }
    }
}
