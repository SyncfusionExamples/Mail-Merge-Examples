using Newtonsoft.Json.Linq;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;

namespace Create_and_send_email_messages
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing
            using (WordDocument template = new WordDocument())
            {
                //Opens the Word template document
                Stream docStream = File.OpenRead(Path.GetFullPath(@"../../../Template.docx"));
                template.Open(docStream, FormatType.Docx);
                docStream.Dispose();

                //Gets the recipient details as DataTable
                List<object> recipients = GetRecipients();
                foreach (var dataRecord in recipients)
                {
                    //Clones the template document for creating new document for each record in the data source
                    WordDocument document = template.Clone();

                    //Performs the mail merge
                    document.MailMerge.Execute(new List<object>() { dataRecord as Dictionary<string, object> });

                    //Save the HTML file as string
                    docStream = new MemoryStream();
                    document.SaveOptions.HtmlExportOmitXmlDeclaration = true;
                    document.Save(docStream, FormatType.Html);
                    //Releases the resources occupied by WordDocument instance
                    document.Dispose();
                    docStream.Position = 0;
                    StreamReader reader = new StreamReader(docStream);
                    string mailBody = reader.ReadToEnd();
                    docStream.Dispose();
                    if (mailBody.StartsWith("<!DOCTYPE"))
                        mailBody = mailBody.Remove(0, 97);
                    //Sends the email message
                    //Update the required e-mail id here
                    SendEMail("MailId@live.in", "RecipientMailId@live.in", "You order #" + (dataRecord as Dictionary<string, object>)["OrderID"].ToString() + " has been shipped", mailBody);
                }
            }
        }
        #region Helper methods
        private static void SendEMail(string from, string recipients, string subject, string body)
        {
            //Creates the email message
            var emailMessage = new MailMessage(from, recipients);
            //Adds the subject for email
            emailMessage.Subject = subject;
            //Sets the HTML string as email body
            emailMessage.IsBodyHtml = true;
            emailMessage.Body = body;
            //Sends the email with prepared message
            using (var client = new SmtpClient())
            {
                //Update your SMTP Server address here
                client.Host = "smtp.live.com";
                client.UseDefaultCredentials = false;
                //Update your email credentials here
                client.Credentials = new System.Net.NetworkCredential(from, "password");
                client.Port = 587;
                client.EnableSsl = true;
                client.Send(emailMessage);
            }
        }
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static List<object> GetRecipients()
        {
            //Reads the JSON object from JSON file.
            JObject jsonObject = JObject.Parse(File.ReadAllText(@"../../../CustomerDetails.json"));
            //Converts JSON object to Dictionary.
            IDictionary<string, object> data = GetData(jsonObject);
            return data["Customers"] as List<object>;
        }

        /// <summary>
        /// Gets data from JSON object.
        /// </summary>
        /// <param name="jsonObject">JSON object.</param>
        /// <returns>Dictionary of data.</returns>
        private static IDictionary<string, object> GetData(JObject jsonObject)
        {
            Dictionary<string, object> dictionary = new Dictionary<string, object>();
            foreach (var item in jsonObject)
            {
                object keyValue = null;
                if (item.Value is JArray)
                    keyValue = GetData((JArray)item.Value);
                else if (item.Value is JToken)
                    keyValue = ((JToken)item.Value).ToObject<string>();
                dictionary.Add(item.Key, keyValue);
            }
            return dictionary;
        }
        /// <summary>
        /// Gets array of items from JSON array.
        /// </summary>
        /// <param name="jArray">JSON array.</param>
        /// <returns>List of objects.</returns>
        private static List<object> GetData(JArray jArray)
        {
            List<object> jArrayItems = new List<object>();
            foreach (var item in jArray)
            {
                object keyValue = null;
                if (item is JObject)
                    keyValue = GetData((JObject)item);
                jArrayItems.Add(keyValue);
            }
            return jArrayItems;
        }
        #endregion
    }
}
