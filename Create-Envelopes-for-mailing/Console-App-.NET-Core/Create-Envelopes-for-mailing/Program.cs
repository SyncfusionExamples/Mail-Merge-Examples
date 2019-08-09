using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.IO;

namespace Create_Envelopes_for_mailing
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing
            using (WordDocument document = new WordDocument())
            {
                //Opens the Word template document
                Stream docStream = File.OpenRead(Path.GetFullPath(@"../../../Template.docx"));
                document.Open(docStream, FormatType.Docx);
                docStream.Dispose();

                //Gets the recipient details as "IEnumerable" collection of .NET objects
                List<Recipient> recipients = GetRecipients();

                //Performs the mail merge
                document.MailMerge.Execute(recipients);

                //Saves the file in the given path
                docStream = File.Create(Path.GetFullPath(@"../../../Sample.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }

        #region Helper methods
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static List<Recipient> GetRecipients()
        {
            List<Recipient> recipients = new List<Recipient>();
            //Initializes the recipient details
            recipients.Add(new Recipient("Nancy", "Davolio", "507 - 20th Ave. E.Apt. 2A", "Seattle", "WA", "98122", "USA"));
            recipients.Add(new Recipient("Andrew", "Fuller", "908 W. Capital Way", "Tacoma", "WA", "98401", "USA"));
            recipients.Add(new Recipient("Janet", "Leverling", "722 Moss Bay Blvd.", "Kirkland", "WA", "98033", "USA"));
            recipients.Add(new Recipient("Margaret", "Peacock", "4110 Old Redmond Rd.", "Redmond", "WA", "98052", "USA"));
            recipients.Add(new Recipient("Steven", "Buchanan", "14 Garrett Hil", "London", "", "SW1 8JR", "UK"));

            return recipients;
        }
        #endregion

    }

    #region Helper class
    /// <summary>
    /// Represents the Recipient details.
    /// </summary>
    class Recipient
    {
        #region Properties
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public string Country { get; set; }
        #endregion

        #region Constructor
        public Recipient(string firstName, string lastName, string address, string city, string state, string zipCode, string country)
        {
            FirstName = firstName;
            LastName = lastName;
            Address = address;
            City = city;
            State = state;
            ZipCode = zipCode;
            Country = country;
        }
        #endregion
    }
    #endregion
}
