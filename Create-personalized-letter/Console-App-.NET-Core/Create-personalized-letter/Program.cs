using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.IO;

namespace Create_personalized_letter
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the Word template document
            Stream docStream = File.OpenRead(Path.GetFullPath(@"../../../LetterTemplate.docx"));
            WordDocument document = new WordDocument(docStream, FormatType.Docx);
            docStream.Dispose();

            //Loads the string arrays with field names and values for mail merge
            string[] fieldNames = { "ContactName", "CompanyName", "Address", "City", "Country", "Phone" };
            string[] fieldValues = { "Nancy Davolio", "Syncfusion", "507 - 20th Ave. E.Apt. 2A", "Seattle, WA", "USA", "(206) 555-9857-x5467" };

            //Performs the mail merge
            document.MailMerge.Execute(fieldNames, fieldValues);

            //Saves the Word document as DOCX format
            docStream = File.Create(Path.GetFullPath(@"../../../Sample.docx"));
            document.Save(docStream, FormatType.Docx);
            docStream.Dispose();
            //Releases the resources occupied by WordDocument instance
            document.Dispose();
        }
    }
}
