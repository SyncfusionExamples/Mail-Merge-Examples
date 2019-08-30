using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;

namespace Generate_multiple_Word_documents
{
    class Program
    {
        static void Main(string[] args)
        {
            //Creates new Word document instance for Word processing
            using (WordDocument template = new WordDocument())
            {
                //Opens the Word template document
                template.Open(Path.GetFullPath(@"../../LetterTemplate.docx"), FormatType.Docx);

                //Gets the recipient details as DataTable
                DataTable recipients = GetRecipients();

                //Creates folder for saving generated documents
                if (!Directory.Exists(Path.GetFullPath(@"../../Result/")))
                    Directory.CreateDirectory(Path.GetFullPath(@"../../Result/"));
                foreach (DataRow dataRow in recipients.Rows)
                {
                    //Clones the template document for creating new document for each record in the data source
                    WordDocument document = template.Clone();

                    //Performs the mail merge
                    document.MailMerge.Execute(dataRow);

                    //Save the file in the given path
                    document.Save(Path.GetFullPath(@"../../Result/Letter_" + dataRow.ItemArray[2].ToString() + ".docx"), FormatType.Docx);
                    //Releases the resources occupied by WordDocument instance
                    document.Dispose();
                }
            }
        }
        #region Helper methods
        /// <summary>
        /// Gets the data to perform mail merge.
        /// </summary>
        /// <returns></returns>
        private static DataTable GetRecipients()
        {
            //Creates new DataTable instance 
            DataTable table = new DataTable();
            //Loads the database
            OleDbConnection conn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + @"../../CustomerDetails.mdb");
            //Opens the database connection
            conn.Open();
            OleDbDataAdapter adapter = new OleDbDataAdapter("Select * from Customers", conn);
            //Gets the data from the database
            adapter.Fill(table);
            //Releases the memory occupied by database connection
            adapter.Dispose();
            conn.Close();
            return table;
        }
        #endregion
    }
}
