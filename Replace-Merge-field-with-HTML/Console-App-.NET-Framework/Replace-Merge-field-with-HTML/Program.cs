using Syncfusion.DocIO.DLS;
using System.Data;
using System.IO;

namespace Replace_Merge_field_with_HTML
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the template document
            WordDocument document = new WordDocument(Path.GetFullPath(@"../../Template.docx"));
            //Creates mail merge events handler to replace merge field with HTML
            document.MailMerge.MergeField += new MergeFieldEventHandler(MergeFieldEvent);
            //Gets data to perform mail merge
            DataTable table = GetDataTable();
            //Performs the mail merge
            document.MailMerge.Execute(table);
            //Saves and closes the Word document instance
            document.Save(Path.GetFullPath(@"../../Sample.docx"));
            document.Close();
            System.Diagnostics.Process.Start(Path.GetFullPath(@"../../Sample.docx"));
        }

        #region Helper methods
        /// <summary>
        /// Replaces merge field with HTML string by using MergeFieldEventHandler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public static void MergeFieldEvent(object sender, MergeFieldEventArgs args)
        {
            if (args.TableName.Equals("HTML"))
            {
                if (args.FieldName.Equals("ProductList") || args.FieldName.Equals("CompanyName"))
                {
                    string text = args.FieldValue as string;
                    WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                    int paraIndex = paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph);
                    int fieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);
                    paragraph.OwnerTextBody.InsertXHTML(args.FieldValue.ToString(), paraIndex, fieldIndex);
                    args.Text = string.Empty;
                }
            }
        }
        /// <summary>
        /// Gets the data to perform mail merge
        /// </summary>
        /// <returns></returns>
        private static DataTable GetDataTable()
        {
            DataTable dataTable = new DataTable("HTML");
            dataTable.Columns.Add("CustomerName");
            dataTable.Columns.Add("Address");
            dataTable.Columns.Add("Phone");
            dataTable.Columns.Add("ProductList");
            dataTable.Columns.Add("CompanyName");
            DataRow datarow = dataTable.NewRow();
            dataTable.Rows.Add(datarow);
            datarow["CustomerName"] = "Nancy Davolio";
            datarow["Address"] = "59 rue de I'Abbaye, Reims 51100, France";
            datarow["Phone"] = "1-888-936-8638";
            datarow["ProductList"] = "<html><body><div style=\"font-family:calibri;\"><ul><li>Mountain-200</li><li>Mountain-300</li><li>Road-150</li></ul></div></body></html>";
            datarow["CompanyName"] = "<html><body><p style=\"font-family:calibri;text-align:right\"><b><font color='blue'>Adventure Works Cycles</font></b></p></body></html>";
            return dataTable;
        }
        #endregion
    }
}
