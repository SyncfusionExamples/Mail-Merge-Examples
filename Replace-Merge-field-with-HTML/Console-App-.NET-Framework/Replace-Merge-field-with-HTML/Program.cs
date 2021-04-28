using Syncfusion.DocIO.DLS;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace Replace_Merge_field_with_HTML
{
    class Program
    {
        static Dictionary<WParagraph, Dictionary<int, string>> paraToInsertHTML = new Dictionary<WParagraph, Dictionary<int, string>>();
        static void Main(string[] args)
        {
            //Opens the template document.
            using (WordDocument document = new WordDocument(Path.GetFullPath(@"../../Template.docx")))
            {
                //Creates mail merge events handler to replace merge field with HTML.
                document.MailMerge.MergeField += new MergeFieldEventHandler(MergeFieldEvent);
                //Gets data to perform mail merge.
                DataTable table = GetDataTable();
                //Performs the mail merge.
                document.MailMerge.Execute(table);
                //Append HTML to paragraph.
                InsertHtml();
                //Removes mail merge events handler.
                document.MailMerge.MergeField -= new MergeFieldEventHandler(MergeFieldEvent);
                //Saves the Word document instance.
                document.Save(Path.GetFullPath(@"../../Sample.docx"));
            }
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
                if (args.FieldName.Equals("ProductList"))
                {
                    //Gets the current merge field owner paragraph.
                    WParagraph paragraph = args.CurrentMergeField.OwnerParagraph;
                    //Gets the current merge field index in the current paragraph.
                    int mergeFieldIndex = paragraph.ChildEntities.IndexOf(args.CurrentMergeField);
                    //Maintain HTML in collection.
                    Dictionary<int, string> fieldValues = new Dictionary<int, string>();
                    fieldValues.Add(mergeFieldIndex, args.FieldValue.ToString());
                    //Maintain paragraph in collection.
                    paraToInsertHTML.Add(paragraph, fieldValues);
                    //Set field value as empty.
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
            DataRow datarow = dataTable.NewRow();
            dataTable.Rows.Add(datarow);
            datarow["CustomerName"] = "Nancy Davolio";
            datarow["Address"] = "59 rue de I'Abbaye, Reims 51100, France";
            datarow["Phone"] = "1-888-936-8638";
            //Reads HTML string from the file.
            string htmlString = File.ReadAllText(@"../../File.html");
            datarow["ProductList"] = htmlString;
            return dataTable;
        }
        /// <summary>
        /// Append HTML to paragraph.
        /// </summary>
        private static void InsertHtml()
        {
            //Iterates through each item in the dictionary.
            foreach (KeyValuePair<WParagraph, Dictionary<int, string>> dictionaryItems in paraToInsertHTML)
            {
                WParagraph paragraph = dictionaryItems.Key as WParagraph;
                Dictionary<int, string> values = dictionaryItems.Value as Dictionary<int, string>;
                //Iterates through each value in the dictionary.
                foreach (KeyValuePair<int, string> valuePair in values)
                {
                    int index = valuePair.Key;
                    string fieldValue = valuePair.Value;
                    //Inserts HTML string at the same position of mergefield in Word document.
                    paragraph.OwnerTextBody.InsertXHTML(fieldValue, paragraph.OwnerTextBody.ChildEntities.IndexOf(paragraph), index);
                }
            }
            paraToInsertHTML.Clear();
        }
        #endregion
    }
}
