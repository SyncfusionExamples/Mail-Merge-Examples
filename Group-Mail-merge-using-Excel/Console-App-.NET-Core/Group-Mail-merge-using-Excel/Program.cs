using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.XlsIO;
using System.Collections.Generic;
using System.IO;

namespace Group_Mail_merge_using_Excel
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

                //Performs the mail merge for group
                document.MailMerge.ExecuteGroup(GetData());

                //Updates fields in the document
                document.UpdateDocumentFields();

                //Saves the file in the given path
                docStream = File.Create(Path.GetFullPath(@"../../../Sample.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }
        #region Helper Method
        /// <summary>
        /// Gets the data from Excel for mail merge
        /// </summary>
        /// <returns></returns>
        private static MailMergeDataTable GetData()
        {
            //Creates new excel engine
            ExcelEngine excelEngine = new ExcelEngine();
            //Creates new excel application
            IApplication application = excelEngine.Excel;

            //Opens the excel to extract data for mail merge
            Stream excelStream = File.OpenRead(Path.GetFullPath(@"../../../StockDetails.xlsx"));
            IWorkbook workbook = application.Workbooks.Open(excelStream);
            excelStream.Dispose();

            //Exports data from worksheet to .NET objects
            IWorksheet sheet = workbook.Worksheets[0];
            List<StockDetail> stockDetails = sheet.ExportData<StockDetail>(1, 1, 31, 5);
            workbook.Close();
            excelEngine.Dispose();
            return new MailMergeDataTable("StockDetails", stockDetails);
        }
        #endregion
    }
    #region Helper Class
    public class StockDetail
    {
        public string TradeNo { get; set; }
        public string CompanyName { get; set; }
        public string CostPrice { get; set; }
        public string SharesCount { get; set; }
        public string SalesPrice { get; set; }

        public StockDetail()
        {
        }
    }
    #endregion
}
