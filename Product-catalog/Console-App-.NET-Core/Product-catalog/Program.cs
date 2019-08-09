using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Drawing;
using System.Data;
using System.IO;

namespace Product_catalog
{
    class Program
    {
        // Create a DataSet.
        static DataSet ds = new DataSet();

        static void Main(string[] args)
        {			
            //Creates new Word document instance for Word processing
            using (WordDocument document = new WordDocument())
            {
                //Opens the Word template document
				Stream docStream = File.OpenRead(Path.GetFullPath(@"../../../Template.docx"));
				document.Open(docStream, FormatType.Docx);
				docStream.Dispose();
				
                //Get the tables from Data Set
				GetDataTable();			
				//Using Merge events to do conditional formatting during runtime
				document.MailMerge.MergeField += new MergeFieldEventHandler(AlternateRows_MergeField);
				document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_ProductImage);
			
				//Execute Mail Merge with groups
				document.MailMerge.ExecuteGroup(ds.Tables["Products"]);
				document.MailMerge.ExecuteGroup(ds.Tables["Product_PriceList"]);
			
				//Saves and closes the Word document
				docStream = File.Create(Path.GetFullPath(@"../../../Sample.docx"));
				document.Save(docStream, FormatType.Docx);
				docStream.Dispose();
            }
        }
        #region Helper Methods

        #region Event Handlers
        /// <summary>
        /// Method to handle MergeField event.
        /// </summary>
        private static void AlternateRows_MergeField(object sender, MergeFieldEventArgs args)
        {
            // Conditionally format data during Merge
            if (args.RowIndex % 2 == 0)
                args.CharacterFormat.TextColor = Color.FromArgb(255, 102, 0);
        }
        /// <summary>
        /// Method to handle MergeImageField event.
        /// </summary>     
        private static void MergeField_ProductImage(object sender, MergeImageFieldEventArgs args)
        {
            // Gets the image from disk during Merge
            if (args.FieldName == "ProductImage")
            {
                //Gets the image file name
                string ProductFileName = args.FieldValue.ToString();
                //Gets image from file system
                FileStream imageStream = new FileStream(@"../../../Data/" + ProductFileName, FileMode.Open, FileAccess.Read);
                //Sets the image for mail merge
                args.ImageStream = imageStream;      
            }
        }
        #endregion
        /// <summary>
        /// Gets the data to perform mail merge
        /// </summary>
        private static void GetDataTable()
        {
            //List of Syncfusion products name        
            string[] products = { "Apple Juice", "Grape Juice", "Hot Soup", "Tender Coconut", "Vennila", "Strawberry", "Cherry", "Cone" };

            //Add new Tables to the data set
            DataRow row;
            ds.Tables.Add();
            ds.Tables.Add();

            //Add fields to the Product_PriceList table.
            ds.Tables[0].TableName = "Product_PriceList";
            ds.Tables[0].Columns.Add("ProductName");
            ds.Tables[0].Columns.Add("Price");

            //Add fields to the Products table.
            ds.Tables[1].TableName = "Products";
            ds.Tables[1].Columns.Add("SNO");
            ds.Tables[1].Columns.Add("ProductName");
            ds.Tables[1].Columns.Add("ProductImage");

            int count = 0;

            //Inserting values to the tables.
            foreach (string product in products)
            {
                row = ds.Tables["Product_PriceList"].NewRow();
                row["ProductName"] = product;
                switch (product)
                {
                    case "Apple Juice":
                        row["Price"] = "$12.00"; break;
                    case "Grape Juice":
                        row["Price"] = "$15.00"; break;
                    case "Hot Soup":
                        row["Price"] = "$20.00"; break;
                    case "Tender coconut":
                        row["Price"] = "$10.00"; break;
                    case "Vennila Ice Cream":
                        row["Price"] = "$15.00"; break;
                    case "Strawberry":
                        row["Price"] = "$18.00"; break;
                    case "Cherry":
                        row["Price"] = "$25.00"; break;
                    default:
                        row["Price"] = "$20.00"; break;
                }

                ds.Tables["Product_PriceList"].Rows.Add(row);

                count++;
                row = ds.Tables["Products"].NewRow();
                row["SNO"] = count.ToString();
                row["ProductName"] = product;
                row["ProductImage"] = string.Concat(product, ".png");
                ds.Tables["Products"].Rows.Add(row);
            }
        }
        #endregion
    }
}
