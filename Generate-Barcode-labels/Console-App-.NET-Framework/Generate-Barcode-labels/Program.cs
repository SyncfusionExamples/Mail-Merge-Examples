using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Pdf.Barcode;
using System.Data;
using System.Drawing;
using System.IO;

namespace Generate_Barcode_labels
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the template document 
            WordDocument document = new WordDocument(Path.GetFullPath(@"../../Template.docx"));
            //Creates mail merge events handler for image fields
            document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(InsertBarcode);
            //Gets data to perform mail merge
            DataTable table = GetDataTable();
            //Performs the mail merge
            document.MailMerge.ExecuteGroup(table);

            //Saves and closes the Word document instance
            document.Save(Path.GetFullPath(@"../../Sample.docx"));
            document.Close();
        }

        #region Helper methods
        /// <summary>
        /// Inserts barcode in the Word document 
        /// </summary>
        private static void InsertBarcode(object sender, MergeImageFieldEventArgs args)
        {
            if (args.FieldName == "Barcode")
            {
                //Generates barcode image for field value.
                Image barcodeImage = GenerateBarcodeImage(args.FieldValue.ToString());
                //Sets barcode image for merge field
                args.Image = barcodeImage;
            }
        }
        /// <summary>
        /// Generates barcode image.
        /// </summary>
        /// <param name="barcodeText">Barcode text</param>
        /// <returns>Barcode image</returns>
        private static Image GenerateBarcodeImage(string barcodeText)
        {
            //Initialize a new PdfCode39Barcode instance
            PdfCode39Barcode barcode = new PdfCode39Barcode();
            //Set the height and text for barcode
            barcode.BarHeight = 45;
            barcode.Text = barcodeText;
            //Convert the barcode to image
            Image barcodeImage = barcode.ToImage(new SizeF(145, 45));
            return barcodeImage;
        }

        /// <summary>
        /// Gets the data to perform mail merge
        /// </summary>
        /// <returns></returns>
        private static DataTable GetDataTable()
        {
            // List of products name.
            string[] products = { "Apple Juice", "Grape Juice", "Hot Soup", "Tender Coconut", "Vennila", "Strawberry", "Cherry", "Cone",
                "Butter", "Milk", "Cheese", "Salt", "Honey", "Soap", "Chocolate", "Edible Oil", "Spices", "Paneer", "Curd", "Bread", "Olive oil", "Vinegar", "Sports Drinks",
                "Vegetable Juice", "Sugar", "Flour", "Jam", "Cake", "Brownies", "Donuts", "Egg", "Tooth Brush", "Talcum powder", "Detergent Soap", "Room Spray", "Tooth paste"};

            DataTable table = new DataTable("Product_PriceList");

            // Add fields to the Product_PriceList table.
            table.Columns.Add("ProductName");
            table.Columns.Add("Price");
            table.Columns.Add("Barcode");
            DataRow row;
            
            int Id =10001;
            // Inserting values to the tables.
            foreach (string product in products)
            {
                row = table.NewRow();
                row["ProductName"] = product;
                switch (product)
                {
                    case "Apple Juice":
                        row["Price"] = "$12.00";
                        break;
                    case "Grape Juice":
                    case "Milk":
                        row["Price"] = "$15.00";
                        break;
                    case "Hot Soup":
                        row["Price"] = "$20.00";
                        break;
                    case "Tender coconut":
                    case "Cheese":
                        row["Price"] = "$10.00";
                        break;
                    case "Vennila Ice Cream":
                        row["Price"] = "$15.00";
                        break;
                    case "Strawberry":
                    case "Butter":
                        row["Price"] = "$18.00";
                        break;
                    case "Cherry":
                    case "Salt":
                        row["Price"] = "$25.00";
                        break;
                    default:
                        row["Price"] = "$20.00";
                        break;
                }
                //Add barcode text
                row["Barcode"] = Id.ToString();
                table.Rows.Add(row);
                Id++;
            }
            return table;
        }
        #endregion
    }
}
