using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Pdf.Barcode;
using Syncfusion.Pdf.Graphics;
using System;
using System.Data;
using System.IO;

namespace Generate_Barcode_labels
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens an existing document from stream through constructor of `WordDocument` class
            FileStream fileStreamPath = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            WordDocument document = new WordDocument(fileStreamPath, FormatType.Automatic);
            //Creates mail merge events handler for image fields
            document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(InsertBarcode);
            //Gets data to perform mail merge
            DataTable table = GetDataTable();
            //Performs the mail merge
            document.MailMerge.ExecuteGroup(table);

            //Saves and closes the Word document instance
            FileStream outputStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite, FileShare.ReadWrite);
            document.Save(outputStream, FormatType.Docx);
            outputStream.Dispose();
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
                Stream imageStream= GenerateBarcodeImage(args.FieldValue.ToString());
                //Sets barcode image stream for merge field
                args.ImageStream = imageStream;
            }
        }
        /// <summary>
        /// Generates barcode image stream.
        /// </summary>
        /// <param name="barcodeText">Barcode text</param>
        /// <returns>Barcode image stream</returns>
        private static Stream GenerateBarcodeImage(string barcodeText)
        {
            //Initialize a new PdfCode39Barcode instance
            PdfCode39Barcode barcode = new PdfCode39Barcode();
            //Set the height and text for barcode
            barcode.BarHeight = 45;
            barcode.Text = barcodeText;
            //Convert the barcode to image 
            Stream imageStream = barcode.ToImage(new Syncfusion.Drawing.SizeF(145, 45));
            return imageStream;
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

            int Id = 10001;
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
