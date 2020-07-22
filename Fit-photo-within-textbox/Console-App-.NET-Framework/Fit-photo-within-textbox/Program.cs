using Syncfusion.DocIO.DLS;
using System.Drawing;
using System.IO;

namespace Fit_photo_within_textbox
{
    class Program
    {
        static void Main(string[] args)
        {
            //Opens the template document
            using (WordDocument document = new WordDocument(@"../../Template.docx"))
            {
                //Uses the mail merge events handler for image fields
                document.MailMerge.MergeImageField += new MergeImageFieldEventHandler(MergeField_ProductImage);
                //Specifies the field names and field values
                string[] fieldNames = new string[] { "Photo" };
                string[] fieldValues = new string[] { "Logo.jpg" };
                //Executes the mail merge with groups
                document.MailMerge.Execute(fieldNames, fieldValues);
                //Unhooks mail merge events handler
                document.MailMerge.MergeImageField -= new MergeImageFieldEventHandler(MergeField_ProductImage);
                //Saves the Word document instance
                document.Save(@"../../Sample.docx");
            }
            System.Diagnostics.Process.Start(Path.GetFullPath(@"../../Sample.docx"));
        }

        #region Helper method
        /// <summary>
        /// Binds the image from file system and fit within text box during Mail merge process by using MergeImageFieldEventHandler.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="args"></param>
        public static void MergeField_ProductImage(object sender, MergeImageFieldEventArgs args)
        {
            //Binds image from file system during mail merge
            if (args.FieldName == "Photo")
            {
                string ProductFileName = args.FieldValue.ToString();
                //Gets the image from file system
                args.Image = Image.FromFile(@"../../" + ProductFileName);
                //Gets the picture, to be merged for image merge field
                WPicture picture = args.Picture;
                //Gets the text box format
                WTextBoxFormat textBoxFormat = (args.CurrentMergeField.OwnerParagraph.OwnerTextBody.Owner as WTextBox).TextBoxFormat;
                //Resizes the picture to fit within text box
                float scalePercentage = 100;
                if (picture.Width != textBoxFormat.Width)
                {
                    //Calculates value for width scale factor
                    scalePercentage = textBoxFormat.Width / picture.Width * 100;
                    //This will resize the width
                    picture.WidthScale *= scalePercentage / 100;
                }
                scalePercentage = 100;
                if (picture.Height != textBoxFormat.Height)
                {
                    //Calculates value for height scale factor
                    scalePercentage = textBoxFormat.Height / picture.Height * 100;
                    //This will resize the height
                    picture.HeightScale *= scalePercentage / 100;
                }
            }
        }
        #endregion
    }
}
