using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Xml;

namespace Generate_order_details_of_customer
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
                document.MailMerge.ExecuteNestedGroup(GetRelationalData());
                //Removes empty page at the end of Word document
                RemoveEmptyPage(document);

                //Saves the file in the given path
                docStream = File.Create(Path.GetFullPath(@"../../../Sample.docx"));
                document.Save(docStream, FormatType.Docx);
                docStream.Dispose();
            }
        }       

        #region Helper Methods
        /// <summary>
        /// Gets the relational data to perform mail merge
        /// </summary>
        /// <returns></returns>
        static MailMergeDataTable GetRelationalData()
        {
            List<ExpandoObject> customers = new List<ExpandoObject>();
            Stream xmlStream = File.OpenRead(Path.GetFullPath(@"../../../CustomerDetails.xml"));
            XmlDocument xmlDocument = new XmlDocument();
            xmlDocument.Load(xmlStream);
            xmlStream.Dispose();

            ExpandoObject customer = new ExpandoObject();
            GetDataAsExpandoObject((xmlDocument as XmlNode).LastChild, ref customer);
            customers = (((customer as IDictionary<string, object>)["CustomerDetails"] as List<ExpandoObject>)[0] as IDictionary<string, object>)["Customers"] as List<ExpandoObject>;
            MailMergeDataTable dataTable = new MailMergeDataTable("Customers", customers);
            return dataTable;
        }

        /// <summary>
        /// Gets the data as ExpandoObject.
        /// </summary>
        /// <param name="reader">The reader.</param>
        /// <returns></returns>
        /// <exception cref="System.Exception">reader</exception>
        /// <exception cref="XmlException">Unexpected xml tag  + reader.LocalName</exception>
        private static void GetDataAsExpandoObject(XmlNode node, ref ExpandoObject dynamicObject)
        {
            if (node.InnerText == node.InnerXml)
                dynamicObject.TryAdd(node.LocalName, node.InnerText);
            else
            {
                List<ExpandoObject> childObjects;
                if ((dynamicObject as IDictionary<string, object>).ContainsKey(node.LocalName))
                    childObjects = (dynamicObject as IDictionary<string, object>)[node.LocalName] as List<ExpandoObject>;
                else
                {
                    childObjects = new List<ExpandoObject>();
                    dynamicObject.TryAdd(node.LocalName, childObjects);
                }
                ExpandoObject childObject = new ExpandoObject();
                foreach (XmlNode childNode in (node as XmlNode).ChildNodes)
                {
                    GetDataAsExpandoObject(childNode, ref childObject);
                }
                childObjects.Add(childObject);
            }           
        }
        /// <summary>
        /// Removes empty paragraphs from the end of Word document.
        /// </summary>
        /// <param name="document">The Word document</param>
        private static void RemoveEmptyPage(WordDocument document)
        {
            WTextBody textBody = document.LastSection.Body;           

            //A flag to determine any renderable item found in the Word document.
            bool IsRenderableItem = false;
            //Iterates text body items.
            for (int itemIndex = textBody.ChildEntities.Count - 1; itemIndex >= 0 && !IsRenderableItem; itemIndex--)
            {
                //Check item is empty paragraph and removes it.
                if (textBody.ChildEntities[itemIndex] is WParagraph)
                {
                    WParagraph paragraph = textBody.ChildEntities[itemIndex] as WParagraph;                    
                    //Iterates into paragraph
                    for (int pIndex = paragraph.Items.Count - 1; pIndex >= 0; pIndex--)
                    {
                        ParagraphItem paragraphItem = paragraph.Items[pIndex];

                        //If page break found in end of document, then remove it to preserve contents in same page
                        if ((paragraphItem is Break && (paragraphItem as Break).BreakType == BreakType.PageBreak))
                            paragraph.Items.RemoveAt(pIndex);

                        //Check paragraph contains any renderable items.
                        else if (!(paragraphItem is BookmarkStart || paragraphItem is BookmarkEnd))
                        {                            
                            IsRenderableItem = true;
                            //Found renderable item and break the iteration.
                            break;
                        }
                    }
                    //Remove empty paragraph and the paragraph with bookmarks only
                    if (paragraph.Items.Count == 0 || !IsRenderableItem)
                        textBody.ChildEntities.RemoveAt(itemIndex);
                }                
            }
        }
        #endregion
    }
}
