# Generate order details of customer in C#

This example illustrates how to generate order details of each customer by performing **nested mail merge** operation for a specified region using [MailMergeDataTable](https://help.syncfusion.com/cr/document-processing/Syncfusion.DocIO.DLS.MailMergeDataTable.html) as the data source by [ExecuteNestedGroup(MailMergeDataTable dataTable)](https://help.syncfusion.com/cr/document-processing/Syncfusion.DocIO.DLS.MailMerge.html#Syncfusion_DocIO_DLS_MailMerge_ExecuteNestedGroup_Syncfusion_DocIO_DLS_MailMergeDataTable_) API.

# How to run the project

1. Download this project to a location in your disk.

2. Open the solution file using Visual Studio.

3. Rebuild the solution to install the required NuGet packages.

4. Run the application.

# Screenshots

By running this application, you will get the order details in a Word document as follows.

<p align="center">
<img src="Images/Generate-order-details-of-customer-output.png" alt="Generate-order-details-of-customer-output"/>
</p>

To create order details by nested mail merge operation, design your template Word document with the nested group of merge fields using Microsoft Word. In the below template, Customers is the owner group and it has two child groups, Orders and Products.

<p align="center">
<img src="Images/Generate-order-details-of-customer-template.png" alt="Generate-order-details-of-customer-template"/>
</p>

Take a moment to peruse the [documentation](https://help.syncfusion.com/document-processing/word/word-library/net/getting-started), where you will find other Word document processing operations along with features like [mail merge](https://www.syncfusion.com/document-sdk/net-word-library/mail-merge), [merge](https://www.syncfusion.com/document-sdk/net-word-library/merge-word-documents), and split documents, [find and replace](https://www.syncfusion.com/document-sdk/net-word-library/find-and-replace) text in the Word document, [protect](https://help.syncfusion.com/document-processing/word/word-library/net/working-with-security) Word documents, and most importantly [PDF](https://www.syncfusion.com/document-sdk/net-word-library/word-to-pdf-conversion) and [image](https://www.syncfusion.com/document-sdk/net-word-library/word-to-image-conversion) conversions with code examples.