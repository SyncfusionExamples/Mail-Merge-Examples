using System.Collections.Generic;
using System.IO;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;

namespace Start_each_record_on_new_page
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fileStream = new FileStream(Path.GetFullPath(@"../../../Template.docx"), FileMode.Open))
            {
                //Loads an existing Word document.
                using (WordDocument document = new WordDocument(fileStream, FormatType.Automatic))
                {
                    //Gets the invoice details as “IEnumerable” collection.
                    List<Invoice> invoice = GetInvoice();
                    //Creates an instance of “MailMergeDataTable” by specifying mail merge group name and “IEnumerable” collection.
                    MailMergeDataTable dataTable = new MailMergeDataTable("Invoice", invoice);
                    //Enables the flag to start each record in new page.
                    document.MailMerge.StartAtNewPage = true;
                    //Performs Mail merge.
                    document.MailMerge.ExecuteNestedGroup(dataTable);
                    //Creates file stream.
                    using (FileStream outputFileStream = new FileStream(Path.GetFullPath(@"../../../Result.docx"), FileMode.Create, FileAccess.ReadWrite))
                    {
                        //Saves the Word document to file stream.
                        document.Save(outputFileStream, FormatType.Docx);
                    }
                }
            }
        }

        #region Helper method
        /// <summary>
        /// Get the data to perform mail merge.
        /// </summary>
        public static List<Invoice> GetInvoice()
        {
            //Creates invoice details.
            List<Invoice> invoices = new List<Invoice>();

            List<Orders> orders = new List<Orders>();
            orders.Add(new Orders("10248", "Vins et alcools Chevalier", "59 rue de l'Abbaye", "Reims", "51100", "France", "VINET", "59 rue de l'Abbaye", "51100", "Reims", "France", "Steven Buchanan", "Vins et alcools Chevalier", "1996-07-04T00:00:00-04:00", "1996-08-01T00:00:00-04:00", "1996-07-16T00:00:00-04:00", "Federal Shipping"));

            List<Order> order = new List<Order>();
            order.Add(new Order("1", "Chai", "14.4", "45", "0.2", "518.4"));
            order.Add(new Order("2", "Boston Crab Meat", "14.7", "40", "0.2", "470.4"));

            List<OrderTotals> orderTotals = new List<OrderTotals>();
            orderTotals.Add(new OrderTotals("440", "32.8", "472.38"));

            invoices.Add(new Invoice(orders, order, orderTotals));

            orders = new List<Orders>();
            orders.Add(new Orders("10249", "Toms Spezialitäten", "Luisenstr. 48", "Münster", "51100", "Germany", "TOMSP", "Luisenstr. 48", "51100", "Münster", "Germany", "Michael Suyama", "Toms Spezialitäten", "1996-07-04T00:00:00-04:00", "1996-08-01T00:00:00-04:00", "1996-07-16T00:00:00-04:00", "Speedy Express"));

            order = new List<Order>();
            order.Add(new Order("1", "Chai", "18", "45", "0.2", "618.4"));
            order.Add(new Order("4", "Alice Mutton", "39", "100", "0", "3900"));

            orderTotals = new List<OrderTotals>();
            orderTotals.Add(new OrderTotals("1863.4", "11.61", "1875.01"));

            invoices.Add(new Invoice(orders, order, orderTotals));

            orders = new List<Orders>();
            orders.Add(new Orders("10250", "Hanari Carnes", "Rua do Paço, 67", "Rio de Janeiro", "05454-876", "Brazil", "VINET", "Rua do Paço, 67", "51100", "Rio de Janeiro", "Brazil", "Margaret Peacock", "Hanari Carnes", "1996-07-04T00:00:00-04:00", "1996-08-01T00:00:00-04:00", "1996-07-16T00:00:00-04:00", "United Package"));

            order = new List<Order>();
            order.Add(new Order("65", "Louisiana Fiery Hot Pepper Sauce", "16.8", "15", "0.15", "214.2"));
            order.Add(new Order("51", "Manjimup Dried Apples", "42.4", "35", "0.15", "1261.4"));

            orderTotals = new List<OrderTotals>();
            orderTotals.Add(new OrderTotals("1552.6", "65.83", "1618.43"));

            invoices.Add(new Invoice(orders, order, orderTotals));

            return invoices;
        }
        #endregion
    }

    #region Helper classes
    /// <summary>
    /// Represents a class to maintain invoice details.
    /// </summary>
    public class Invoice
    {
        #region Fields
        private List<Orders> orders;
        private List<Order> order;
        private List<OrderTotals> orderTotal;
        #endregion

        #region Properties
        public List<Orders> Orders
        {
            get { return orders; }
            set { orders = value; }
        }
        public List<Order> Order
        {
            get { return order; }
            set { order = value; }
        }
        public List<OrderTotals> OrderTotals
        {
            get { return orderTotal; }
            set { orderTotal = value; }
        }
        #endregion

        #region Constructor
        public Invoice(List<Orders> orders, List<Order> order, List<OrderTotals> orderTotals)
        {
            Orders = orders;
            Order = order;
            OrderTotals = orderTotals;
        }
        #endregion
    }
    /// <summary>
    /// Represents a class to maintain orders details.
    /// </summary>
    public class Orders
    {
        #region Fields
        private string orderID;
        private string shipName;
        private string shipAddress;
        private string shipCity;
        private string shipPostalCode;
        private string shipCountry;
        private string customerID;
        private string address;
        private string postalCode;
        private string city;
        private string country;
        private string salesPerson;
        private string customersCompanyName;
        private string orderDate;
        private string requiredDate;
        private string shippedDate;
        private string shippersCompanyName;
        #endregion

        #region Properties
        public string ShipName
        {
            get { return shipName; }
            set { shipName = value; }
        }
        public string ShipAddress
        {
            get { return shipAddress; }
            set { shipAddress = value; }
        }
        public string ShipCity
        {
            get { return shipCity; }
            set { shipCity = value; }
        }
        public string ShipPostalCode
        {
            get { return shipPostalCode; }
            set { shipPostalCode = value; }
        }
        public string PostalCode
        {
            get { return postalCode; }
            set { postalCode = value; }
        }
        public string ShipCountry
        {
            get { return shipCountry; }
            set { shipCountry = value; }
        }
        public string CustomerID
        {
            get { return customerID; }
            set { customerID = value; }
        }
        public string Customers_CompanyName
        {
            get { return customersCompanyName; }
            set { customersCompanyName = value; }
        }
        public string Address
        {
            get { return address; }
            set { address = value; }
        }
        public string City
        {
            get { return city; }
            set { city = value; }
        }
        public string Country
        {
            get { return country; }
            set { country = value; }
        }
        public string Salesperson
        {
            get { return salesPerson; }
            set { salesPerson = value; }
        }
        public string OrderID
        {
            get { return orderID; }
            set { orderID = value; }
        }
        public string OrderDate
        {
            get { return orderDate; }
            set { orderDate = value; }
        }
        public string RequiredDate
        {
            get { return requiredDate; }
            set { requiredDate = value; }
        }
        public string ShippedDate
        {
            get { return shippedDate; }
            set { shippedDate = value; }
        }
        public string Shippers_CompanyName
        {
            get { return shippersCompanyName; }
            set { shippersCompanyName = value; }
        }
        #endregion

        #region Constructor
        public Orders(string orderID, string shipName, string shipAddress, string shipCity,
         string shipPostalCode, string shipCountry, string customerID, string address,
         string postalCode, string city, string country, string salesPerson, string customersCompanyName,
         string orderDate, string requiredDate, string shippedDate, string shippersCompanyName)
        {
            OrderID = orderID;
            ShipName = shipName;
            ShipAddress = shipAddress;
            ShipCity = shipCity;
            ShipPostalCode = shipPostalCode;
            ShipCountry = shipCountry;
            CustomerID = customerID;
            Address = address;
            PostalCode = postalCode;
            City = city;
            Country = country;
            Salesperson = salesPerson;
            Customers_CompanyName = customersCompanyName;
            OrderDate = orderDate;
            RequiredDate = requiredDate;
            ShippedDate = shippedDate;
            Shippers_CompanyName = shippersCompanyName;
        }
        #endregion
    }
    /// <summary>
    /// Represents a class to maintain order details.
    /// </summary>
    public class Order
    {
        #region Fields
        private string productID;
        private string productName;
        private string unitPrice;
        private string quantity;
        private string discount;
        private string extendedPrice;
        #endregion

        #region Properties
        public string ProductID
        {
            get { return productID; }
            set { productID = value; }
        }
        public string ProductName
        {
            get { return productName; }
            set { productName = value; }
        }
        public string UnitPrice
        {
            get { return unitPrice; }
            set { unitPrice = value; }
        }
        public string Quantity
        {
            get { return quantity; }
            set { quantity = value; }
        }
        public string Discount
        {
            get { return discount; }
            set { discount = value; }
        }
        public string ExtendedPrice
        {
            get { return extendedPrice; }
            set { extendedPrice = value; }
        }
        #endregion

        #region Constructor       
        public Order(string productID, string productName, string unitPrice, string quantity,
         string discount, string extendedPrice)
        {
            ProductID = productID;
            ProductName = productName;
            UnitPrice = unitPrice;
            Quantity = quantity;
            Discount = discount;
            ExtendedPrice = extendedPrice;
        }
        #endregion
    }
    /// <summary>
    /// Represents a class to maintain order totals details.
    /// </summary>
    public class OrderTotals
    {
        #region Fields
        private string subTotal;
        private string freight;
        private string total;
        #endregion

        #region Properties
        public string Subtotal
        {
            get { return subTotal; }
            set { subTotal = value; }
        }
        public string Freight
        {
            get { return freight; }
            set { freight = value; }
        }
        public string Total
        {
            get { return total; }
            set { total = value; }
        }
        #endregion

        #region Constructor       
        public OrderTotals(string subTotal, string freight, string total)
        {
            Subtotal = subTotal;
            Freight = freight;
            Total = total;
        }
        #endregion
    }
    #endregion
}
