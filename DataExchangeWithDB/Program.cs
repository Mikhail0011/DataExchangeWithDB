using System.Data;
using System.Data.OleDb;
using System.Xml.Linq;
using System.Text;

string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=ECommerce.accdb;";
OleDbConnection connection = new(connectionString);

try
{
    Console.WriteLine("Start loading data from DB table to XML...");
    LoadOrdersDataFromDBToXML(connection);
    Console.WriteLine("Done!\n");

    Console.WriteLine("Start loading data from XML to DB table...");
    LoadDataFromXMLToDB(connection);
    Console.WriteLine("Done!\n");
}
catch (Exception ex)
{
    Console.WriteLine("\nError occurred: " + ex);
}
finally
{
    Console.WriteLine("\nPress any key to close the program...");
    Console.ReadKey();
    connection.Dispose();
}

static void LoadOrdersDataFromDBToXML(OleDbConnection connection)
{
    var dtOrders = new DataTable();
    var dtProducts = new DataTable();
    var dtUsers = new DataTable();

    using (OleDbDataAdapter adapter = new("select * from Orders", connection))
    {
        adapter.Fill(dtOrders);
    }

    using (OleDbDataAdapter adapter = new("select * from Products", connection))
    {
        adapter.Fill(dtProducts);
    }

    using (OleDbDataAdapter adapter = new("select * from Users", connection))
    {
        adapter.Fill(dtUsers);
    }

    var xmlElement = new XElement("Orders");

    foreach (DataRow row in dtOrders.Rows)
    {
        var product = dtProducts.Select("Product_id = " + row["Product_id"])[0];
        var user = dtUsers.Select("User_id = " + row["User_id"])[0];

        xmlElement.Add(new XElement("Order",
                                     new XElement("Order_id", row["Order_id"]),
                                     new XElement("Order_date", row["Order_date"]),
                                     new XElement("Order_sum", row["Order_sum"]),
                                     new XElement("Product_quantity", row["Product_quantity"]),
                                     new XElement("Product",
                                                  new XElement("Product_name", product["Product_name"]),
                                                  new XElement("Price", product["Price"]),
                                                  new XElement("Description", product["Description"])),
                                    new XElement("User",
                                                 new XElement("First_name", user["First_name"]),
                                                 new XElement("Second_name", user["Second_name"]),
                                                 new XElement("Third_name", user["Third_name"]),
                                                 new XElement("Birthdate", user["Birthdate"]),
                                                 new XElement("Phone", user["Phone"]),
                                                 new XElement("Email", user["Email"]),
                                                 new XElement("Address", user["Address"]))));
    }

    var xmlDoc = new XDocument();
    xmlDoc.Add(xmlElement);

    FileStream streamWrite = new("OrdersData.xml", FileMode.Create);
    xmlDoc.Save(streamWrite);
    streamWrite.Close();

    dtOrders.Clear();
    dtProducts.Clear();
    dtUsers.Clear();
}

static void LoadDataFromXMLToDB(OleDbConnection connect)
{
    connect.Open();

    foreach (var element in XElement.Load("Orders.xml").Elements("Order"))
    {
        var userNode = element.Element("User");

        var sqlGetUserId = new StringBuilder($"select User_id from Users where First_name = '{userNode.Element("First_name").Value}' and Second_name = '{userNode.Element("Second_name").Value}'");

        if (!String.IsNullOrEmpty(userNode.Element("Third_name").Value))
            sqlGetUserId.Append($" and Third_name = '{userNode.Element("Third_name").Value}'");
       
        sqlGetUserId.Append($" and Birthdate = #{DateTime.Parse(userNode.Element("Birthdate").Value).ToString("dd'/'MM'/'yyyy")}#;");

        var userId = (int)new OleDbCommand(sqlGetUserId.ToString(), connect).ExecuteScalar();

        var productId = (int)new OleDbCommand($"select Product_id from Products where Product_name = '{element.Element("Product").Element("Product_name").Value}'", connect).ExecuteScalar();

        string sqlCommand = $"insert into Orders values ({element.Element("Order_id").Value},@Date,{element.Element("Order_sum").Value},{userId},{productId},{element.Element("Product_quantity").Value})";
        
        OleDbCommand command = new(sqlCommand, connect);
        command.Parameters.AddWithValue("@Date", DateTime.Parse(element.Element("Order_date").Value));
        command.ExecuteNonQuery();
    }

    connect.Close();
}