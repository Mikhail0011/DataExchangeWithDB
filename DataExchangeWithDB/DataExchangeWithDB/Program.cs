using System.Data;
using System.Data.OleDb;
using System.Xml.Linq;

string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=ECommerce.accdb;";
OleDbConnection connection = new(connectionString);

try
{
    Console.WriteLine("Start loading data from DB table to XML...");
    LoadTableFromDBToXML(connection, "Orders");
    Console.WriteLine("Done!\n");

    Console.WriteLine("Start loading data from XML to DB table...");
    LoadDataFromXMLToDB(connection);
    Console.WriteLine("Done!\n");
}
catch (Exception ex)
{
    Console.WriteLine("\nError occurred: " + ex.Message);
}
finally
{
    Console.WriteLine("\nPress any key to close the program...");
    Console.ReadKey();
    connection.Dispose();
}


static void LoadTableFromDBToXML(OleDbConnection connection, string tableName)
{
    using (OleDbDataAdapter adapter = new("select * from " + tableName, connection))
    {
        var ds = new DataSet();

        connection.Open();
        adapter.Fill(ds, "Orders");
        connection.Close();

        FileStream streamWrite = new("OrdersDataFromDB.xml", FileMode.Create);
        ds.WriteXml(streamWrite);
        streamWrite.Close();
    }
}

static void LoadDataFromXMLToDB(OleDbConnection connect)
{
    connect.Open();

    foreach (var element in XElement.Load("Orders.xml").Elements("Orders"))
    {
        string sqlCommand = $"insert into Orders values ({element.Element("Order_id").Value},@Date,{element.Element("Order_sum").Value},{element.Element("User_id").Value},{element.Element("Product_id").Value},{element.Element("Product_quantity").Value})";
        OleDbCommand command = new(sqlCommand, connect);
        command.Parameters.AddWithValue("@Date", DateTime.Parse(element.Element("Order_date").Value));
        command.ExecuteNonQuery();
    }

    connect.Close();
}