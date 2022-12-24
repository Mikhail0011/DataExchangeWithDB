using System.Data;
using System.Data.OleDb;
using System.Xml.Linq;

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

    using (OleDbDataAdapter adapter = new("SELECT * FROM Orders", connection))
    {
        adapter.Fill(dtOrders);
    }

    using (OleDbDataAdapter adapter = new("SELECT * FROM Products", connection))
    {
        adapter.Fill(dtProducts);
    }

    using (OleDbDataAdapter adapter = new("SELECT * FROM Users", connection))
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

static void LoadDataFromXMLToDB(OleDbConnection connection)
{
    connection.Open();

    foreach (var element in XElement.Load("Orders.xml").Elements("Order"))
    {
        string sqlCommand = $"INSERT INTO Orders VALUES (@OrderId, @Date, @OrderSum, @UserId, @ProductId, @ProductQuantity)";

        OleDbCommand command = new(sqlCommand, connection);

        command.Parameters.AddWithValue("@OrderId", element.Element("Order_id").Value);

        command.Parameters.AddWithValue("@Date", DateTime.Parse(element.Element("Order_date").Value));

        command.Parameters.AddWithValue("@OrderSum", element.Element("Order_sum").Value);

        command.Parameters.AddWithValue("@UserId", GetUserId(element.Element("User"), connection));

        command.Parameters.AddWithValue("@ProductId", GetProductId(element.Element("Product"), connection));

        command.Parameters.AddWithValue("@ProductQuantity", element.Element("Product_quantity").Value);

        command.ExecuteNonQuery();
    }

    connection.Close();
}

static int GetUserId(XElement userInfo, OleDbConnection connection)
{
    string sqlGetUserId = "SELECT * FROM Users WHERE First_name = @FirstName AND Second_name = @SecondName AND Birthdate = @Birthdate ";

    if (String.IsNullOrEmpty(userInfo.Element("Third_name").Value))
        sqlGetUserId += "AND Third_name IS NULL";
    else sqlGetUserId += "AND Third_name = @ThirdName";

    var getUserIdCommand = new OleDbCommand(sqlGetUserId, connection);

    getUserIdCommand.Parameters.AddWithValue("@FirstName", userInfo.Element("First_name").Value);
    getUserIdCommand.Parameters.AddWithValue("@SecondName", userInfo.Element("Second_name").Value);
    getUserIdCommand.Parameters.AddWithValue("@Birthdate", DateTime.Parse(userInfo.Element("Birthdate").Value));
    getUserIdCommand.Parameters.AddWithValue("@ThirdName", userInfo.Element("Third_name").Value);

    var reader = getUserIdCommand.ExecuteReader();

    if (reader.Read())
    {
        if (!userInfo.Element("Phone").Value.Equals(reader["Phone"]) || !userInfo.Element("Email").Value.Equals(reader["Email"]) || !userInfo.Element("Address").Value.Equals(reader["Address"]))
        {
            string sqlUpdateUser = "UPDATE Users SET Users.Phone = @Phone, Users.Email = @Email, Users.Address = @Address WHERE Users.User_id = @UserId;";

            var updateUserCommand = new OleDbCommand(sqlUpdateUser, connection);

            updateUserCommand.Parameters.AddWithValue("@Phone", userInfo.Element("Phone").Value);
            updateUserCommand.Parameters.AddWithValue("@Email", userInfo.Element("Email").Value);
            updateUserCommand.Parameters.AddWithValue("@Address", userInfo.Element("Address").Value);
            updateUserCommand.Parameters.AddWithValue("@UserId", reader["User_id"]);

            updateUserCommand.ExecuteNonQuery();
        }
        return (int)reader["User_id"];
    };

    string sqlInsertUser = "INSERT INTO Users (First_name, Second_name, Third_name, Birthdate, Phone, Email, Address) VALUES(@FirstName, @SecondName, @ThirdName, @Birthdate, @Phone, @Email, @Address);";

    var insertUserCommand = new OleDbCommand(sqlInsertUser, connection);

    insertUserCommand.Parameters.AddWithValue("@FirstName", userInfo.Element("First_name").Value);
    insertUserCommand.Parameters.AddWithValue("@SecondName", userInfo.Element("Second_name").Value);
    insertUserCommand.Parameters.AddWithValue("@ThirdName", String.IsNullOrEmpty(userInfo.Element("Third_name").Value) ? DBNull.Value : userInfo.Element("Third_name").Value);
    insertUserCommand.Parameters.AddWithValue("@Birthdate", DateTime.Parse(userInfo.Element("Birthdate").Value));
    insertUserCommand.Parameters.AddWithValue("@Phone", userInfo.Element("Phone").Value);
    insertUserCommand.Parameters.AddWithValue("@Email", userInfo.Element("Email").Value);
    insertUserCommand.Parameters.AddWithValue("@Address", userInfo.Element("Address").Value);

    insertUserCommand.ExecuteNonQuery();
    insertUserCommand.CommandText = "SELECT @@IDENTITY";

    return (int)insertUserCommand.ExecuteScalar();
}

static int GetProductId(XElement productInfo, OleDbConnection connection)
{
    string sqlGetProductId = "SELECT * FROM Products WHERE Product_name = @ProductName";
    var getProductIdCommand = new OleDbCommand(sqlGetProductId, connection);
    getProductIdCommand.Parameters.AddWithValue("@ProductName", productInfo.Element("Product_name").Value);

    var reader = getProductIdCommand.ExecuteReader();

    if (reader.Read())
    {
        if (!productInfo.Element("Price").Value.Equals(reader["Price"]) || !productInfo.Element("Description").Value.Equals(reader["Description"]))
        {
            string sqlUpdateUser = "UPDATE Products SET Products.Price = @Price, Products.Description = @Description WHERE Products.Product_id = @ProductId;";

            var updateUserCommand = new OleDbCommand(sqlUpdateUser, connection);

            updateUserCommand.Parameters.AddWithValue("@Price", productInfo.Element("Price").Value);
            updateUserCommand.Parameters.AddWithValue("@Description", productInfo.Element("Description").Value);
            updateUserCommand.Parameters.AddWithValue("@ProductId", reader["Product_id"]);

            updateUserCommand.ExecuteNonQuery();
        }
        return (int)reader["Product_id"];
    };

    string sqlInsertProduct = "INSERT INTO Products (Product_name, Price, Description) VALUES(@ProductName, @Price, @Description);";

    var insertProductCommand = new OleDbCommand(sqlInsertProduct, connection);

    insertProductCommand.Parameters.AddWithValue("@Product_name", productInfo.Element("Product_name").Value);
    insertProductCommand.Parameters.AddWithValue("@Price", productInfo.Element("Price").Value);
    insertProductCommand.Parameters.AddWithValue("@Description", String.IsNullOrEmpty(productInfo.Element("Description").Value) ? DBNull.Value : productInfo.Element("Description").Value);

    insertProductCommand.ExecuteNonQuery();
    insertProductCommand.CommandText = "SELECT @@IDENTITY";

    return (int)insertProductCommand.ExecuteScalar();
}