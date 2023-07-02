using MySql.Data.MySqlClient;
using SwansonParser;
using System.Data.SqlClient;

public class ProductRepository
{
    private readonly string _connectionString;

    public ProductRepository(string connectionString)
    {
        _connectionString = connectionString;
    }
    public void InsertProduct(Product product)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            connection.Open();

            var query = "INSERT INTO Products (Number, Title, Vendor, Price) " +
                        "VALUES (@Number, @Title, @Vendor, @Price)";

            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@Number", product.Number);
                command.Parameters.AddWithValue("@Title", product.Title);
                command.Parameters.AddWithValue("@Vendor", product.Vendor);
                command.Parameters.AddWithValue("@Price", product.Price);

                command.ExecuteNonQuery();
            }
        }
    }
}