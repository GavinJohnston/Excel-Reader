using System;
using System.Data.SqlClient;

namespace ExcelReader
{
    class Program
    {
        static string connectionString = @"Server=localhost,1433; User=sa; Password=someThingComplicated1234";
        static void Main(string[] args)
        {
            using(var connection = new SqlConnection(connectionString)) {
                connection.Open();

                var createDb = connection.CreateCommand();

                createDb.CommandText = @"
                IF EXISTS (SELECT * FROM sys.databases WHERE name = 'ExcelReader')
                BEGIN
                DROP DATABASE ExcelReader;
                END;
                CREATE DATABASE ExcelReader;";

                createDb.ExecuteNonQuery();

                connection.Close();
            }

            Console.WriteLine("Welcome to Excel Reader. Press any key to retrieve data.");

            Console.ReadLine();

            ManageData.getExcel();

            Controller.sendData();

            ManageData.showData();
        }
    }
}
