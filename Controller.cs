using System;
using System.Collections.Generic;
using System.Data.SqlClient;

namespace ExcelReader
{
    public class Controller
    {
        static string connectionString = @"Server=localhost,1433; Database=ExcelReader; User=sa; Password=someThingComplicated1234";
        public static void sendData() {
            using(var connection = new SqlConnection(connectionString)) {

                Console.WriteLine("Create Table..");

                connection.Open();

                var createTables = connection.CreateCommand();

                createTables.CommandText = @"
                    CREATE TABLE ExcelData (id INTEGER PRIMARY KEY IDENTITY, FirstName TEXT, LastName TEXT, Age TEXT, Occupation TEXT, Address TEXT);";

                createTables.ExecuteNonQuery();

                Console.WriteLine("Table Created.");

                Console.WriteLine("Populating Table..");

                var ExcelData = ManageData.getExcel();

                foreach (ManageData.ExcelRecords item in ExcelData)
                {
                    var addData = connection.CreateCommand();

                    addData.CommandText = $"INSERT INTO ExcelData (FirstName, LastName, Age, Occupation, Address) VALUES('{item.FirstName}', '{item.LastName}', '{item.Age}', '{item.Occupation}', '{item.Address}')";

                    addData.ExecuteNonQuery();
                }

                Console.WriteLine("Table Populated.");
            }
        }
        public static List<ManageData.ExcelRecords> getData() {
            using(var connection = new SqlConnection(connectionString)) {
                Console.WriteLine("Retrieving Data From Database..");

                connection.Open();

                var getData = connection.CreateCommand();

                getData.CommandText = $"SELECT * FROM ExcelData";

                List<ManageData.ExcelRecords> Records = new();

                SqlDataReader reader = getData.ExecuteReader();

                if (reader.HasRows) {
                while (reader.Read()) {
                    Records.Add(
                        new ManageData.ExcelRecords {
                            FirstName = reader.GetString(1),
                            LastName = reader.GetString(2),
                            Age = reader.GetString(3),
                            Occupation = reader.GetString(4),
                            Address = reader.GetString(5)
                        }); 
                }
                Console.WriteLine("Retrieved Data..");
                return Records;
            } else {
                return null;
            }
            }
        }
    }
}