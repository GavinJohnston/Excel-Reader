using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ConsoleTableExt;
using OfficeOpenXml;

namespace ExcelReader
{
    public class ManageData {

            public static List<ExcelRecords> getExcel() {

            Console.WriteLine("Retrieving Excel Data..."); 

            // designates path to file
            string ExcelPath = "Spreadsheet.xlsx";
            // allows access to file
            FileInfo fileInfo = new FileInfo(ExcelPath);
            // allows access to excel document
            ExcelPackage package = new ExcelPackage(fileInfo);
            // allows access to specific worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            // defines max rows and columns, then cycles through each row of the document and outputs relevent data to a list
            int rows = worksheet.Dimension.Rows;

            int columns = worksheet.Dimension.Columns;

            List<ExcelRecords> ExcelData = new();

            for (int i = 2; i <= rows; i++) {
                // for (int j = 2; j <= columns; j++) {
                //     // writes the value of each cell to a string
                //     string content = worksheet.Cells[i, j].Value.ToString();

                    ExcelData.Add(
                        new ExcelRecords {
                            FirstName = worksheet.Cells[i, 1].Value.ToString(),
                            LastName = worksheet.Cells[i, 2].Value.ToString(),
                            Age = worksheet.Cells[i, 3].Value.ToString(),
                            Occupation = worksheet.Cells[i, 4].Value.ToString(),
                            Address = worksheet.Cells[i, 5].Value.ToString(),        
                        });

                // }
            }

            Console.WriteLine("Retrieval Complete."); 

            return ExcelData;
        }

        public static void showData() {

            var DatabaseData = Controller.getData();

            ConsoleTableBuilder
            .From(DatabaseData)
            .ExportAndWriteLine();

            Console.WriteLine("\nData Displayed, Press any Key to close the app..");

            Console.ReadLine();
        }

        public class ExcelRecords {
            public string FirstName {get; set;}
            public string LastName {get; set;}
            public string Age {get; set;}
            public string Occupation {get; set;}
            public string Address {get; set;}
        }
    }
}