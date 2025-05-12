using System;
using ExcelGen;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Dummy data to test the ExportData method (8 fields: fname, lname, age, city, email, phone, address, country)
            var strExportPath = @"C:\Users\deepa\OneDrive\Desktop\";
            var strSheetTitle = "Sheet1";
            var Data = new object[][]
            {
                new object[] { "mad", "Mad", 30, "'Ne'.'.  w Y ////'.'m .hj////*$/#@ ork'", "max@example.com", "123-456-7890", "123 Elm Street", "USA" },
                new object[] { "max", "Doe", 25, "'Ne'.'.  w Y  ork'", "john@example.com", "234-567-8901", "456 Maple Avenue", "USA" }
            };
            var strEFN = "Final.xls";
            var strTemplatePath = @"C:\Users\deepa\OneDrive\Desktop\X\";
            var strFileName = "Book1.xls";
            var strDSNtype = "System";
            var strDSNname = "ExcelFileDSN";
            bool blogEnabled = true;

            // DSN name
            //var dsnName = "Excel Files";  // Your DSN name here

            // Call the ExportData method from the ExcelGen library
            Return result = MyExcelGeneration.ExcelGenaration(strExportPath, strSheetTitle, Data, strEFN, strTemplatePath, strFileName, strDSNtype, strDSNname, blogEnabled);

            // Output the result (success or error message)
            if (result.m_bSuccess)
            {
                Console.WriteLine($"Success: {result.m_bSuccess}, Message: {result.m_strErrorMessage}");
            }
            else
            {
                Console.WriteLine($"Error: {result.m_strErrorMessage}");
            }
        }
    }
}
