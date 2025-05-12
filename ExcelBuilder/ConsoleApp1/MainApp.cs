using System;
using System.IO;
using System.Data.Odbc;

namespace ExcelGen
{
    public class MyExcelGeneration
    {
        public static Return ExcelGenaration(string strExportPath, string strSheetTitle, object[][] Data, string strEFN, string strTemplatePath, string strFileName, string strDSNtype, string strDSNname, bool blogEnabled)
        {
            Return objReturn = new Return();  // Default failure
            int rowStart = 0;
            try
            {
                // Generate full file paths for template and export file
                string strTemplateCopy = Path.Combine(strExportPath, strFileName);
                string strTemplateFilePath = Path.Combine(strTemplatePath, strFileName);

                // Validate the template file path
                if (!File.Exists(strTemplateFilePath))
                {
                    objReturn.m_strErrorMessage = "Template file not found at the specified path.";
                    return objReturn;
                }

                // Check if the file exists. If not, copy the template file
                if (!File.Exists(strTemplateCopy))
                {
                    if (!CopyTemplateFile(strTemplateFilePath, strTemplateCopy))
                    {
                        objReturn.m_strErrorMessage = "Failed to copy template file.";
                        return objReturn;  // Exit the function
                    }
                    rowStart = 2;  // Start from row 2 (since row 1 is headers)
                }
                else
                {
                    // Get the next available row index
                    rowStart = GetNextAvailableRowIndex(strTemplateCopy, strSheetTitle, strDSNtype, strDSNname);
                }

                // Open an ODBC connection to the Excel file using the User DSN
                string connString = GetConnString(strTemplateCopy, strDSNtype, strDSNname); // Accessing "Excel File" DSN under User DSN
                using (OdbcConnection conn = new OdbcConnection(connString))
                {
                    conn.Open();

                    // Insert each row of data into Excel, handling each cell as part of the row
                    foreach (var record in Data)
                    {
                        // Sanitize the record values before building the SQL query
                        string query = BuildInsertSQL(strSheetTitle, record, rowStart);
                        using (OdbcCommand cmd = new OdbcCommand(query, conn))
                        {
                            try
                            {
                                cmd.ExecuteNonQuery();
                                rowStart++;  // Increment rowStart after each row is inserted
                            }
                            catch (Exception ex)
                            {
                                objReturn.m_strErrorMessage = $"Error inserting record at row {rowStart}: {ex.Message}";
                                return objReturn;
                            }
                        }
                    }
                }

                // Set success message
                objReturn.m_bSuccess = true;
                objReturn.m_strErrorMessage = "Data Export Successful";
            }
            catch (Exception ex)
            {
                objReturn.m_strErrorMessage = $"Error: {ex.Message}";
            }

            // Return the result (success or error message)
            return objReturn;
        }

        // Helper function to copy the template file
        public static bool CopyTemplateFile(string sourceFile, string targetFile)
        {
            try
            {
                if (File.Exists(sourceFile))
                {
                    File.Copy(sourceFile, targetFile, overwrite: true);
                    return true;
                }
                else
                {
                    throw new FileNotFoundException("Source template file not found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error copying template file: {ex.Message}");
                return false;
            }
        }

        // Helper function to get the next available row index in the Excel sheet
        public static int GetNextAvailableRowIndex(string strTemplateCopy, string sheetName, string dsntype, string dsnName)
        {
            // Build the connection string using the passed parameters
            string connString = GetConnString(strTemplateCopy, dsntype, dsnName);

            using (OdbcConnection conn = new OdbcConnection(connString))
            {
                conn.Open();
                string query = $"SELECT COUNT(*) FROM [{sheetName}$]";
                using (OdbcCommand cmd = new OdbcCommand(query, conn))
                {
                    try
                    {
                        int rowCount = Convert.ToInt32(cmd.ExecuteScalar());
                        return rowCount + 1;  // Adding 1 because rowCount gives the current last row
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"Error getting row count: {ex.Message}");
                    }
                }
            }
        }

        // Helper function to get the Excel connection string using DSN
        public static string GetConnString(string filePath, string dsnType, string dsnName)
        {
            if (dsnType == "File")
            {
                // If File DSN, specify the full path to the .dsn file
                return $"Driver={{Microsoft Excel Driver (*.xls)}};DBQ={filePath};DSN={dsnName};ReadOnly=FALSE;";
            }
            else if (dsnType == "User" || dsnType == "System")
            {
                // If User or System DSN, specify the DSN name directly
                return $"Driver={{Microsoft Excel Driver (*.xls)}};DSN={dsnName};DBQ={filePath};ReadOnly=FALSE;";
            }
            else
            {
                throw new ArgumentException("Invalid DSN type specified. Use 'System', 'User', or 'File'.");
            }
        }

        // Helper function to build the SQL insert query for a row (instead of cell-by-cell)
        public static string BuildInsertSQL(string sheetName, object[] rowData, int rowIndex)
        {
            // Escape special characters in the row data
            for (int i = 0; i < rowData.Length; i++)
            {
                if (rowData[i] != null)
                {
                    rowData[i] = skipSplChar(rowData[i].ToString());
                }
            }

            // Excel columns are 1-indexed (A=1, B=2, C=3, etc.)
            return $"INSERT INTO [{sheetName}$] (fname, lname, age, city, email, phone, address, country) " +
                   $"VALUES ('{rowData[0]}', '{rowData[1]}', {rowData[2]}, '{rowData[3]}', '{rowData[4]}', '{rowData[5]}', " +
                   $"'{rowData[6]}', '{rowData[7]}')";
        }

        // Helper function to sanitize and escape special characters for SQL query
        public static string skipSplChar(string input) => input == null ? null :
        input.Replace("'", "''").Replace("\\", @"\\").Replace("\"", "\"\"").Replace("\n", " ").Replace("\r", " ")
         .Replace(";", " ").Replace(",", " ").Replace("*", " ").Replace("#", " ");

    }

    // Return class to hold the result of the operation
    public class Return
    {
        public bool m_bSuccess { get; set; }
        public string m_strErrorMessage { get; set; }

        public Return()
        {
            m_bSuccess = false;
            m_strErrorMessage = string.Empty;
        }
    }
}
