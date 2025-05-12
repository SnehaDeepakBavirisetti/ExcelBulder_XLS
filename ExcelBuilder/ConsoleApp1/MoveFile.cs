using System;
using System.IO;

class MoveFile
{
    static void Main()
    {
        // Set the paths for Folder A and Folder B
        string folderA = @"C:\Users\deepa\OneDrive\Desktop\";
        string folderB = @"C:\Users\deepa\OneDrive\Desktop\Archive\";

        // Set the filename you want to check in Folder A
        string fileName = "Book1.xls";  // Change this to your file name

        // Create the full path for the file in Folder A
        string filePathA = Path.Combine(folderA, fileName);

        // Check if the file exists in Folder A
        if (File.Exists(filePathA))
        {
            try
            {
                // Get yesterday's date
                string yesterday = DateTime.Now.AddDays(-1).ToString("yyyyMMdd");

                // Generate the new filename by appending yesterday's date to the original filename
                string newFileName = Path.GetFileNameWithoutExtension(fileName) + "_" + yesterday + Path.GetExtension(fileName);

                // Create the full path for the new file in Folder B
                string filePathB = Path.Combine(folderB, newFileName);

                // Rename the file by moving it to Folder B
                File.Move(filePathA, filePathB);

                Console.WriteLine($"File moved and renamed to: {filePathB}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine($"File '{fileName}' does not exist in Folder A.");
        }
    }
}
