using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace FileColumnReader
{
    public class FileHelper
    {

        public static void ValidateFileHeaderForSelectedContentSize(byte[] fileBytes, int maxNumberOfTransactionRows = 250)
        {
            if (fileBytes == null || fileBytes.Length == 0)
                throw new Exception("Invalid uploaded file parameters.");

            //var header = "Ct.;Date of Wallk-In;First Name;Full Surname Patrilineal-Matrilineal(as applicable);Language  (drop Down);Other Language, please specify;Email;Cell Phone Number;Other Phone Number;DOB;A Number;Family Unit size;Referral?;Referred By";

            Dictionary<string, string> fieldRequirements = new Dictionary<string, string>
            {
                { "Ct.", "NotRequired" },
                { "Date of Wallk-In", "NotRequired" },
                { "First Name", "Required" },
                { "Full Surname Patrilineal-Matrilineal(as applicable)", "Required" },
                { "Language  (drop Down)", "NotRequired" },
                { "Other Language, please specify", "NotRequired" },
                { "Email", "Required" },
                { "Cell Phone Number", "Required" },
                { "Other Phone Number", "NotRequired" },
                { "DOB", "NotRequired" },
                { "A Number", "Required" },
                { "Family Unit size", "NotRequired" },
                { "Referral?", "NotRequired" },
                { "Referred By", "NotRequired" },
                {"Location", "Required" }
            };

            var _columnName = fieldRequirements.Keys.ToArray()[0];
            // Split the header to get the expected column names
            //List<string> orderedHeader = header.Split(';').ToList();

            // Create a MemoryStream from the byte array
            using (var stream = new MemoryStream(fileBytes))
            {
                using (var package = new ExcelPackage(stream))
                {
                    var sheet = package.Workbook.Worksheets.First();

                    int noOfCol = 0;
                    for (int column = 1; column <= sheet.Dimension.End.Column; column++)  // Loop through all columns
                    {
                        var cellValue = sheet.Cells[6, column].Value;  // Check cell in the 6th row (headers)
                        if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        {
                            noOfCol++;
                        }
                    }

                    int noOfRow = 0;
                    for (int row = 7; row <= sheet.Dimension.End.Row; row++)  // Start from row 7 as per your requirement
                    {
                        var cellValue = sheet.Cells[row, 4].Value;  // Check column 4 for data
                        if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                        {
                            noOfRow++;
                        }
                    }

                    if (fieldRequirements.Count == noOfCol)
                    {
                        // Get the column names from the file (header row, 6th row)
                        List<string> columns = new List<string>();
                        for (int column = 1; column <= noOfCol; column++)
                        {
                            var cellValue = sheet.Cells[6, column].Value;
                            if (cellValue != null && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                            {
                                // Add non-empty column to the list
                                columns.Add(cellValue.ToString().Replace("\n", " ").Trim());
                            }
                            else
                            {
                                // Handle empty columns: print the column name (from dictionary) and that it's empty
                                var columnName = fieldRequirements.Keys.ToArray()[column - 1];  // Get corresponding dictionary key
                                Console.WriteLine($"Column '{columnName}' is empty in the file.");
                            }
                        }

                        // Check if each column name matches the corresponding key in the dictionary
                        for (int i = 0; i < columns.Count; i++)
                        {
                            if (columns[i] != fieldRequirements.Keys.ToArray()[i])
                            {
                                var mesg = $"Mismatch at index {i}: Column '{columns[i]}' does not match name in the template '{fieldRequirements.Keys.ToArray()[i]}'.";
                                Console.WriteLine(mesg);
                            }
                        }

                    }
                    else
                    {
                        Console.WriteLine("The number of columns in the file does not match the number of keys in the dictionary.");
                    }

                    //var noOfRow = sheet.Dimension.End.Row;

                    //if (noOfCol != orderedHeader.Count())
                    //    throw new Exception("The uploaded template does not match the accepted template");

                    if (noOfRow < 1)
                        throw new Exception("Empty file template was uploaded!");

                    if (noOfRow > maxNumberOfTransactionRows + 1)
                        throw new Exception($"The uploaded template contains too many transactions. Maximum allowed is {maxNumberOfTransactionRows} records");

                    // Get a list of column indexes for required fields
                    List<int> requiredColumnIndexes = new List<int>();
                    List<string> requiredColumns = new List<string>();

                    // Loop through the dictionary and find the indexes of required fields
                    foreach (var field in fieldRequirements)
                    {
                        if (field.Value == "Required")
                        {
                            requiredColumns.Add(field.Key);  // Store the column name
                        }
                    }

                    // Get the column headers from the sheet (assumed to be in row 6)
                    List<string> columnHeaders = new List<string>();
                    for (int column = 1; column <= noOfCol; column++)
                    {
                        var cellValue = sheet.Cells[6, column].Text;  // Get header text from the 6th row
                        columnHeaders.Add(cellValue.Replace("\n", " ").Trim());
                    }

                    // Find the index of each required field column in the sheet
                    foreach (var requiredColumn in requiredColumns)
                    {
                        int columnIndex = columnHeaders.IndexOf(requiredColumn) + 1; // +1 because EPPlus is 1-based indexing
                        if (columnIndex > 0) // If the column is found in the header row
                        {
                            requiredColumnIndexes.Add(columnIndex);  // Store the column index for required fields
                        }
                    }

                    // Loop through the rows and check for missing values in required fields
                    for (int row = 7; row <= noOfRow; row++)  // Starting from row 7 (data starts after row 6)
                    {
                        foreach (var columnIndex in requiredColumnIndexes)
                        {
                            var cellValue = sheet.Cells[row, columnIndex].Text.Trim();  // Get cell value in the required column
                            if (string.IsNullOrWhiteSpace(cellValue))
                            {
                                // If the cell is empty and the column is required, log an error
                                Console.WriteLine($"Error: Missing value in required column '{columnHeaders[columnIndex - 1]}' for row {row}.");
                            }
                        }
                    }

                }
            }
        }


        public static byte[] CreateExcelFromFailedRecords(List<string> failedRecords)
        {
                // Create a memory stream to store the Excel file
                using (var memoryStream = new MemoryStream())
                {
                    // Initialize EPPlus to create the Excel file
                    using (var package = new ExcelPackage(memoryStream))
                    {
                        // Add a worksheet to the Excel file
                        var worksheet = package.Workbook.Worksheets.Add("Failed Records");

                        // Add headers to the worksheet (assuming the structure of failed record details)
                        worksheet.Cells[1, 1].Value = "Row Position";
                        worksheet.Cells[1, 2].Value = "First Name";
                        worksheet.Cells[1, 3].Value = "Last Name";
                        worksheet.Cells[1, 4].Value = "A Number";
                        worksheet.Cells[1, 5].Value = "Error";

                        // Fill the worksheet with the failed records
                        int row = 2; // Start from the second row (because the first row is headers)
                        foreach (var record in failedRecords)
                        {
                            var recordParts = record.Split(',');  // Split each failed record into parts
                            for (int col = 0; col < recordParts.Length; col++)
                            {
                                worksheet.Cells[row, col + 1].Value = recordParts[col].Trim(); // Write to the cells
                            }
                            row++;
                        }

                        // Save the package to the memory stream
                        package.Save();
                    }

                    // Return the byte array of the Excel file
                    return memoryStream.ToArray();
                }
        }
        

    }
}
