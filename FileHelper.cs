using DHS.CMPP.Plugins.Helper;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace DHS.CMPP.Plugins2.Helper
{
    public class FileHelper
    {
        public static validationResponse ValidateFile(byte[] fileBytes, int maxNumberOfTransactionRows = 100)
        {
            List<string> errorList = new List<string>();

            // Check for null or empty file
            if (fileBytes == null || fileBytes.Length == 0)
            {
                errorList.Add("No content found in the imported file.");
            }

            // Check if the file is an Excel file (either .xlsx or .xls)
            if (!IsExcelFile(fileBytes))
            {
                errorList.Add("The uploaded file is not an Excel file. Please upload a valid Excel file.");
            }


            // expected header columns
            var orderedHeader = new List<string>
            {
                "Ct.", "Date of Wallk-In", "First Name", "Full Surname Patrilineal-Matrilineal(as applicable)", "Language  (drop Down)", "Other Language, please specify",
                "Email", "Cell Phone Number", "Other Phone Number", "DOB", "A Number", "Family Unit size", "Referral?", "Referred By", "Location"
            };

            // Open the byte array as a memory stream and load the document
            using (var memoryStream = new MemoryStream(fileBytes))
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(memoryStream, false))
                {
                    // Get the first worksheet
                    var sheet = document.WorkbookPart.Workbook.Sheets.Elements<Sheet>().First();
                    var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    int noOfCol = 0;

                    // Get the header row (6th row, zero-based index is 5)
                    var headerRow = sheetData.Elements<Row>().ElementAt(5);
                    foreach (var cell in headerRow.Elements<Cell>())
                    {
                        var cellValue = GetCellValue(document, cell);
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            noOfCol++;
                        }
                    }

                    int noOfRow = 0;

                    // Validate rows starting from row 7 (zero-based index is 6)
                    foreach (var row in sheetData.Elements<Row>().Skip(6))
                    {
                        // Get the cell values for columns 2 to 15 (zero-based indices 1 to 14)
                        var columnValues = row.Elements<Cell>()
                            .Skip(1) // Skip the first column (serial number column)
                            .Take(14 - 1)  // Take the next 14 columns (columns 2 to 15)
                            .Select(cell => GetCellValue(document, cell))  // Get value for each cell
                            .ToList();

                        // Check if at least one column has a non-empty value
                        if (columnValues.Any(value => !string.IsNullOrEmpty(value)))
                        {
                            // Increment row count if at least one of the columns 2-15 has a value
                            noOfRow++;
                        }
                    }



                    // Ensure the column count matches the expected header count
                    if (noOfCol != orderedHeader.Count)
                    {
                        errorList.Add("The number of columns do not match with the corresponding number in the template.");
                    }

                    // Check if there are any rows with data
                    if (noOfRow < 1)
                    {
                        errorList.Add("No Participant record was found in the file!");
                    }

                    // Ensure the row count does not exceed the allowed limit
                    if (noOfRow > maxNumberOfTransactionRows + 1)
                    {
                        errorList.Add($"The uploaded template contains too many records. Maximum allowed is {maxNumberOfTransactionRows} records.");
                    }

                    // Validate that each column header matches the expected header and check for missing columns
                    var missingColumns = new List<string>();
                    for (int columnPosition = 0; columnPosition < orderedHeader.Count; columnPosition++)
                    {
                        var cell = headerRow.Elements<Cell>().ElementAtOrDefault(columnPosition);
                        if (cell == null || GetCellValue(document, cell).Trim() != orderedHeader[columnPosition])
                        {
                            missingColumns.Add(orderedHeader[columnPosition]);
                        }
                    }

                    if (missingColumns.Any())
                    {
                        errorList.Add($"The following columns are missing or incorrect: {string.Join(", ", missingColumns)}");
                    }
                }
            }

            // If any validation error exists, return a failure response
            if (errorList.Any())
            {
                return new validationResponse { IsSuccess = false, failureReason = errorList };
            }

            return new validationResponse { IsSuccess = true, failureReason = null };
        }

        // Helper function to check if the file is an Excel file (either .xlsx or .xls)
        private static bool IsExcelFile(byte[] fileBytes)
        {
            try
            {
                // Read the file header to check for Excel signature
                var fileSignature = fileBytes.Take(4).ToArray();
                // Check for .xlsx (.xml-based format) signature
                var xlsxSignature = new byte[] { 80, 75, 3, 4 }; // PK\x03\x04 for .xlsx
                // Check for .xls (binary format) signature (file starts with D0 CF 11 E0 A1 B1 1A E1)
                var xlsSignature = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };

                // Check if the file is either .xlsx or .xls based on signature
                return fileSignature.SequenceEqual(xlsxSignature) || fileSignature.Take(xlsSignature.Length).SequenceEqual(xlsSignature);
            }
            catch
            {
                return false;
            }
        }

        // Helper function to retrieve the value of a cell, handling shared strings and other types
        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell == null || cell.DataType == null) return string.Empty;

            var value = cell.CellValue?.Text;
            if (cell.DataType == CellValues.SharedString)
            {
                var sharedStringItem = document.WorkbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value));
                return sharedStringItem.Text?.Text ?? sharedStringItem.InnerText;
            }

            return value ?? string.Empty;
        }
    }
}
