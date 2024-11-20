using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Xrm.Sdk;

namespace FileColumnReader
{
    public class FileCreate
    {
        public static void ProcessFileAndCreateRecords(byte[] fileBytes, IOrganizationService crmService, EntityReference baseEntity)
        {
            if (fileBytes == null || fileBytes.Length == 0)
                throw new Exception("Invalid uploaded file parameters.");

            using (var stream = new MemoryStream(fileBytes))
            {
                using (var document = SpreadsheetDocument.Open(stream, false))
                {
                    var workbookPart = document.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    var rows = sheetData.Elements<Row>().ToList();

                    // Get column headers from the 6th row (index 5)
                    var headerRow = rows[5]; // Row 6, 0-indexed
                    List<string> columns = new List<string>();

                    foreach (Cell cell in headerRow.Elements<Cell>())
                    {
                        var cellValue = GetCellValue(cell, workbookPart);
                        if (!string.IsNullOrWhiteSpace(cellValue))
                        {
                            columns.Add(cellValue);
                        }
                    }

                    // Process rows starting from row 7 (index 6)
                    foreach (Row row in rows.Skip(6))  // Skip first 6 rows (header + potential empty rows)
                    {
                        var importRecord = new Entity("importrecord"); // Assuming entity is "importrecord"

                        // Save the data to variables first
                        var walkInDate = convertDateValue(GetCellValueFromRow(row, 2, workbookPart));  // Cell 7,1 -> "firstname"
                        var firstName = GetCellValueFromRow(row, 3, workbookPart);  // Cell 7,2 -> "lastname"
                        var lastname = convertDateValue(GetCellValueFromRow(row, 4, workbookPart));
                        var language = GetCellValueFromRow(row, 5, workbookPart);     // Cell 7,4 -> "email"
                        var otherLanguage = GetCellValueFromRow(row, 6, workbookPart); // Cell 7,6 -> "phonenumber"
                        var email = GetCellValueFromRow(row, 7, workbookPart);        // Cell 7,7 -> "DOB"
                        var cellPhone = GetCellValueFromRow(row, 8, workbookPart);    // Cell 7,8 -> "language"
                        var otherPhone = GetCellValueFromRow(row, 9, workbookPart);     // Cell 7,9 -> "Anumber"
                        var DOB = convertDateValue(GetCellValueFromRow(row, 10, workbookPart));   // Cell 7,11 -> "location"
                        var aNumber = GetCellValueFromRow(row, 11, workbookPart);
                        var familyUnitSize = GetCellValueFromRow(row, 12, workbookPart);// Cell 7,12 -> "familyunitsize"
                        var referral = GetCellValueFromRow(row, 13, workbookPart);  // Cell 7,14 -> "referral"
                        var referredby = GetCellValueFromRow(row, 14, workbookPart);  // Cell 7,15 -> "language" (same as above)
                        var location = GetCellValueFromRow(row, 15, workbookPart);  // Cell 7,14 -> "referral"

                        // Now assign the saved values to the importRecord fields
                        // Assign values to CRM entity if not empty or null
                        if (!string.IsNullOrEmpty(firstName))
                            importRecord["cmpp_name"] = firstName;

                        if (!string.IsNullOrEmpty(lastname))
                            importRecord["cmpp_fullsurname"] = lastname;

                        if (!string.IsNullOrEmpty(email))
                            importRecord["cmpp_email"] = email;

                        if (!string.IsNullOrEmpty(cellPhone))
                            importRecord["cmpp_familyunit"] = cellPhone;

                        if (!string.IsNullOrEmpty(otherPhone))
                            importRecord["cmpp_otherphonenumber"] = otherPhone;

                        if (!string.IsNullOrEmpty(DOB))
                            importRecord["cmpp_dob"] = DOB;

                        if (!string.IsNullOrEmpty(otherLanguage))
                            importRecord["cmpp_otherLanguage"] = otherLanguage;

                        if (!string.IsNullOrEmpty(language))
                            importRecord["cmpp_language"] = language;

                        if (!string.IsNullOrEmpty(aNumber))
                            importRecord["cmpp_anumber"] = aNumber;

                        if (!string.IsNullOrEmpty(location))
                            importRecord["cmpp_location"] = location;

                        if (!string.IsNullOrEmpty(familyUnitSize))
                            importRecord["cmpp_familyunit"] = familyUnitSize;

                        if (!string.IsNullOrEmpty(referral))
                            importRecord["cmpp_referral"] = referral;

                        if (!string.IsNullOrEmpty(referredby))
                            importRecord["cmpp_referredby"] = referredby;

                        if (!string.IsNullOrEmpty(walkInDate))
                            importRecord["cmpp_dateofwalkin"] = walkInDate;

                        importRecord["cmpp_batchimport"] = baseEntity;

                        try
                        {
                            crmService.Create(importRecord); // Create the record in CRM
                        }
                        catch (Exception ex)
                        {
                            var mesg = ex.Message;
                        }
                    }
                }
            }
        }

        // Helper method to get the cell value from a specific cell in a row
        private static string GetCellValueFromRow(Row row, int columnIndex, WorkbookPart workbookPart)
        {
            var cell = row.Elements<Cell>().ElementAtOrDefault(columnIndex - 1); // 0-based index adjustment
            return GetCellValue(cell, workbookPart);
        }

        // Helper method to get the cell value from the Excel file
        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null || cell.CellValue == null)
                return string.Empty;

            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                int index = int.Parse(cell.CellValue.Text);
                var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                return sharedStringTablePart?.SharedStringTable.ElementAt(index)?.InnerText ?? string.Empty;
            }

            return cell.CellValue.Text;
        }

        private static string convertDateValue(string val)
        {
            int serialDate = int.Parse(val);  // The serial date number
            DateTime startDate = new DateTime(1900, 1, 1);  // Excel's start date
            DateTime resultDate = startDate.AddDays(serialDate - 2); // Subtract 2 to account for Excel's leap year bug

            // Format the DateTime to the desired string format
            string formattedDate = resultDate.ToString("M/d/yyyy");

            return formattedDate;
        }
    }
}
