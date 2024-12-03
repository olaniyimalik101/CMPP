public static validationResponse ValidateFile(byte[] fileBytes, ITracingService logger, int maxNumberOfTransactionRows = 100)
{
    List<string> errorList = new List<string>();

    // Check for null or empty file
    if (fileBytes == null || fileBytes.Length == 0)
    {
        errorList.Add("No content found in the imported file.");
    }

    logger.Trace("Checking File Type");
    // Check if the file is an Excel file (either .xlsx or .xls)
    if (!IsExcelFile(fileBytes))
    {
        errorList.Add("The uploaded file is not an Excel file. Please upload a valid Excel file.");
    }
    logger.Trace("File Type validated");

    // Expected header columns
    var orderedHeader = new List<string>
    {
        "Ct.", "Date of Wallk-In", "First Name", "Full Surname Patrilineal-Matrilineal(as applicable)", "Language  (drop Down)", "Other Language, please specify",
        "Email", "Cell Phone Number", "Other Phone Number", "DOB", "A Number", "Family Unit size", "Referral?", "Referred By", "Location"
    };
    logger.Trace("Obtained total Header as: " + orderedHeader.Count);

    logger.Trace("Opening file with memory stream");
    // Open the byte array as a memory stream and load the document
    using (var memoryStream = new MemoryStream(fileBytes))
    {
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(memoryStream, false))
        {
            if (document == null)
            {
                errorList.Add("Failed to open the Excel file.");
                logger.Trace("Failed to open Excel file.");
                return new validationResponse { IsSuccess = false, failureReason = errorList };
            }

            logger.Trace("Getting the worksheet and sheet data");
            // Get the first worksheet
            var sheet = document.WorkbookPart.Workbook.Sheets.Elements<Sheet>().FirstOrDefault();
            if (sheet == null)
            {
                errorList.Add("No sheet found in the Excel file.");
                return new validationResponse { IsSuccess = false, failureReason = errorList };
            }

            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id);
            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
            if (sheetData == null)
            {
                errorList.Add("No sheet data found in the Excel file.");
                return new validationResponse { IsSuccess = false, failureReason = errorList };
            }

            logger.Trace("Getting the rows and column numbers");
            int noOfCol = 0;

            // Get the header row (6th row, zero-based index is 5)
            var headerRow = sheetData.Elements<Row>().ElementAtOrDefault(5);
            if (headerRow == null)
            {
                errorList.Add("Header row (6th row) not found.");
                return new validationResponse { IsSuccess = false, failureReason = errorList };
            }

            foreach (var cell in headerRow.Elements<Cell>())
            {
                var cellValue = GetCellValue(document, cell);
                if (!string.IsNullOrEmpty(cellValue))
                {
                    noOfCol++;
                }
            }

            logger.Trace("Number of columns is: " + noOfCol);
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

            logger.Trace("Number of rows is: " + noOfRow);

            // Ensure the column count matches the expected header count
            if (noOfCol != orderedHeader.Count)
            {
                errorList.Add("The number of columns do not match with the corresponding number in the template.");
            }
            logger.Trace("Column Number validated");

            // Check if there are any rows with data
            if (noOfRow < 1)
            {
                errorList.Add("No Participant record was found in the file!");
            }

            logger.Trace("Number of rows Validated");

            // Ensure the row count does not exceed the allowed limit
            if (noOfRow > maxNumberOfTransactionRows + 1)
            {
                errorList.Add($"The uploaded template contains too many records. Maximum allowed is {maxNumberOfTransactionRows} records.");
            }

            logger.Trace("Validating same Column Headers as template");
            // Validate that each column header matches the expected header and check for missing columns
            var missingColumns = new List<string>();
            for (int columnPosition = 0; columnPosition < orderedHeader.Count; columnPosition++)
            {
                var cell = headerRow.Elements<Cell>().ElementAtOrDefault(columnPosition);
                if (cell == null || GetCellValue(document, cell).Replace("\n", " ").Trim() != orderedHeader[columnPosition])
                {
                    missingColumns.Add(orderedHeader[columnPosition]);
                }
            }

            logger.Trace("Column header validated, issue with " + missingColumns.Count + " Columns");
            if (missingColumns.Any())
            {
                errorList.Add($"The following columns are missing or incorrect: {string.Join(", ", missingColumns)}");
            }
        }
    }

    logger.Trace("Total number of errors returned is: " + errorList.Count);
    logger.Trace("The following are the errors:\n " + string.Join(", ", errorList));

    // If any validation error exists, return a failure response
    if (errorList.Any())
    {
        return new validationResponse { IsSuccess = false, failureReason = errorList };
    }

    return new validationResponse { IsSuccess = true, failureReason = null };
}
