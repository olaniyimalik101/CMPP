using Interswitch.CRM.SharedFunctions.DataModels;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace Interswitch.CRM.CustomerPortal.Helpers
{
    public class FileHelper
    {
        public static void ValidateFileHeaderContentSize(HttpPostedFileBase file, TransactionType transactionType, int maxNumberOfTransactionRows = 1000)
        {
            if (transactionType == null)
                throw new Exception("Invalid transaction type selected.");

            List<string> orderedHeader = transactionType.TemplateHeaders.Split(',').ToList();
            if (file == null || orderedHeader == null || orderedHeader.Count < 1)
                throw new Exception("Invalid uploaded file parameters.");

            string fileName = file.FileName;
            string fileContentType = file.ContentType;
            byte[] fileBytes = new byte[file.ContentLength];

            using (var package = new ExcelPackage(file.InputStream))
            {
                var sheet = package.Workbook.Worksheets.First();
                var noOfCol = sheet.Dimension.End.Column;
                var noOfRow = sheet.Dimension.End.Row;

                if (noOfCol != orderedHeader.Count())
                    throw new Exception("The uploaded template does not match the selected transaction type.");

                if (noOfRow < 2)
                    throw new Exception("Empty transaction template was uploaded!");

                if (noOfRow > maxNumberOfTransactionRows + 1)
                    throw new Exception("The uploaded template contains too many transaction records. Maximum allowed is " + maxNumberOfTransactionRows + " records");

                for (int columnPosition = 0; columnPosition < orderedHeader.Count(); columnPosition++)
                {
                    var cellValue = sheet.Cells[1, columnPosition + 1].Value.ToString();
                    //var abc = orderedHeader[columnPosition];
                    if (!cellValue.Equals(orderedHeader[columnPosition]))
                        throw new Exception("Invalid column Header " + cellValue + " was found in the upload template. Upload template for the selected transaction type.");
                }

                // Validate all uploaded rows for required fields.
                // The fields are ordered on the databse (Ordering Column on the 'TransactionTypeFormFieldValidation' table ) such that the required fields come first so only the first set of columns are validated.
                int numberOfRequiredFields = transactionType.TransactionTypeFormFieldValidations.Where(a => a.IsRequired).Count();
                bool hasError = false;
                for (int rowPosition = 2; rowPosition <= noOfRow; rowPosition++)
                {
                    for (int columnPosition = 1; columnPosition <= numberOfRequiredFields; columnPosition++)
                    {
                        var value = sheet.Cells[rowPosition, columnPosition].Value;
                        if (value == null)
                        {
                            var columnHeader = sheet.Cells[1, columnPosition].Value.ToString();
                            throw new Exception("The column '" + columnHeader + "' is required. Kindly complete the uploaded template and try again.");
                        }
                    }
                }

                file.InputStream.Seek(0, System.IO.SeekOrigin.Begin);
            }
        }

        public static void ValidateFileHeaderForRequestSend(HttpPostedFileBase file, int maxNumberOfTransactionRows = 1000)
        {
            if (file == null)
                throw new Exception("Invalid uploaded file parameters.");

            int orderedHeader = 1;


            string fileName = file.FileName;
            string fileContentType = file.ContentType;
            byte[] fileBytes = new byte[file.ContentLength];

            using (var package = new ExcelPackage(file.InputStream))
            {
                var sheet = package.Workbook.Worksheets.First();
                var noOfCol = sheet.Dimension.End.Column;
                var noOfRow = sheet.Dimension.End.Row;

                if (noOfCol != orderedHeader)
                    throw new Exception("The uploaded template does not match the selected issue type.");

                if (noOfRow < 2)
                    throw new Exception("Empty attachment template was uploaded!");

                if (noOfRow < 3)
                    throw new Exception("single log code was uploaded!. Please use the single option from the log code volume List");

                if (noOfRow > maxNumberOfTransactionRows + 1)
                    throw new Exception("The uploaded template contains too many Log codes. Maximum allowed is " + maxNumberOfTransactionRows + " records");

                file.InputStream.Seek(0, System.IO.SeekOrigin.Begin);
            }
        }

        public static void ValidateFileHeaderForSelectedContentSize(HttpPostedFileBase file, string Issuetype, int maxNumberOfTransactionRows = 1000)
        {
            if (file == null)
                throw new Exception("Invalid uploaded file parameters.");

            var _header = string.Empty;
            int numberOfRequiredFields = 0;

            if (Issuetype == "UnavailableonArbiter")
            {
                _header = "First Six Digits and Last Four Digit of Maskedpan,STAN,RRN,Terminal ID,Amount,Date";
                numberOfRequiredFields = 6;
            }
            else if (Issuetype == "transactionSettlementDateInq")
            {
                _header = "Channel Transaction Reference No,Amount,Date";

                numberOfRequiredFields = 3;
            }
            else if (Issuetype == "requestForReport")
            {
                _header = "Termnal ID,Start Date,End Date";

                numberOfRequiredFields = 3;
            }

            //var header_ = "First Six Digits and Last Four Digit of Maskedpan,STAN,RRN,Terminal ID,Amount,Date";

            List<string> orderedHeader = _header.Split(',').ToList();
            //List<string> orderedHeader = list;
            //int orderedHeader = 6;

            string fileName = file.FileName;
            string fileContentType = file.ContentType;
            byte[] fileBytes = new byte[file.ContentLength];

            using (var package = new ExcelPackage(file.InputStream))
            {
                var sheet = package.Workbook.Worksheets.First();
                var noOfCol = sheet.Dimension.End.Column;
                var noOfRow = sheet.Dimension.End.Row;

                if (noOfCol != orderedHeader.Count())
                    throw new Exception("The uploaded template does not match the selected issue type.");

                if (noOfRow < 2)
                    throw new Exception("Empty attachment template was uploaded!");

                if (noOfRow < 3)
                    throw new Exception("single transaction info was uploaded!. Please use the single option from the Volume/issue List");

                if (noOfRow > maxNumberOfTransactionRows + 1)
                    throw new Exception("The uploaded template contains too many transactions. Maximum allowed is " + maxNumberOfTransactionRows + " records");

                for (int columnPosition = 0; columnPosition < orderedHeader.Count(); columnPosition++)
                {
                    var cellValue = sheet.Cells[1, columnPosition + 1].Value.ToString();
                    //var abc = orderedHeader[columnPosition];
                    if (!cellValue.Equals(orderedHeader[columnPosition]))
                        throw new Exception("Invalid column Header " + cellValue + " was found in the upload template. Upload template for the selected transaction type.");
                }

                // Validate all uploaded rows for required fields.
                // The fields are ordered on the databse (Ordering Column on the 'TransactionTypeFormFieldValidation' table ) such that the required fields come first so only the first set of columns are validated.
                //int numberOfRequiredFields = 6;
                bool hasError = false;
                for (int rowPosition = 2; rowPosition <= noOfRow; rowPosition++)
                {
                    for (int columnPosition = 1; columnPosition <= numberOfRequiredFields; columnPosition++)
                    {
                        var value = sheet.Cells[rowPosition, columnPosition].Value;
                        if (value == null)
                        {
                            var columnHeader = sheet.Cells[1, columnPosition].Value.ToString();
                            throw new Exception("The column '" + columnHeader + "' is required. Kindly complete the uploaded template and try again.");
                        }
                    }
                }

                file.InputStream.Seek(0, System.IO.SeekOrigin.Begin);
            }
        }

    }
}