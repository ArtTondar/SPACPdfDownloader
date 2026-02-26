using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace BusinessLogicLayer
{
    public class ExcelService
    {
        public ExcelService()
        {

        }

        public List<Report> ReadFirstTwentyReports(string excelFilePath, string primaryColumn, string fallbackColumn)
        {
            List<Report> reportList = new List<Report>();

            using var workbook = new XLWorkbook(excelFilePath);
            var sheet = workbook.Worksheet(1);

            int count = 0;
            foreach (var row in sheet.RowsUsed().Skip(1)) // spring header over
            {
                string primaryUrl = row.Cell(primaryColumn).GetString();
                string fallbackUrl = row.Cell(fallbackColumn).GetString();

                var report = new Report()
                {
                    PrimaryUrl = string.IsNullOrWhiteSpace(primaryUrl) ? null : primaryUrl,
                    FallbackUrl = string.IsNullOrWhiteSpace(fallbackUrl) ? null : fallbackUrl,
                    Status = StatusMessage.NotDownloaded
                };

                reportList.Add(report);

                count++;
                if (count >= 20) break; // stop efter 10
            }

            return reportList;
        }

        public List<Report> ReadReports(string excelFilePath, string primaryColumn, string fallbackColumn)
        {
            List<Report> reportList = new List<Report>();

            using var workbook = new XLWorkbook(excelFilePath);
            var sheet = workbook.Worksheet(1);

            foreach (var row in sheet.RowsUsed().Skip(1))
            {
                string primaryUrl = row.Cell(primaryColumn).GetString();
                string fallbackUrl = row.Cell(fallbackColumn).GetString();

                var report = new Report()
                {
                    PrimaryUrl = string.IsNullOrWhiteSpace(primaryUrl) ? null : primaryUrl,
                    FallbackUrl = string.IsNullOrWhiteSpace(fallbackUrl) ? null : fallbackUrl,
                    Status = StatusMessage.NotDownloaded
                };

                reportList.Add(report);
            }

            return reportList;
        }

        public void WriteReports(List<Report> reports, string outputPath)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Reports");

            // Header
            sheet.Cell(1, 1).Value = "BRNumber";
            sheet.Cell(1, 2).Value = "PrimaryUrl";
            sheet.Cell(1, 3).Value = "FallbackUrl";
            sheet.Cell(1, 4).Value = "Status";
            sheet.Cell(1, 5).Value = "LocalPath";
            sheet.Cell(1, 6).Value = "FileSizeKB";
            sheet.Cell(1, 7).Value = "DownloadTimeSeconds";

            int row = 2;
            foreach (var report in reports)
            {
                sheet.Cell(row, 1).Value = report.BRNumber;
                sheet.Cell(row, 2).Value = report.PrimaryUrl;
                sheet.Cell(row, 3).Value = report.FallbackUrl;
                sheet.Cell(row, 4).Value = report.Status.ToString();
                sheet.Cell(row, 5).Value = report.LocalPath;
                sheet.Cell(row, 6).Value = report.FileSizeKB;
                sheet.Cell(row, 7).Value = report.DownloadTimeSeconds;
                row++;
            }

            workbook.SaveAs(outputPath);
        }

        public bool ValidateColumns(string excelFilePath, string primaryColumn, string fallbackColumn)
        {
            using var workbook = new XLWorkbook(excelFilePath);
            var sheet = workbook.Worksheet(1);

            var headerRow = sheet.Row(1);
            bool hasPrimary = headerRow.Cells().Any(c => c.Address.ColumnLetter == primaryColumn);
            bool hasFallback = headerRow.Cells().Any(c => c.Address.ColumnLetter == fallbackColumn);

            return hasPrimary && hasFallback;
        }
    }
}
