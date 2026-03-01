using ClosedXML.Excel;
using Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BusinessLogicLayer
{
    public class ExcelService
    {
        public ExcelService()
        {

        }

        public List<Report> ReadFirstTwoHundredReports(string excelFilePath, string idColumn, string primaryColumn, string fallbackColumn, string yearColumn)
        {
            using var stream = File.OpenRead(excelFilePath);
            return ReadFirstTwoHundredReports(stream, idColumn, primaryColumn, fallbackColumn, yearColumn);
        }

        public List<Report> ReadFirstTwoHundredReports(Stream excelStream, string idColumn, string primaryColumn, string fallbackColumn, string yearColumn)
        {
            List<Report> reportList = new List<Report>();

            using var workbook = new XLWorkbook(excelStream);
            var sheet = workbook.Worksheet(1);

            int count = 0;
            foreach (var row in sheet.RowsUsed().Skip(1))
            {
                string brNumber = row.Cell(idColumn).GetString();
                string primaryUrl = row.Cell(primaryColumn).GetString();
                string fallbackUrl = row.Cell(fallbackColumn).GetString();
                string yearString = row.Cell(yearColumn).GetString();

                int.TryParse(yearString, out int year); // Beholder som i originalen

                reportList.Add(new Report
                {
                    BRNumber = brNumber,
                    PrimaryUrl = string.IsNullOrWhiteSpace(primaryUrl) ? null : primaryUrl,
                    FallbackUrl = string.IsNullOrWhiteSpace(fallbackUrl) ? null : fallbackUrl,
                    Year = year,
                    Status = StatusMessage.NotDownloaded
                });

                count++;
                if (count >= 200) break;
            }

            return reportList;
        }

        public List<Report> ReadReports(string excelFilePath, string idColumn, string primaryColumn, string fallbackColumn, string yearColumn)
        {
            using var stream = File.OpenRead(excelFilePath);
            return ReadReports(stream, idColumn, primaryColumn, fallbackColumn, yearColumn);
        }

        public List<Report> ReadReports(Stream excelStream, string idColumn, string primaryColumn, string fallbackColumn, string yearColumn)
        {
            List<Report> reportList = new List<Report>();

            using var workbook = new XLWorkbook(excelStream);
            var sheet = workbook.Worksheet(1);

            foreach (var row in sheet.RowsUsed().Skip(1))
            {
                string brNumber = row.Cell(idColumn).GetString();
                string primaryUrl = row.Cell(primaryColumn).GetString();
                string fallbackUrl = row.Cell(fallbackColumn).GetString();
                string yearString = row.Cell(yearColumn).GetString();

                int.TryParse(yearString, out int year);

                reportList.Add(new Report
                {
                    BRNumber = brNumber,
                    PrimaryUrl = string.IsNullOrWhiteSpace(primaryUrl) ? null : primaryUrl,
                    FallbackUrl = string.IsNullOrWhiteSpace(fallbackUrl) ? null : fallbackUrl,
                    Year = year,
                    Status = StatusMessage.NotDownloaded
                });
            }

            return reportList;
        }

        public void WriteReports(List<Report> reports, string outputPath)
        {
            using var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
            WriteReports(reports, stream);
        }

        public void WriteReports(List<Report> reports, Stream stream)
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

            workbook.SaveAs(stream);
            stream.Flush();
        }

        public bool ValidateInputFile(string path)
        {
            if (!IsValidPath(path)) return false;
            if (!File.Exists(path)) return false;
            if (IsFileLocked(path)) return false;

            return true;
        }

        public bool ValidateOutputFile(string path)
        {
            if (!IsValidPath(path)) return false;
            if (File.Exists(path) && IsFileLocked(path)) return false;
            if (!CanWriteToFile(path)) return false;

            return true;
        }

        public bool IsValidPath(string path)
        {
            try
            {
                Path.GetFullPath(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public bool IsFileLocked(string path)
        {
            try
            {
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return true;
            }
        }

        public bool CanWriteToFile(string path)
        {
            try
            {
                using (var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None))
                {
                    return true;
                }
            }
            catch (IOException)
            {
                return false;
            }
        }

        public bool ValidateColumns(string excelFilePath, params string[] columns)
        {
            using var stream = File.OpenRead(excelFilePath);
            return ValidateColumns(stream, columns);
        }

        public bool ValidateColumns(Stream excelStream, params string[] columns)
        {
            using var workbook = new XLWorkbook(excelStream);
            var sheet = workbook.Worksheet(1);

            foreach (var column in columns)
            {
                var rows = sheet.RowsUsed().Skip(1);
                bool hasData = rows.Any(r => !string.IsNullOrWhiteSpace(r.Cell(column).GetString()));
                if (!hasData) return false;
            }

            return true;
        }
    }
}