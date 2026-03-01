using ClosedXML.Excel;
using Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace BusinessLogicLayer
{
    /// <summary>
    /// Service til at læse og skrive rapporter fra/til Excel-filer.
    /// </summary>
    public class ExcelService
    {
        /// <summary>
        /// Standardkonstruktør.
        /// </summary>
        public ExcelService()
        {
        }

        /// <summary>
        /// Læser de første 200 rapporter fra en Excel-fil.
        /// </summary>
        /// <param name="excelFilePath">Stien til Excel-filen.</param>
        /// <param name="idColumn">Kolonne med BRNumber.</param>
        /// <param name="primaryColumn">Kolonne med PrimaryUrl.</param>
        /// <param name="fallbackColumn">Kolonne med FallbackUrl.</param>
        /// <param name="yearColumn">Kolonne med årstal.</param>
        /// <returns>Liste af op til 200 rapporter.</returns>
        public List<Report> ReadFirstTwoHundredReports(string excelFilePath, string idColumn, string primaryColumn, string fallbackColumn, string yearColumn)
        {
            using var stream = File.OpenRead(excelFilePath);
            return ReadFirstTwoHundredReports(stream, idColumn, primaryColumn, fallbackColumn, yearColumn);
        }

        /// <summary>
        /// Læser de første 200 rapporter fra et Excel-stream.
        /// </summary>
        public List<Report> ReadFirstTwoHundredReports(Stream excelStream, string idColumn, string primaryColumn, string fallbackColumn, string yearColumn)
        {
            List<Report> reportList = new List<Report>();

            using var workbook = new XLWorkbook(excelStream);
            var sheet = workbook.Worksheet(1);

            int count = 0;
            foreach (var row in sheet.RowsUsed().Skip(1)) // Spring header-række over
            {
                string brNumber = row.Cell(idColumn).GetString();
                string primaryUrl = row.Cell(primaryColumn).GetString();
                string fallbackUrl = row.Cell(fallbackColumn).GetString();
                string yearString = row.Cell(yearColumn).GetString();

                int.TryParse(yearString, out int year); // Gem år, hvis muligt

                reportList.Add(new Report
                {
                    BRNumber = brNumber,
                    PrimaryUrl = string.IsNullOrWhiteSpace(primaryUrl) ? null : primaryUrl,
                    FallbackUrl = string.IsNullOrWhiteSpace(fallbackUrl) ? null : fallbackUrl,
                    Year = year,
                    Status = StatusMessage.NotDownloaded
                });

                count++;
                if (count >= 200) break; // Stop efter 200
            }

            return reportList;
        }

        /// <summary>
        /// Læser alle rapporter fra en Excel-fil.
        /// </summary>
        public List<Report> ReadReports(string excelFilePath, string idColumn, string primaryColumn, string fallbackColumn, string yearColumn)
        {
            using var stream = File.OpenRead(excelFilePath);
            return ReadReports(stream, idColumn, primaryColumn, fallbackColumn, yearColumn);
        }

        /// <summary>
        /// Læser alle rapporter fra et Excel-stream.
        /// </summary>
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

        /// <summary>
        /// Skriver rapporter til en Excel-fil på disk.
        /// </summary>
        public void WriteReports(List<Report> reports, string outputPath)
        {
            using var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
            WriteReports(reports, stream);
        }

        /// <summary>
        /// Skriver rapporter til et Excel-stream.
        /// </summary>
        public void WriteReports(List<Report> reports, Stream stream)
        {
            using var workbook = new XLWorkbook();
            var sheet = workbook.Worksheets.Add("Reports");

            // Opret header-række
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
                // Skriv data-rækker
                sheet.Cell(row, 1).Value = report.BRNumber;
                sheet.Cell(row, 2).Value = report.PrimaryUrl;
                sheet.Cell(row, 3).Value = report.FallbackUrl;
                sheet.Cell(row, 4).Value = report.Status.ToString();
                sheet.Cell(row, 5).Value = report.LocalPath;
                sheet.Cell(row, 6).Value = report.FileSizeKB;
                sheet.Cell(row, 7).Value = report.DownloadTimeSeconds;
                row++;
            }

            workbook.SaveAs(stream); // Gem til stream
            stream.Flush();
        }

        /// <summary>
        /// Tjekker om en input-fil kan bruges (eksisterer, sti er gyldig, ikke låst).
        /// </summary>
        public bool ValidateInputFile(string path)
        {
            if (!IsValidPath(path)) return false;
            if (!File.Exists(path)) return false;
            if (IsFileLocked(path)) return false;

            return true;
        }

        /// <summary>
        /// Tjekker om en output-fil kan skrives til.
        /// </summary>
        public bool ValidateOutputFile(string path)
        {
            if (!IsValidPath(path)) return false;
            if (File.Exists(path) && IsFileLocked(path)) return false;
            if (!CanWriteToFile(path)) return false;

            return true;
        }

        /// <summary>
        /// Tjekker om en sti er gyldig.
        /// </summary>
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

        /// <summary>
        /// Tjekker om en fil er låst af et andet program.
        /// </summary>
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

        /// <summary>
        /// Tjekker om en fil kan åbnes til skrivning (Opretter fil hvis den ikke findes).
        /// </summary>
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

        /// <summary>
        /// Validerer at angivne kolonner har data i Excel-filen.
        /// </summary>
        /// <param name="excelFilePath">Sti til Excel-fil.</param>
        /// <param name="columns">Liste af kolonner, f.eks. "A", "B".</param>
        public bool ValidateColumns(string excelFilePath, params string[] columns)
        {
            using var stream = File.OpenRead(excelFilePath);
            return ValidateColumns(stream, columns);
        }

        /// <summary>
        /// Validerer at angivne kolonner har data i Excel-stream.
        /// </summary>
        public bool ValidateColumns(Stream excelStream, params string[] columns)
        {
            using var workbook = new XLWorkbook(excelStream);
            var sheet = workbook.Worksheet(1);

            foreach (var column in columns)
            {
                // RowsUsed ignorerer tomme rækker, men vi tjekker data i alle brugte rækker
                var rows = sheet.RowsUsed().Skip(1);
                bool hasData = rows.Any(r => !string.IsNullOrWhiteSpace(r.Cell(column).GetString()));
                if (!hasData) return false;
            }

            return true;
        }
    }
}