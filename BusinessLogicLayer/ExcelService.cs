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
    /// Denne version bruger dynamiske kolonnenavne i stedet for hardcodede kolonnebogstaver.
    /// Det gør det muligt at skifte inputfiler uden at ændre koden.
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
        /// Finder kolonneindeks ud fra kolonneheader-navn i første række.
        /// Returnerer null, hvis header ikke findes.
        /// </summary>
        private int? GetColumnIndexByHeader(IXLWorksheet sheet, string headerName)
        {
            var headerRow = sheet.Row(1);

            // Loop gennem alle brugte celler i header-rækken for at finde kolonne
            foreach (var cell in headerRow.CellsUsed())
            {
                // Trim og ignore case for at være robust overfor ekstra whitespace eller små/big letters
                if (cell.GetString().Trim().Equals(headerName, StringComparison.OrdinalIgnoreCase))
                {
                    return cell.Address.ColumnNumber; // Returnér 1-baseret kolonneindeks
                }
            }

            return null; // kolonne ikke fundet
        }

        /// <summary>
        /// Læser de første 200 rapporter fra en Excel-fil på disk.
        /// </summary>
        public List<Report> ReadFirstTwoHundredReports(string excelFilePath, string idColumnName, string primaryColumnName, string fallbackColumnName, string yearColumnName)
        {
            using var stream = File.OpenRead(excelFilePath);
            return ReadFirstTwoHundredReports(stream, idColumnName, primaryColumnName, fallbackColumnName, yearColumnName);
        }

        /// <summary>
        /// Læser de første 200 rapporter fra et Excel-stream.
        /// </summary>
        public List<Report> ReadFirstTwoHundredReports(Stream excelStream, string idColumnName, string primaryColumnName, string fallbackColumnName, string yearColumnName)
        {
            List<Report> reportList = new List<Report>();
            using var workbook = new XLWorkbook(excelStream);
            var sheet = workbook.Worksheet(1);

            // Find kolonneindeks dynamisk ud fra header-navne i stedet for hardcodede bogstaver
            int? idCol = GetColumnIndexByHeader(sheet, idColumnName);
            int? primaryCol = GetColumnIndexByHeader(sheet, primaryColumnName);
            int? fallbackCol = GetColumnIndexByHeader(sheet, fallbackColumnName);
            int? yearCol = string.IsNullOrWhiteSpace(yearColumnName) ? null : GetColumnIndexByHeader(sheet, yearColumnName);

            // Tjek at alle nødvendige kolonner findes
            if (!idCol.HasValue || !primaryCol.HasValue || !fallbackCol.HasValue)
                throw new Exception("En af de nødvendige kolonner findes ikke i inputfilen.");

            int count = 0;

            // Loop gennem alle brugte rækker i arket, spring header-række over
            foreach (var row in sheet.RowsUsed().Skip(1))
            {
                string brNumber = row.Cell(idCol.Value).GetString();
                string primaryUrl = row.Cell(primaryCol.Value).GetString();
                string fallbackUrl = row.Cell(fallbackCol.Value).GetString();
                string yearString = yearCol.HasValue ? row.Cell(yearCol.Value).GetString() : null;

                // Brug TryParse for at håndtere tomme eller ugyldige årstal
                int.TryParse(yearString, out int year);

                // Tilføj rapport til listen
                reportList.Add(new Report
                {
                    BRNumber = brNumber,
                    PrimaryUrl = string.IsNullOrWhiteSpace(primaryUrl) ? null : primaryUrl,
                    FallbackUrl = string.IsNullOrWhiteSpace(fallbackUrl) ? null : fallbackUrl,
                    Year = year,
                    Status = StatusMessage.NotDownloaded
                });

                count++;
                if (count >= 200) break; // Stop efter 200 for prototype
            }

            return reportList;
        }

        /// <summary>
        /// Læser alle rapporter fra en Excel-fil på disk.
        /// </summary>
        public List<Report> ReadReports(string excelFilePath, string idColumnName, string primaryColumnName, string fallbackColumnName, string yearColumnName)
        {
            using var stream = File.OpenRead(excelFilePath);
            return ReadReports(stream, idColumnName, primaryColumnName, fallbackColumnName, yearColumnName);
        }

        /// <summary>
        /// Læser alle rapporter fra et Excel-stream.
        /// </summary>
        public List<Report> ReadReports(Stream excelStream, string idColumnName, string primaryColumnName, string fallbackColumnName, string yearColumnName)
        {
            List<Report> reportList = new List<Report>();
            using var workbook = new XLWorkbook(excelStream);
            var sheet = workbook.Worksheet(1);

            // Find kolonneindeks dynamisk ud fra header-navne
            int? idCol = GetColumnIndexByHeader(sheet, idColumnName);
            int? primaryCol = GetColumnIndexByHeader(sheet, primaryColumnName);
            int? fallbackCol = GetColumnIndexByHeader(sheet, fallbackColumnName);
            int? yearCol = string.IsNullOrWhiteSpace(yearColumnName) ? null : GetColumnIndexByHeader(sheet, yearColumnName);

            // Tjek at alle nødvendige kolonner findes
            if (!idCol.HasValue || !primaryCol.HasValue || !fallbackCol.HasValue)
                throw new Exception("En af de nødvendige kolonner findes ikke i inputfilen.");

            foreach (var row in sheet.RowsUsed().Skip(1))
            {
                string brNumber = row.Cell(idCol.Value).GetString();
                string primaryUrl = row.Cell(primaryCol.Value).GetString();
                string fallbackUrl = row.Cell(fallbackCol.Value).GetString();
                string yearString = yearCol.HasValue ? row.Cell(yearCol.Value).GetString() : null;

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
        /// Tjekker om en fil kan åbnes til skrivning (opretter fil hvis den ikke findes).
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
        /// Validerer at angivne kolonner har data i Excel-filen (header-navne i stedet for bogstaver).
        /// </summary>
        public bool ValidateColumns(string excelFilePath, params string[] columns)
        {
            using var stream = File.OpenRead(excelFilePath);
            return ValidateColumns(stream, columns);
        }

        /// <summary>
        /// Validerer at angivne kolonner har data i Excel-stream.
        /// Finder kolonneindeks dynamisk ud fra header-navn.
        /// </summary>
        public bool ValidateColumns(Stream excelStream, params string[] headerNames)
        {
            using var workbook = new XLWorkbook(excelStream);
            var sheet = workbook.Worksheet(1);

            var headerRow = sheet.Row(1);

            foreach (var headerName in headerNames)
            {
                // Find kolonne ud fra header-navn
                var cell = headerRow.CellsUsed()
                    .FirstOrDefault(c => c.GetString().Trim().Equals(headerName, StringComparison.OrdinalIgnoreCase));

                if (cell == null)
                {
                    // Kolonnen findes ikke → validering fejler
                    return false;
                }

                int colIndex = cell.Address.ColumnNumber;

                // Tjek at der er data i kolonnen
                bool hasData = sheet.RowsUsed().Skip(1)
                    .Any(r => !string.IsNullOrWhiteSpace(r.Cell(colIndex).GetString()));

                if (!hasData) return false;
            }

            return true;
        }
    }
}