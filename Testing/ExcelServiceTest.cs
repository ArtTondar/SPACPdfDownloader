using BusinessLogicLayer;
using ClosedXML.Excel;
using Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace BusinessLogicLayer.Tests
{
    public class ExcelServiceTests : IDisposable
    {
        private readonly List<string> _tempFiles = new();

        public void Dispose()
        {
            foreach (string file in _tempFiles)
            {
                try { File.Delete(file); } catch { }
            }
        }

        #region Read/Write Tests

        [Fact]
        public void WriteReports_CreatesValidExcel()
        {
            List<Report> reports = new List<Report>
    {
        new Report { BRNumber = "BR1", PrimaryUrl = "http://a", FallbackUrl = "http://b", Year = 2024, Status = StatusMessage.NotDownloaded },
        new Report { BRNumber = "BR2", PrimaryUrl = "http://c", FallbackUrl = null, Year = 2023, Status = StatusMessage.NotDownloaded }
    };

            ExcelService service = new ExcelService();
            using MemoryStream stream = new MemoryStream();

            service.WriteReports(reports, stream);

            stream.Position = 0; // vigtigt!

            using XLWorkbook workbook = new XLWorkbook(stream);
            IXLWorksheet sheet = workbook.Worksheet(1);

            // Læs de kolonner, som faktisk er i WriteReports
            Assert.Equal("BR1", sheet.Cell(2, 1).GetString());      // BRNumber
            Assert.Equal("http://a", sheet.Cell(2, 2).GetString()); // PrimaryUrl
            Assert.Equal("http://b", sheet.Cell(2, 3).GetString()); // FallbackUrl
            Assert.Equal("NotDownloaded", sheet.Cell(2, 4).GetString()); // Status

            Assert.Equal("BR2", sheet.Cell(3, 1).GetString());      // BRNumber
            Assert.Equal("http://c", sheet.Cell(3, 2).GetString()); // PrimaryUrl
            Assert.Equal("", sheet.Cell(3, 3).GetString());         // FallbackUrl
            Assert.Equal("NotDownloaded", sheet.Cell(3, 4).GetString()); // Status
        }

        [Fact]
        public void ReadReports_ReadsCorrectly()
        {
            // Arrange
            using MemoryStream stream = new MemoryStream();
            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet sheet = workbook.Worksheets.Add("Reports");
                sheet.Cell(1, 1).Value = "BRNumber";
                sheet.Cell(1, 2).Value = "PrimaryUrl";
                sheet.Cell(1, 3).Value = "FallbackUrl";
                sheet.Cell(1, 4).Value = "Year";

                sheet.Cell(2, 1).Value = "BR1";
                sheet.Cell(2, 2).Value = "http://a";
                sheet.Cell(2, 3).Value = "http://b";
                sheet.Cell(2, 4).Value = "2024";

                sheet.Cell(3, 1).Value = "BR2";
                sheet.Cell(3, 2).Value = "http://c";
                sheet.Cell(3, 3).Value = "";
                sheet.Cell(3, 4).Value = "2023";

                workbook.SaveAs(stream);
            }

            stream.Position = 0;
            ExcelService service = new ExcelService();

            // Act
            List<Report> reports = service.ReadReports(stream, "BRNumber", "PrimaryUrl", "FallbackUrl", "Year");

            // Assert
            Assert.Equal(2, reports.Count);
            Assert.Equal("BR1", reports[0].BRNumber);
            Assert.Equal("http://a", reports[0].PrimaryUrl);
            Assert.Equal("http://b", reports[0].FallbackUrl);
            Assert.Equal(2024, reports[0].Year);

            Assert.Equal("BR2", reports[1].BRNumber);
            Assert.Equal("http://c", reports[1].PrimaryUrl);
            Assert.Null(reports[1].FallbackUrl);
            Assert.Equal(2023, reports[1].Year);
        }

        [Fact]
        public void ReadFirstFiftyReports_LimitsTo50()
        {
            // Arrange
            using MemoryStream stream = new MemoryStream();
            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet sheet = workbook.Worksheets.Add("Reports");
                sheet.Cell(1, 1).Value = "BRNumber";
                sheet.Cell(1, 2).Value = "PrimaryUrl";
                sheet.Cell(1, 3).Value = "FallbackUrl";
                sheet.Cell(1, 4).Value = "Year";

                for (int i = 1; i <= 250; i++)
                {
                    sheet.Cell(i + 1, 1).Value = $"BR{i}";
                    sheet.Cell(i + 1, 2).Value = $"http://primary{i}";
                    sheet.Cell(i + 1, 3).Value = $"http://fallback{i}";
                    sheet.Cell(i + 1, 4).Value = "2024";
                }

                workbook.SaveAs(stream);
            }

            stream.Position = 0;
            ExcelService service = new ExcelService();

            // Act
            List<Report> reports = service.ReadFirstFiftyReports(stream, "BRNumber", "PrimaryUrl", "FallbackUrl", "Year");

            // Assert
            Assert.Equal(50, reports.Count);
            Assert.Equal("BR1", reports[0].BRNumber);
            Assert.Equal("BR50", reports[49].BRNumber);
        }

        #endregion

        #region File Validation Tests

        [Fact]
        public void IsValidPath_ReturnsTrueForValidPath()
        {
            var service = new ExcelService();
            Assert.True(service.IsValidPath(Path.GetTempPath()));
        }

        [Fact]
        public void IsFileLocked_ReturnsTrue_WhenFileIsOpen()
        {
            string tempFile = Path.GetTempFileName();
            _tempFiles.Add(tempFile);

            using var fs = new FileStream(tempFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            var service = new ExcelService();

            bool locked = service.IsFileLocked(tempFile);
            Assert.True(locked);
        }

        [Fact]
        public void CanWriteToFile_ReturnsTrue_WhenFileWritable()
        {
            string tempFile = Path.GetTempFileName();
            _tempFiles.Add(tempFile);

            var service = new ExcelService();
            bool canWrite = service.CanWriteToFile(tempFile);
            Assert.True(canWrite);
        }

        [Fact]
        public void ValidateInputFile_ReturnsFalse_WhenFileDoesNotExist()
        {
            var service = new ExcelService();
            bool result = service.ValidateInputFile("C:\\nonexistentfile.xlsx");
            Assert.False(result);
        }

        [Fact]
        public void ValidateOutputFile_ReturnsFalse_WhenFileIsLocked()
        {
            string tempFile = Path.GetTempFileName();
            _tempFiles.Add(tempFile);

            using var fs = new FileStream(tempFile, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            var service = new ExcelService();

            bool result = service.ValidateOutputFile(tempFile);
            Assert.False(result);
        }

        #endregion

        #region Column Validation Tests

        [Fact]
        public void ValidateColumns_ReturnsTrue_WhenColumnsHaveData()
        {
            using MemoryStream stream = new MemoryStream();
            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet sheet = workbook.Worksheets.Add("Reports");
                sheet.Cell(1, 1).Value = "BRNumber";
                sheet.Cell(1, 2).Value = "PrimaryUrl";
                sheet.Cell(2, 1).Value = "BR1";
                sheet.Cell(2, 2).Value = "http://a";

                workbook.SaveAs(stream);
            }

            stream.Position = 0;
            ExcelService service = new ExcelService();

            bool result = service.ValidateColumns(stream, "BRNumber", "PrimaryUrl");
            Assert.True(result);
        }

        [Fact]
        public void ValidateColumns_ReturnsFalse_WhenColumnEmpty()
        {
            using MemoryStream stream = new MemoryStream();
            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet sheet = workbook.Worksheets.Add("Reports");
                sheet.Cell(1, 1).Value = "BRNumber";
                sheet.Cell(1, 2).Value = "PrimaryUrl";
                sheet.Cell(2, 1).Value = "";
                sheet.Cell(2, 2).Value = "";

                workbook.SaveAs(stream);
            }

            stream.Position = 0;
            ExcelService service = new ExcelService();

            bool result = service.ValidateColumns(stream, "BRNumber", "PrimaryUrl");
            Assert.False(result);
        }

        #endregion
    }
}