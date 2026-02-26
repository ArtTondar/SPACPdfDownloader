using BusinessLogicLayer;
using Models;

ExcelService excelService = new ExcelService();

// Hardcoded Excel-fil og kolonner
string excelFilePath = @"C:\Users\Jannie\Downloads\GRI_2017_2020 (1).xlsx";
string primaryColumn = "AM";   // kolonne for PrimaryUrl
string fallbackColumn = "AL";  // kolonne for FallbackUrl

// 1️ Tjek først at kolonnerne findes
bool columnsValid = excelService.ValidateColumns(excelFilePath, primaryColumn, fallbackColumn);
if (!columnsValid)
{
    Console.WriteLine("Error: Excel-file missing required columns.");
    return; // stop programmet
}

// 2️ Læs rapporterne (fx første 10)
List<Report> reports = excelService.ReadFirstTwentyReports(excelFilePath, primaryColumn, fallbackColumn);

Console.WriteLine("Reports loaded from Excel:");

foreach (var report in reports)
{
    Console.WriteLine($"PrimaryUrl: {report.PrimaryUrl}, FallbackUrl: {report.FallbackUrl}, Status: {report.Status}");
}