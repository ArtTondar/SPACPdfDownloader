using BusinessLogicLayer;
using Models;

ExcelService excelService = new ExcelService();

// Hardcoded Excel-fil og kolonner - på sigt ændres til brugerinput
string inputPath = @"C:\Users\Jannie\Downloads\GRI_2017_2020 (1).xlsx";
string outputPath = @"C:\Users\Jannie\Downloads\TestExcelOutput2.xlsx";
string idColumn = "A"; // kolonne for BRNummer
string primaryColumn = "AL";   // kolonne for PrimaryUrl
string fallbackColumn = "AM";  // kolonne for FallbackUrl

// 1️ Validate input file
if (!excelService.ValidateInputFile(inputPath))
{
    Console.WriteLine("Input fil er ugyldig, findes ikke eller er låst.");
    return;
}

// 2️ Validate output file
if (!excelService.ValidateOutputFile(outputPath))
{
    Console.WriteLine("Output fil er låst eller kan ikke skrives til.");
    return;
}

// 3️ Validate columns dynamically
string[] requiredColumns = { primaryColumn, fallbackColumn };
if (!excelService.ValidateColumns(inputPath, requiredColumns))
{
    Console.WriteLine("Input Excel mangler en eller flere nødvendige kolonner.");
    return;
}

// 4 Read reports
List<Report> reports = excelService.ReadFirstTwentyReports(inputPath, idColumn, primaryColumn, fallbackColumn);
Console.WriteLine($"Læst {reports.Count} rapporter.");

// 5 Her kan downloades PDF'er, opdatere report.Status, LocalPath, FileSizeKB osv.
foreach (var report in reports)
{
    //dummy data
    report.Status = StatusMessage.Downloaded;
    report.LocalPath = @"C:\Temp\Dummy.pdf";
    report.FileSizeKB = 123;
    report.DownloadTimeSeconds = 2.5;
}

// 6️ Write reports to output
excelService.WriteReports(reports, outputPath);
Console.WriteLine($"Rapporter gemt til {outputPath}");