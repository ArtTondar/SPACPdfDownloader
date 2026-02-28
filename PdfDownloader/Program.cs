using BusinessLogicLayer;
using Models;

ExcelService es = new ExcelService();
PdfDownloaderService ds = new PdfDownloaderService();

// Hardcoded Excel-fil og kolonner - på sigt ændres til brugerinput
string excelInputPath = @"C:\pdf\GRI_2017_2020 (1).xlsx";
string excelOutputPath = @"C:\pdf\ExcelOutput.xlsx";
string pdfOutputPath = @"C:\pdf\downloaded_pdf_reports";
string idColumn = "A"; // kolonne for BRNummer
string primaryColumn = "AL";   // kolonne for PrimaryUrl
string fallbackColumn = "AM";  // kolonne for FallbackUrl
string yearColumn = "N"; // kolonne for publication year

// Validate excel input file
if (!es.ValidateInputFile(excelInputPath))
{
    Console.WriteLine("Input fil er ugyldig, findes ikke eller er låst.");
    return;
}

// Validate excel output file
if (!es.ValidateOutputFile(excelOutputPath))
{
    Console.WriteLine("Output fil er låst eller kan ikke skrives til.");
    return;
}

// Validate columns dynamically
string[] requiredColumns = { primaryColumn, fallbackColumn };
if (!es.ValidateColumns(excelInputPath, requiredColumns))
{
    Console.WriteLine("Input Excel mangler en eller flere nødvendige kolonner.");
    return;
}

// Read reports
List<Report> reports = es.ReadReports(excelInputPath, idColumn, primaryColumn, fallbackColumn, yearColumn);
Console.WriteLine($"Læst {reports.Count} rapporter.");

// Download pdf's
Console.WriteLine("Downloader pdf'er - dette kan godt tage lang tid...");
await ds.DownloadReportsAsync(reports, pdfOutputPath);

// Write reports to output
es.WriteReports(reports, excelOutputPath);
Console.WriteLine($"Rapporter gemt til {excelOutputPath}");