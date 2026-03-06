using BusinessLogicLayer;
using Models;
using Microsoft.Extensions.Configuration;

// ------------------------------------------------------------
// Læs konfiguration fra appsettings.json
// ------------------------------------------------------------
// ConfigurationBuilder henter værdier fra JSON-filen.
// optional: false → programmet fejler, hvis filen ikke findes
var config = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory()) // Base path = kørselens folder
    .AddJsonFile("appsettings.json", optional: false)
    .Build();

// ------------------------------------------------------------
// Opret services
// ------------------------------------------------------------
// ExcelService håndterer læsning og skrivning af Excel-filer
// PdfDownloaderService håndterer PDF-downloads via HttpClient
ExcelService es = new ExcelService();
PdfDownloaderService ds = new PdfDownloaderService(new HttpClient());

// ------------------------------------------------------------
// Læs paths og kolonnenavne fra konfiguration
// ------------------------------------------------------------
string excelInputPath = config["ExcelInputPath"];
string excelOutputPath = config["ExcelOutputPath"];
string pdfOutputPath = config["PdfOutputPath"];

// Kolonnenavne i header-rækken (ikke bogstaver)
// Gør det muligt at skifte inputfiler uden at ændre koden
string idColumn = config["Columns:BRNumber"];
string primaryColumn = config["Columns:PrimaryUrl"];
string fallbackColumn = config["Columns:FallbackUrl"];
string yearColumn = config["Columns:Year"];

// Maks antal samtidige downloads (dynamisk via JSON)
int maxParallel = int.Parse(config["MaxParallelDownloads"]);

// ------------------------------------------------------------
// Valider input- og outputfiler
// ------------------------------------------------------------
if (!es.ValidateInputFile(excelInputPath))
{
    Console.WriteLine("Input fil er ugyldig, findes ikke eller er låst.");
    return;
}

if (!es.ValidateOutputFile(excelOutputPath))
{
    Console.WriteLine("Output fil er låst eller kan ikke skrives til.");
    return;
}

// ------------------------------------------------------------
// Valider at nødvendige kolonner findes og har data
// ------------------------------------------------------------
string[] requiredColumns = { idColumn, primaryColumn, fallbackColumn };
if (!es.ValidateColumnsHasData(excelInputPath, requiredColumns))
{
    Console.WriteLine("Input Excel mangler en eller flere nødvendige kolonner.");
    return;
}

// ------------------------------------------------------------
// Læs rapporter fra Excel
// ------------------------------------------------------------
// Læs kun de første 200 rapporter i prototypen
// Dette kan senere ændres til at læse alle (ReadReports)
List<Report> reports = es.ReadFirstFiftyReports(
    excelInputPath,
    idColumn,
    primaryColumn,
    fallbackColumn,
    yearColumn
);

Console.WriteLine($"Læst {reports.Count} rapporter.");

// ------------------------------------------------------------
// Download PDF'er
// ------------------------------------------------------------
// Dynamisk antal samtidige downloads via maxParallel
// DownloadReportsAsync håndterer automatisk fallback URLs og statusopdatering
Console.WriteLine("Downloader pdf'er - dette kan godt tage lang tid...");
await ds.DownloadReportsAsync(reports, pdfOutputPath, maxParallel);

// ------------------------------------------------------------
// Skriv rapporter tilbage til Excel
// ------------------------------------------------------------
es.WriteReports(reports, excelOutputPath);
Console.WriteLine($"Rapporter gemt til {excelOutputPath}");