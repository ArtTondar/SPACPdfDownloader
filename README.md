# PDF Report Downloader

Et C# konsolprojekt til at læse rapporter fra en Excel-fil, downloade PDF-filer fra URLs og gemme resultaterne i en output Excel-fil. 
Projektet håndterer både primære og fallback URLs, og logger downloadstatus samt filstatistik.

---

## Projektstruktur

ProjectRoot/
   Program.cs # Hovedprogrammet, starter processen
   Models/
     Report.cs # Report-klassen og StatusMessage enum
   BusinessLogicLayer/
     ExcelService.cs # Læser, skriver og validerer Excel-filer
     PdfDownloaderService.cs # Downloader PDF-filer og logger status
   README.md # Denne fil

---

### Klasseoversigt

- **ExcelService**
  - Læser rapportdata fra Excel (`ReadReports`, `ReadFirstTwoHundredReports`)
  - Skriver rapportdata til Excel (`WriteReports`)
  - Validerer input/output filer og nødvendige kolonner (`ValidateInputFile`, `ValidateOutputFile`, `ValidateColumns`)

- **PdfDownloaderService**
  - Downloader PDF-filer asynkront (`DownloadReportsAsync`)
  - Håndterer både primære og fallback URLs
  - Logger succes, fejl og tidsforbrug

- **Report**
  - Indeholder metadata for hver rapport: BRNumber, URLs, Year, Status, LocalPath, filstørrelse og downloadtid
  - StatusMessage enum: `NotDownloaded`, `Downloaded`, `Failed`

- **Program.cs**
  - Starter applikationen
  - Validerer input Excel og output-sti
  - Læser de første 200 rapporter fra Excel
  - Downloader PDF-filer
  - Gemmer opdaterede rapporter tilbage til Excel

---

## Krav

- .NET 6.0 eller nyere
- NuGet pakker:
  - `EPPlus` (eller anden Excel-læse/skrivebibliotek)
  - `System.Net.Http` (standard for HttpClient)

---

## Sådan køres projektet

1. **Opdater hardcoded stier og kolonner** i `Program.cs`:
   ```csharp
   string excelInputPath = @"C:\pdf\GRI_2017_2020 (1).xlsx";
   string excelOutputPath = @"C:\pdf\ExcelOutput.xlsx";
   string pdfOutputPath = @"C:\pdf\downloaded_pdf_reports";
   string idColumn = "A";
   string primaryColumn = "AL";
   string fallbackColumn = "AM";
   string yearColumn = "N";

2. Build projektet i Visual Studio eller via CLI:
   dotnet build

3. Kør applikationen:
   dotnet run

4. Output
    Downloadede PDF-filer gemmes i pdfOutputPath
    Opdateret Excel med status og lokal sti gemmes i excelOutputPath

---

## Noter
    Projektet håndterer op til de første 200 rapporter i Excel for hurtig testkørsel (ReadFirstTwoHundredReports).
    Download sker asynkront og rapporterer løbende succes/fejl.
    Hvis en PDF ikke kan downloades fra PrimaryUrl, forsøger systemet FallbackUrl.
    Excel kolonner valideres før læsning for at undgå runtime-fejl.

---

## Fremtidige forbedringer
    Tilføj brugerinput til stier og kolonner i stedet for hardcoding.
    Mulighed for at genoptage tidligere downloads.