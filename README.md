# PDF Report Downloader

Et C# konsolprojekt til at læse rapporter fra en Excel-fil, downloade PDF-filer fra URLs og gemme resultaterne i en output Excel-fil. 
Projektet håndterer både primære og fallback URLs, og logger downloadstatus samt filstatistik.

---

## Projektstruktur

- ProjectRoot/
   - Program.cs # Hovedprogrammet, starter processen
   - appsettings.json
   - README.md # Denne fil
   
   - Models/
     - Report.cs # Report-klassen og StatusMessage enum
     
   - BusinessLogicLayer/
     - ExcelService.cs # Læser, skriver og validerer Excel-filer
     - PdfDownloaderService.cs # Downloader PDF-filer og logger status

---

### Klasseoversigt

- **ExcelService**
  - Læser rapportdata fra Excel (`ReadReports`, `ReadFirstTwoHundredReports`)
  - Skriver rapportdata til Excel (`WriteReports`)
  - Validerer input/output filer og nødvendige kolonner (`ValidateInputFile`, `ValidateOutputFile`, `ValidateColumns`)

- **PdfDownloaderService**
  - Downloader PDF-filer asynkront (`DownloadReportsAsync`)
  - Begrænser parallelle downloads via SemaphoreSlim
  - Logger progress hvert X sekund via PeriodicTimer
  - Logger slutstatus med samlet runtime
  - Måler downloadtid pr. rapport (`DownloadTimeSeconds`)
  - Understøtter fallback URL hvis primær URL fejler

- **Report**
  - Indeholder metadata for hver rapport: `BRNumber`, `PrimaryUrl`, `FallbackUrl`, `Year`, `Status`, `LocalPath`, `FileSizeKB`, `DownloadTimeSeconds`
  - StatusMessage enum: `NotDownloaded`, `Downloaded`, `Failed`

- **Program.cs**
  - Starter applikationen
  - Validerer input Excel og output-stier
  - Læser rapporter fra Excel
  - Starter parallel PDF-download
  - Gemmer opdaterede rapporter tilbage til Excel

---

## Krav

- .NET 6.0 eller nyere
- NuGet pakker:
  - `EPPlus` (eller anden Excel-læse/skrivebibliotek)
  - `System.Net.Http` (standard for HttpClient)

---

## appsettings.json konfiguration

Hvis projektet bruger appsettings.json, skal filen konfigureres korrekt i Visual Studio.
VIGTIGT: Sørg for at filen kopieres til output-mappen
   - Højreklik på appsettings.json
   - Vælg Properties
   - Sæt:
     - Build Action = Content
     - Copy to Output Directory = Copy if newer
Når du bygger/kører projektet, kopieres filen automatisk til: bin\Debug\net6.0\
Hvis dette ikke er sat korrekt, vil konfigurationsindlæsning fejle ved runtime.

--- 

## Sådan køres projektet

1. **Opdater variabler** i `appsettings.json`:
   ```csharp
   "ExcelInputPath": "C:\\pdf\\GRI_2017_2020.xlsx",
   "ExcelOutputPath": "C:\\pdf\\ExcelOutput.xlsx",
   "PdfOutputPath": "C:\\pdf\\downloaded_pdf_reports",
   "Columns": {
     "BRNumber": "BRnum", // Header i Excel for BRNumber
     "PrimaryUrl": "Pdf_URL", // Header i Excel for PrimaryUrl
     "FallbackUrl": "Report Html Address", // Header i Excel for FallbackUrl: null,
     "Year": "Publication Year" // Header i Excel for Year (valgfri)
   },
   "MaxParallelDownloads": 10, // Hvor mange PDF'er der kan hentes samtidigt

2. Build projektet i Visual Studio eller via CLI:
   dotnet build

3. Kør applikationen:
   dotnet run

4. Output
    - Downloadede PDF-filer gemmes i:
      - PdfOutputPath\\{Year}\\{BRNumber}.pdf
    - Den opdaterede Excel-fil indeholder:
      - Status (Downloaded / Failed)
      - Lokal filsti
      - Filstørrelse i KB
      - Downloadtid i sekunder

---

Logging

   - Under kørsel vises:
     - Processed / Total
     - Success / Failed
     - Elapsed time
     - Estimeret resterende tid (ETA)

   Heartbeat logges periodisk via PeriodicTimer.

   - Når alle downloads er færdige, vises en samlet summary:
      Download færdig.
      Total tid: 00:12:34
      Success: 180
      Failed: 20

---

## Noter
   - Download sker asynkront og parallelt
   - SemaphoreSlim begrænser antal samtidige downloads
   - Hvis PrimaryUrl fejler, forsøges automatisk FallbackUrl
   - Excel-kolonner valideres før læsning
   - Downloadtid måles pr. rapport

---

## Fremtidige forbedringer
   - ExcelOutputPath skal opdateres efter hver enkelt download i stedet for til sidst efter alle
   - Retry-strategi med eksponentiel backoff
   - Mulighed for at genoptage afbrudte downloads
   - Logging via ILogger i stedet for Console.WriteLine
