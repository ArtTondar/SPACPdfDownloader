using Models;
using System.Diagnostics;

/// <summary>
/// Service til at downloade PDF-rapporter parallelt fra angivne URLs.
/// </summary>
public class PdfDownloaderService
{
    private readonly HttpClient _httpClient;
    private readonly object _heartbeatLock = new object();

    /// <summary>
    /// Initialiserer tjenesten med en HttpClient.
    /// HttpClient bør genbruges for performance og for at undgå socket exhaustion.
    /// </summary>
    /// <param name="httpClient">HttpClient til download af filer.</param>
    public PdfDownloaderService(HttpClient httpClient)
    {
        _httpClient = httpClient;
    }

    /// <summary>
    /// Downloader en liste af rapporter parallelt og gemmer dem i den angivne output-mappe.
    /// 
    /// Metoden:
    /// - Begrænser antal samtidige downloads via SemaphoreSlim
    /// - Logger løbende fremdrift via PeriodicTimer (heartbeat)
    /// - Måler samlet køretid
    /// - Logger en samlet slutstatus når alle downloads er færdige
    /// </summary>
    /// <param name="reports">
    /// Listen af rapporter der skal downloades.
    /// Hver rapport opdateres (ønsket løbende) med status, filsti, filstørrelse og downloadtid.
    /// </param>
    /// <param name="outputFolder">
    /// Sti til output-mappen, hvor PDF-filer gemmes i undermapper pr. år.
    /// </param>
    /// <param name="maxParallel">
    /// Maksimalt antal samtidige downloads.
    /// Bruges til at undgå overbelastning af netværk eller server.
    /// </param>
    public async Task DownloadReportsAsync(List<Report> reports, string outputFolder, int maxParallel)
    {
        // Sørg for at output-mappen eksisterer før download starter
        Directory.CreateDirectory(outputFolder);

        // Semaphore bruges til at begrænse antal samtidige downloads
        SemaphoreSlim semaphore = new SemaphoreSlim(maxParallel);

        // Tællere opdateres trådsikkert via Interlocked
        int processed = 0; // Antal færdigbehandlede rapporter (success + failed)
        int success = 0;   // Antal succesfulde downloads
        int failed = 0;    // Antal fejlede downloads

        // Stopwatch måler samlet køretid for hele batchen
        Stopwatch totalStopwatch = Stopwatch.StartNew();

        // CancellationToken bruges til at stoppe heartbeat-timeren,
        // når alle downloads er færdige
        using CancellationTokenSource cts = new CancellationTokenSource();

        // Starter en baggrundsopgave, der logger fremdrift periodisk
        Task heartbeatTask = Task.Run(async () =>
        {
            // PeriodicTimer udløses hvert 3. sekund
            PeriodicTimer timer = new PeriodicTimer(TimeSpan.FromSeconds(3));

            try
            {
                while (await timer.WaitForNextTickAsync(cts.Token))
                {
                    // Undgå division med nul før første rapport er behandlet
                    if (processed == 0)
                        continue;

                    // Beregn gennemsnitlig behandlingstid pr. rapport
                    double avg = totalStopwatch.Elapsed.TotalSeconds / processed;

                    // Estimer resterende tid (ETA)
                    double remaining = (reports.Count - processed) * avg;

                    // Log aktuel fremdrift
                    Console.WriteLine(
                        $"[{DateTime.Now:HH:mm:ss}] " +
                        $"Processed: {processed}/{reports.Count} | " +
                        $"Success: {success} | Failed: {failed} | " +
                        $"Elapsed: {totalStopwatch.Elapsed:hh\\:mm\\:ss} | " +
                        $"ETA: {TimeSpan.FromSeconds(remaining):hh\\:mm\\:ss}");
                }
            }
            catch (OperationCanceledException)
            {
                // Forventet når vi stopper timeren efter afsluttet batch
            }
        });

        // Opretter en asynkron opgave pr. rapport
        IEnumerable<Task> downloadTasks = reports.Select(async report =>
        {
            // Vent hvis maks. antal samtidige downloads er nået
            await semaphore.WaitAsync();

            try
            {
                // Forsøg at downloade og gem rapporten
                bool downloaded = await ProcessSingleReportAsync(report, outputFolder);

                // Opdater succes/fejl-tællere trådsikkert
                if (downloaded)
                    Interlocked.Increment(ref success);
                else
                    Interlocked.Increment(ref failed);
            }
            finally
            {
                // Marker rapport som behandlet
                Interlocked.Increment(ref processed);

                // Frigiv plads i semaphore, så en ny download kan starte
                semaphore.Release();
            }
        });

        // Vent på at alle downloads færdiggøres
        await Task.WhenAll(downloadTasks);

        // Stop heartbeat-logging
        cts.Cancel();

        // Vent på at heartbeat-task afsluttes korrekt
        await heartbeatTask;

        // Log samlet slutstatus (total runtime + success/failed)
        LogFinalSummary(totalStopwatch, success, failed);
    }

    /// <summary>
    /// Downloader og gemmer en enkelt rapport.
    /// </summary>
    /// <param name="report">Rapporten der skal behandles.</param>
    /// <param name="outputFolder">Output-mappen.</param>
    /// <returns>True hvis download lykkedes, ellers false.</returns>
    private async Task<bool> ProcessSingleReportAsync(Report report, string outputFolder)
    {
        Stopwatch sw = Stopwatch.StartNew(); // Start tidtagning

        string?[] urls = new[] { report.PrimaryUrl, report.FallbackUrl };

        byte[]? fileBytes = await TryDownloadFromUrlsAsync(urls);

        sw.Stop(); // Stop tidtagning

        report.DownloadTimeSeconds = sw.Elapsed.TotalSeconds; // Gem downloadtid

        if (fileBytes == null)
        {
            report.Status = StatusMessage.Failed;
            return false;
        }

        string fullPath = await SaveFileAsync(report, fileBytes, outputFolder);

        report.LocalPath = fullPath;
        report.FileSizeKB = fileBytes.Length / 1024.0;
        report.Status = StatusMessage.Downloaded;

        return true;
    }

    /// <summary>
    /// Prøver at downloade filen fra flere URLs med fallback.
    /// </summary>
    /// <param name="urls">Liste af URLs (primær + fallback).</param>
    /// <returns>Byte-array med PDF-indhold, eller null hvis ingen URL virkede.</returns>
    private async Task<byte[]?> TryDownloadFromUrlsAsync(IEnumerable<string> urls)
    {
        foreach (string? url in urls.Where(u => !string.IsNullOrWhiteSpace(u)))
        {
            try
            {
                return await DownloadPdfAsync(url); // Forsøg download
            }
            catch
            {
                // Hvis download fejler, prøv næste URL
            }
        }

        return null; // Ingen URL lykkedes
    }

    /// <summary>
    /// Downloader en PDF fra en given URL.
    /// </summary>
    /// <param name="url">URL til PDF'en.</param>
    /// <returns>Byte-array med PDF-indhold.</returns>
    /// <exception cref="Exception">Hvis filen ikke er en PDF eller HTTP-status fejler.</exception>
    private async Task<byte[]> DownloadPdfAsync(string url)
    {
        using HttpResponseMessage response = await _httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode(); // Kaster hvis HTTP-fejl

        string? contentType = response.Content.Headers.ContentType?.MediaType;

        if (contentType == null || !contentType.Contains("pdf"))
            throw new Exception("Not a PDF");

        return await response.Content.ReadAsByteArrayAsync();
    }

    /// <summary>
    /// Gemmer en PDF på disk i en år-baseret undermappe.
    /// </summary>
    /// <param name="report">Rapporten der skal gemmes.</param>
    /// <param name="fileBytes">Byte-array med PDF.</param>
    /// <param name="outputFolder">Base output-mappen.</param>
    /// <returns>Den fulde sti hvor filen blev gemt.</returns>
    private async Task<string> SaveFileAsync(Report report, byte[] fileBytes, string outputFolder)
    {
        string yearFolder = Path.Combine(outputFolder, report.Year.ToString());
        Directory.CreateDirectory(yearFolder); // Sørg for at år-mappen eksisterer

        string fullPath = Path.Combine(yearFolder, $"{report.BRNumber}.pdf");

        await File.WriteAllBytesAsync(fullPath, fileBytes); // Async skriv for performance

        return fullPath;
    }

    /// <summary>
    /// Logger slutstatus for alle downloads.
    /// </summary>
    private void LogFinalSummary(Stopwatch stopwatch, int success, int failed)
    {
        Console.WriteLine("\nDownload færdig.");
        Console.WriteLine($"Total tid: {stopwatch.Elapsed:hh\\:mm\\:ss}");
        Console.WriteLine($"Success: {success}");
        Console.WriteLine($"Failed: {failed}");
    }
}