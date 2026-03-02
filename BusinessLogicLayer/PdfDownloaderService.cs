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
    /// Downloader en liste af rapporter parallelt og gemmer dem i outputFolder.
    /// </summary>
    /// <param name="reports">Listen af rapporter, der skal downloades.</param>
    /// <param name="outputFolder">Sti til output-mappen, hvor PDF'er gemmes i undermapper pr. år.</param>
    /// <param name="maxParallel">Maks antal samtidige downloads (dynamisk parameter, fx fra JSON-konfiguration).</param>
    public async Task DownloadReportsAsync(List<Report> reports, string outputFolder, int maxParallel)
    {
        Directory.CreateDirectory(outputFolder);

        var semaphore = new SemaphoreSlim(maxParallel);

        int processed = 0;
        int success = 0;
        int failed = 0;

        var totalStopwatch = Stopwatch.StartNew();

        // CancellationToken for heartbeat-timer
        using var cts = new CancellationTokenSource();

        // Start heartbeat logger
        var heartbeatTask = Task.Run(async () =>
        {
            var timer = new PeriodicTimer(TimeSpan.FromSeconds(3));
            try
            {
                while (await timer.WaitForNextTickAsync(cts.Token))
                {
                    if (processed == 0) continue;

                    double avg = totalStopwatch.Elapsed.TotalSeconds / processed;
                    double remaining = (reports.Count - processed) * avg;

                    Console.WriteLine(
                        $"[{DateTime.Now:HH:mm:ss}] " +
                        $"Processed: {processed}/{reports.Count} | " +
                        $"Success: {success} | Failed: {failed} | " +
                        $"Elapsed: {totalStopwatch.Elapsed:hh\\:mm\\:ss} | " +
                        $"ETA: {TimeSpan.FromSeconds(remaining):hh\\:mm\\:ss}");
                }
            }
            catch (OperationCanceledException) { }
        });

        // Download tasks
        var downloadTasks = reports.Select(async report =>
        {
            await semaphore.WaitAsync();
            try
            {
                bool downloaded = await ProcessSingleReportAsync(report, outputFolder);

                if (downloaded)
                    Interlocked.Increment(ref success);
                else
                    Interlocked.Increment(ref failed);
            }
            finally
            {
                Interlocked.Increment(ref processed);
                semaphore.Release();
            }
        });

        await Task.WhenAll(downloadTasks);

        // Stop heartbeat-loop
        cts.Cancel();
        await heartbeatTask;

        // Log slutstatus
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
        var sw = Stopwatch.StartNew(); // Start tidtagning

        var urls = new[] { report.PrimaryUrl, report.FallbackUrl };

        var fileBytes = await TryDownloadFromUrlsAsync(urls);

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
        foreach (var url in urls.Where(u => !string.IsNullOrWhiteSpace(u)))
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
        using var response = await _httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode(); // Kaster hvis HTTP-fejl

        var contentType = response.Content.Headers.ContentType?.MediaType;

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