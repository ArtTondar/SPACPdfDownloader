using Models;
using System.Diagnostics;

public class PdfDownloaderService
{
    private readonly HttpClient _httpClient;

    public PdfDownloaderService(HttpClient httpClient)
    {
        _httpClient = httpClient;
    }

    public async Task DownloadReportsAsync(List<Report> reports, string outputFolder)
    {
        Directory.CreateDirectory(outputFolder);

        int maxParallel = 8;
        var semaphore = new SemaphoreSlim(maxParallel);

        int processed = 0;
        int success = 0;
        int failed = 0;

        var totalStopwatch = Stopwatch.StartNew();
        var heartbeatTimer = Stopwatch.StartNew();

        var tasks = reports.Select(async report =>
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

                LogHeartbeatIfNeeded(
                    heartbeatTimer,
                    totalStopwatch,
                    processed,
                    reports.Count,
                    success,
                    failed);
            }
        });

        await Task.WhenAll(tasks);

        totalStopwatch.Stop();
        LogFinalSummary(totalStopwatch, success, failed);
    }

    private async Task<bool> ProcessSingleReportAsync(Report report, string outputFolder)
    {
        var urls = new[] { report.PrimaryUrl, report.FallbackUrl };

        var fileBytes = await TryDownloadFromUrlsAsync(urls);

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

    private async Task<byte[]?> TryDownloadFromUrlsAsync(IEnumerable<string> urls)
    {
        foreach (var url in urls.Where(u => !string.IsNullOrWhiteSpace(u)))
        {
            try
            {
                return await DownloadPdfAsync(url);
            }
            catch
            {
                // prøv næste URL
            }
        }

        return null;
    }

    private async Task<byte[]> DownloadPdfAsync(string url)
    {
        using var response = await _httpClient.GetAsync(url);
        response.EnsureSuccessStatusCode();

        var contentType = response.Content.Headers.ContentType?.MediaType;

        if (contentType == null || !contentType.Contains("pdf"))
            throw new Exception("Not a PDF");

        return await response.Content.ReadAsByteArrayAsync();
    }

    private async Task<string> SaveFileAsync(Report report, byte[] fileBytes, string outputFolder)
    {
        string yearFolder = Path.Combine(outputFolder, report.Year.ToString());
        Directory.CreateDirectory(yearFolder);

        string fullPath = Path.Combine(yearFolder, $"{report.BRNumber}.pdf");

        await File.WriteAllBytesAsync(fullPath, fileBytes);

        return fullPath;
    }

    private void LogHeartbeatIfNeeded(
        Stopwatch heartbeatTimer,
        Stopwatch totalStopwatch,
        int processed,
        int total,
        int success,
        int failed)
    {
        if (heartbeatTimer.Elapsed.TotalSeconds < 15)
            return;

        double avg = totalStopwatch.Elapsed.TotalSeconds / processed;
        double remaining = (total - processed) * avg;

        Console.WriteLine(
            $"[{DateTime.Now:HH:mm:ss}] " +
            $"Processed: {processed}/{total} | " +
            $"Success: {success} | Failed: {failed} | " +
            $"Elapsed: {totalStopwatch.Elapsed:hh\\:mm\\:ss} | " +
            $"ETA: {TimeSpan.FromSeconds(remaining):hh\\:mm\\:ss}");

        heartbeatTimer.Restart();
    }

    private void LogFinalSummary(Stopwatch stopwatch, int success, int failed)
    {
        Console.WriteLine("\nDownload færdig.");
        Console.WriteLine($"Total tid: {stopwatch.Elapsed:hh\\:mm\\:ss}");
        Console.WriteLine($"Success: {success}");
        Console.WriteLine($"Failed: {failed}");
    }
}