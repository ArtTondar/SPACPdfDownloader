using DocumentFormat.OpenXml.Bibliography;
using Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace BusinessLogicLayer
{
    public class PdfDownloaderService
    {
        private readonly HttpClient _httpClient;

        public PdfDownloaderService()
        {
            _httpClient = new HttpClient();
        }

        public async Task DownloadReportsAsync(List<Report> reports, string outputFolder)
        {
            Directory.CreateDirectory(outputFolder);

            foreach (var report in reports)
            {
                string[] urls = { report.PrimaryUrl, report.FallbackUrl };
                bool downloaded = false;

                foreach (var url in urls.Where(u => !string.IsNullOrWhiteSpace(u)))
                {
                    try
                    {
                        var stopwatch = Stopwatch.StartNew();

                        using var response = await _httpClient.GetAsync(url);
                        response.EnsureSuccessStatusCode();

                        // Tjek at vi faktisk får en PDF
                        var contentType = response.Content.Headers.ContentType?.MediaType;
                        if (contentType == null || !contentType.Contains("pdf"))
                            throw new Exception("Response is not a PDF");

                        byte[] fileBytes = await response.Content.ReadAsByteArrayAsync();

                        // Opret year-mappe
                        string yearFolder = Path.Combine(outputFolder, report.Year.ToString());
                        Directory.CreateDirectory(yearFolder);

                        string fullPath = Path.Combine(yearFolder, $"{report.BRNumber}.pdf");

                        await File.WriteAllBytesAsync(fullPath, fileBytes);

                        stopwatch.Stop();

                        report.LocalPath = fullPath;
                        report.FileSizeKB = fileBytes.Length / 1024.0;
                        report.DownloadTimeSeconds = stopwatch.Elapsed.TotalSeconds;
                        report.Status = StatusMessage.Downloaded;

                        downloaded = true;
                        break; // stop fallback loop
                    }
                    catch (Exception)
                    {
                        // Prøv næste URL (fallback)
                    }
                }

                if (!downloaded)
                {
                    report.Status = StatusMessage.Failed;
                }
            }
        }
    }
}
