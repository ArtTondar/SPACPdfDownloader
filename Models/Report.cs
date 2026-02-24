using System;
using System.Collections.Generic;
using System.Text;

namespace Models
{
    public class Report
    {
        public int BRNumber { get; set; }
        public string? PrimaryUrl { get; set; }
        public string? FallbackUrl { get; set; }
        public int Year { get; set; }
        public StatusMessage Status { get; set; } = StatusMessage.NotDownloaded;
        public string LocalPath { get; set; }
        public bool IsLocalFile { get; set; }

        //LastAttempt og RetryCount - resume/retry
        public DateTime? LastAttempt { get; set; }
        public int RetryCount { get; set; }

        //Logging og statistik
        public double? FileSizeKB { get; set; }
        public string Domain { get; set; }
        public double? DownloadTimeSeconds { get; set; }

    }
    public enum StatusMessage
    {
        NotDownloaded,
        Downloaded,
        Failed
    }
}
