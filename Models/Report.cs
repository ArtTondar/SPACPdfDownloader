using System;
using System.Collections.Generic;
using System.Text;

namespace Models
{
    public class Report
    {
        public string BRNumber { get; set; }
        public string? PrimaryUrl { get; set; }
        public string? FallbackUrl { get; set; }
        public int Year { get; set; }
        public StatusMessage Status { get; set; }
        public string? LocalPath { get; set; }

        //Logging og statistik
        public double FileSizeKB { get; set; }
        public double DownloadTimeSeconds { get; set; }

    }
    public enum StatusMessage
    {
        NotDownloaded,
        Downloaded,
        Failed
    }
}
