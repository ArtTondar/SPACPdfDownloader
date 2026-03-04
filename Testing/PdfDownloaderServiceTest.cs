using BusinessLogicLayer;
using Models;
using Moq;
using Moq.Protected;
using System.Net;

namespace Testing
{
    public class PdfDownloaderServiceTest
    {
        private HttpClient CreateMockHttpClient(Func<HttpRequestMessage, HttpResponseMessage> sendAsyncFunc)
        {
            var handlerMock = new Mock<HttpMessageHandler>(MockBehavior.Strict);

            handlerMock
               .Protected()
               .Setup<Task<HttpResponseMessage>>(
                   "SendAsync",
                   ItExpr.IsAny<HttpRequestMessage>(),
                   ItExpr.IsAny<CancellationToken>()
               )
               .ReturnsAsync((HttpRequestMessage request, CancellationToken token) =>
               {
                   return sendAsyncFunc(request);
               });

            return new HttpClient(handlerMock.Object);
        }

        [Fact]
        public async Task DownloadReportsAsync_SuccessfulDownload_SetsReportProperties()
        {
            // Arrange
            var report = new Report
            {
                BRNumber = "123",
                Year = 2025,
                PrimaryUrl = "http://example.com/file.pdf"
            };

            var httpClient = CreateMockHttpClient(req =>
            {
                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(new byte[] { 1, 2, 3 }) // lille fake PDF
                    {
                        Headers = { ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf") }
                    }
                };
            });

            var service = new PdfDownloaderService(httpClient);

            string tempFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            int maxParallel = 8;

            Directory.CreateDirectory(tempFolder);

            // Act
            await service.DownloadReportsAsync(new List<Report> { report }, tempFolder, maxParallel);

            // Assert
            Assert.Equal(StatusMessage.Downloaded, report.Status);
            Assert.NotNull(report.LocalPath);
            Assert.True(File.Exists(report.LocalPath));
            Assert.True(report.FileSizeKB > 0);
        }

        [Fact]
        public async Task DownloadReportsAsync_FallbackUrlUsed_WhenPrimaryFails()
        {
            // Arrange
            var report = new Report
            {
                BRNumber = "456",
                Year = 2025,
                PrimaryUrl = "http://example.com/fail.pdf",
                FallbackUrl = "http://example.com/success.pdf"
            };

            var httpClient = CreateMockHttpClient(req =>
            {
                if (req.RequestUri!.ToString().Contains("fail"))
                {
                    return new HttpResponseMessage(HttpStatusCode.InternalServerError);
                }

                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new ByteArrayContent(new byte[] { 1, 2, 3 })
                    {
                        Headers = { ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/pdf") }
                    }
                };
            });

            var service = new PdfDownloaderService(httpClient);
            string tempFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            int maxParallel = 8;
            Directory.CreateDirectory(tempFolder);

            // Act
            await service.DownloadReportsAsync(new List<Report> { report }, tempFolder, maxParallel);

            // Assert
            Assert.Equal(StatusMessage.Downloaded, report.Status);
            Assert.NotNull(report.LocalPath);
            Assert.True(File.Exists(report.LocalPath));
        }

        [Fact]
        public async Task DownloadReportsAsync_AllUrlsFail_SetsFailedStatus()
        {
            // Arrange
            var report = new Report
            {
                BRNumber = "789",
                Year = 2025,
                PrimaryUrl = "http://example.com/fail.pdf",
                FallbackUrl = "http://example.com/fail2.pdf"
            };

            var httpClient = CreateMockHttpClient(req =>
            {
                return new HttpResponseMessage(HttpStatusCode.InternalServerError);
            });

            var service = new PdfDownloaderService(httpClient);
            string tempFolder = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            int maxParallel = 8;
            Directory.CreateDirectory(tempFolder);

            // Act
            await service.DownloadReportsAsync(new List<Report> { report }, tempFolder, maxParallel);

            // Assert
            Assert.Equal(StatusMessage.Failed, report.Status);
            Assert.Null(report.LocalPath);
        }
    }
}