using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using OfficeOpenXml;
using System.Diagnostics;

namespace GoogleAPI
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var driveService = CreateDriveService();

            var timer = new Stopwatch();
            
            Task.Run(async () =>
            {
                timer.Start();

                do
                {
                    DownloadAndWrite(driveService, timer.Elapsed);
                    await Task.Delay(5000);
                } while (timer.Elapsed.TotalMilliseconds < 5 * 60 * 1000);
            });

            timer.Stop();

            Console.ReadKey();
        }

        private static void DownloadAndWrite(DriveService driveService, TimeSpan time)
        {
            var ret = DriveDownloadFile(driveService, Constantes._FILE_ID);

            using var package = new ExcelPackage(ret.Result);
            Console.WriteLine($"[{time:m\\:ss}] CellValue: {package.Workbook.Worksheets[0].Cells[2, 1].Value}");
        }

        #region GoogleDrive

        private static DriveService CreateDriveService()
        {
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
            {
                ClientId = Constantes._CLIENT_ID,
                ClientSecret = Constantes._CLIENT_SECRET
            },
                new[] { DriveService.Scope.Drive, DriveService.Scope.DriveFile },
                "user",
                CancellationToken.None,
                new FileDataStore("Drive.Auth.Store")).Result;

            return new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "Drive Service account Authentication Sample",
            });
        }

        private static async Task<MemoryStream?> DriveDownloadFile(DriveService driveService, string fileId)
        {
            try
            {
                var request = driveService.Files.Get(fileId);
                var stream = new MemoryStream();

                //request.MediaDownloader.ProgressChanged +=
                //    progress =>
                //    {
                //        switch (progress.Status)
                //        {
                //            case DownloadStatus.Downloading:
                //                Console.WriteLine(progress.BytesDownloaded);
                //                break;
                //            case DownloadStatus.Completed:
                //                Console.WriteLine("Download complete.");
                //                break;
                //            case DownloadStatus.Failed:
                //                Console.WriteLine("Download failed.");
                //                break;
                //        }
                //    };

                await request.DownloadAsync(stream);

                return stream;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

            return null;
        }

        #endregion
    }
}