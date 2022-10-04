using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using OfficeOpenXml;
using System.Diagnostics;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using Google.Apis.Sheets.v4.Data;

namespace GoogleAPI
{
    internal static class Program
    {
        private static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            //var driveService = CreateDriveService();
            var sheetService = CreateSheetService();

            var timer = new Stopwatch();
            
            Task.Run(async () =>
            {
                timer.Start();

                do
                {
                    //DownloadAndWrite(driveService, timer.Elapsed);
                    await SheetReadFileAndWrite(sheetService, timer.Elapsed);
                    
                    await Task.Delay(5000);
                } while (timer.Elapsed.TotalMilliseconds < 5 * 60 * 1000);
            });

            timer.Stop();

            Console.ReadKey();
        }

        private static async Task DownloadAndWrite(DriveService driveService, TimeSpan time)
        {
            var ret = await DriveDownloadFile(driveService, Constantes._FILE_ID_SHEET);

            using var package = new ExcelPackage(ret);
            Console.WriteLine($"[{time:m\\:ss}] CellValue: {package.Workbook.Worksheets[0].Cells[2, 1].Value}");
        }

        #region Google Drive Excel

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
                //var request = driveService.Files.Export(fileId, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); //Not working
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

        #region Google Drive Sheet

        private static SheetsService CreateSheetService()
        {
            var credential = GoogleWebAuthorizationBroker.AuthorizeAsync(new ClientSecrets
                {
                    ClientId = Constantes._CLIENT_ID,
                    ClientSecret = Constantes._CLIENT_SECRET
                }, 
                new[] { SheetsService.Scope.SpreadsheetsReadonly },
                "user",
                CancellationToken.None, 
                new FileDataStore("Drive.Auth.Store")).Result;

            return new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = credential,
                ApplicationName = "Spreadsheet Service account Authentication Sample",
            });
        }

        private static async Task<Spreadsheet> SheetReadFile(SheetsService sheetService, string fileId)
        {
            var range = "Sheet1!A2";

            var request = sheetService.Spreadsheets.Get(fileId);
            request.Ranges = range;
            request.IncludeGridData = true;

            return await request.ExecuteAsync();

        }

        private static async Task SheetReadFileAndWrite(SheetsService sheetService, TimeSpan time)
        {
            var ret = await SheetReadFile(sheetService, Constantes._FILE_ID_SHEET);

            Console.WriteLine($"[{time:m\\:ss}] CellValue: {JsonConvert.SerializeObject(ret.Sheets[0].Data[0].RowData[0].Values[0].EffectiveValue.NumberValue )}");
        }

        #endregion
    }
}