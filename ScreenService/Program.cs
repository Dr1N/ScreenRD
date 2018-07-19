using System;
using System.Linq;
using System.Threading;
using Windows.Foundation.Collections;
using Windows.ApplicationModel.AppService;
using ScreenCapture;
using System.IO;
using System.IO.Compression;
using System.Diagnostics;

namespace ScreenService
{
    class Program
    {
        static AppServiceConnection connection = null;
        static byte[] buffer;
        static int width;
        static int height;
        static object locker = new object();
        static int sleepTime = 0;
        static int frameCounter;
        static Stopwatch stopwatch = new Stopwatch();

        /// <summary>
        /// Creates an app service thread
        /// </summary>
        static void Main(string[] args)
        {
            Thread appServiceThread = new Thread(AppServiceProc) {
                IsBackground = true,
            };
            appServiceThread.Start();

            Thread screenThread = new Thread(ScreenProc)
            {
                IsBackground = true,
            };
            //screenThread.Start();

            Console.ReadLine();
        }

        /// <summary>
        /// Creates the app service connection
        /// </summary>
        static async void AppServiceProc()
        {
            Console.WriteLine("App Service Thread running...");
            Console.WriteLine();

            connection = new AppServiceConnection
            {
                AppServiceName = "CommunicationService",
                PackageFamilyName = Windows.ApplicationModel.Package.Current.Id.FamilyName
            };
            connection.RequestReceived += Connection_RequestReceived;

            AppServiceConnectionStatus status = await connection.OpenAsync();
            switch (status)
            {
                case AppServiceConnectionStatus.Success:
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Connection established - waiting for requests");
                    Console.ResetColor();
                    Console.WriteLine();
                    break;
                case AppServiceConnectionStatus.AppNotInstalled:
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("The app AppServicesProvider is not installed.");
                    return;
                case AppServiceConnectionStatus.AppUnavailable:
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("The app AppServicesProvider is not available.");
                    return;
                case AppServiceConnectionStatus.AppServiceUnavailable:
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(string.Format("The app AppServicesProvider is installed but it does not provide the app service {0}.", connection.AppServiceName));
                    return;
                case AppServiceConnectionStatus.Unknown:
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(string.Format("An unkown error occurred while we were trying to open an AppServiceConnection."));
                    return;
            }
        }

        /// <summary>
        /// Create Screen Shots
        /// </summary>
        static void ScreenProc()
        {
            Console.WriteLine("Screen Thread running...");
            while (true)
            {
                if (stopwatch.IsRunning == false) stopwatch.Start();
                ScreenShot screenShotData= ScreenGrabber.Win32ScreenShot(2.0);
                lock (locker)
                {
                    buffer = screenShotData.Bytes;
                }
                width = screenShotData.Width;
                height = screenShotData.Height;
                PrintFps();
                Thread.Sleep(sleepTime);
            }
        }

        /// <summary>
        /// Receives message from UWP app and sends a response back
        /// </summary>
        private static void Connection_RequestReceived(AppServiceConnection sender, AppServiceRequestReceivedEventArgs args)
        {
            string key = args.Request.Message.First().Key;
            if (key == "screen")
            {
                bool value = (bool)args.Request.Message.First().Value;
                var data = new byte[buffer.Length];
                lock (locker)
                {
                    Array.Copy(buffer, data, buffer.Length);
                }
                if (value == true)
                {
                    data = Compress(data);
                    PrintCompressData(buffer.Length, data.Length);
                }
                ValueSet valueSet = new ValueSet
                {
                    { "screen", data },
                    { "screenWidth", width },
                    { "screenHeight", height }
                };
                args.Request.SendResponseAsync(valueSet).Completed += delegate { };
            }
            else if (key == "convert")
            {
                string[] value = args.Request.Message.First().Value as string[];
                Console.WriteLine("Convert Response:");
                Console.WriteLine($"\tSRC: {value[0]}\n\tDST: {value[1]}");
                bool success = false;
                try
                {
                    success = ConvertDocument(value[0], value[1]);
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine(ex.Message);
                    Console.ResetColor();
                }

                Console.WriteLine(success ? "Success" : "Error");

                ValueSet valueSet = new ValueSet
                {
                    { "success", success },
                };
                args.Request.SendResponseAsync(valueSet).Completed += delegate { };
            }
        }

        /// <summary>
        /// Compresses byte array to new byte array.
        /// </summary>
        private static byte[] Compress(byte[] raw)
        {
            using (MemoryStream memory = new MemoryStream())
            {
                using (GZipStream gzip = new GZipStream(memory, CompressionMode.Compress, true))
                {
                    gzip.Write(raw, 0, raw.Length);
                }
                return memory.ToArray();
            }
        }

        private static void PrintCompressData(int uLength, int cLength)
        {
            var uncopressedLength = Math.Round((double)uLength / 1024, 2);
            var compressedLength = Math.Round((double)cLength / 1024, 2);
            var compressRate = Math.Round(100.0 * compressedLength / uncopressedLength, 2);
            Console.WriteLine($"UData: {uncopressedLength} Kb CData: {compressedLength} Kb Rate: {compressRate} %");
        }

        private static void PrintFps()
        {
            if (++frameCounter >= 10)
            {
                var time = stopwatch.Elapsed.Milliseconds;
                var fps = Math.Round(1000.0 * frameCounter / time, 1);
                stopwatch.Restart();
                frameCounter = 0;
                Console.WriteLine($"FPS: {fps}");
            }
        }

        private static bool ConvertDocument(string src, string dst)
        {
            string extension = Path.GetExtension(src);
            switch (extension)
            {
                case ".docx":
                    OfficeConverterLib.OfficeConverter.ConvertWordDocumentToPdf(src, dst);
                    break;
                case ".xlsx":
                    OfficeConverterLib.OfficeConverter.ConvertExcelDocumentToPdf(src, dst);
                    break;
                case ".ppt":
                    OfficeConverterLib.OfficeConverter.ConvertPowerPointDocumentToPdf(src, dst);
                    break;
                default:
                    return false;
            }

            return true;
        }
    }
}
