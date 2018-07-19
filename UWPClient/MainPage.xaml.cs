using Microsoft.Graphics.Canvas;
using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;
using Windows.ApplicationModel.AppService;
using Windows.Foundation.Collections;
using Windows.Graphics.Display;
using Windows.Storage;
using Windows.Storage.Pickers;
using Windows.UI;
using Windows.UI.Popups;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace SDKTemplate
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        static CanvasDevice device;
        static CanvasSwapChain swapChain;
        static float dpi;
        static CancellationToken token;
        static CancellationTokenSource source;
        static Stopwatch stopwatch = new Stopwatch();
        static bool isCompressData = false;

        public MainPage()
        {
            this.InitializeComponent();
            Loaded += MainPage_Loaded;
        }

        private async Task InitShapChain(int width, int height)
        {
            await Dispatcher.TryRunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, () =>
            {
                swapChain = new CanvasSwapChain(device, width, height, dpi);
                SwapChainPanel.SwapChain = swapChain;
            });
        }

        private void MainPage_Loaded(object sender, RoutedEventArgs e)
        {
            dpi = DisplayInformation.GetForCurrentView().LogicalDpi;
            device = CanvasDevice.GetSharedDevice();
        }

        private async void StartServiceBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                await Windows.ApplicationModel.FullTrustProcessLauncher.LaunchFullTrustProcessForCurrentAppAsync();
                Debug.WriteLine("[[[ Service Launched... ]]]");
                StartBtn.IsEnabled = true;
                StartServiceBtn.IsEnabled = false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                MessageDialog dialog = new MessageDialog("Rebuild the solution and make sure the BackgroundProcess is in your AppX folder");
                await dialog.ShowAsync();
            }
        }

        private void StartBtn_Click(object sender, RoutedEventArgs e)
        {
            source = new CancellationTokenSource();
            token = source.Token;
            StopBtn.IsEnabled = true;
            StartBtn.IsEnabled = false;
            Task.Run(async () =>
            {
                Debug.WriteLine("Start Screen Cast");
                while (token.IsCancellationRequested == false)
                {
                    ValueSet valueSet = new ValueSet
                    {
                        { "screen", isCompressData }
                    };

                    if (App.Connection != null)
                    {
                        stopwatch.Start();
                        AppServiceResponse response = await App.Connection.SendMessageAsync(valueSet);
                        var responseTime = stopwatch.Elapsed.Milliseconds;

                        var screenBytes = response.Message["screen"] as byte[];
                        var width = (int)response.Message["screenWidth"];
                        var height = (int)response.Message["screenHeight"];

                        double len = Math.Round((double)screenBytes.Length / 1024, 2);

                        if (swapChain == null)
                        {
                            await InitShapChain(width, height);
                        }

                        stopwatch.Restart();

                        var decompressTime = 0;
                        if (isCompressData == true)
                        {
                            screenBytes = Decompress(screenBytes);
                            decompressTime = stopwatch.Elapsed.Milliseconds;
                            stopwatch.Restart();
                        }

                        using (CanvasDrawingSession ds = SwapChainPanel.SwapChain.CreateDrawingSession(Colors.Transparent))
                        {
                            using (CanvasBitmap screen = CanvasBitmap.CreateFromBytes(device, screenBytes, width, height, Windows.Graphics.DirectX.DirectXPixelFormat.B8G8R8A8UIntNormalized))
                            {
                                ds.DrawImage(screen, new Windows.Foundation.Rect(0, 0, 800, 600));
                            }
                        }
                        SwapChainPanel.SwapChain.Present();
                        var drawTime = stopwatch.Elapsed.Milliseconds;
                        stopwatch.Stop();
                        
                        Debug.WriteLine($"Response: {responseTime} ms Draw:{drawTime} ms Decompress: {decompressTime} ms Length: {len} Kb");
                    }
                    await Task.Delay(TimeSpan.FromMilliseconds(25));
                }
                Debug.WriteLine("Stop Screen Cast");
            }, token);
        }

        private void StopBtn_Click(object sender, RoutedEventArgs e)
        {
            if (token.CanBeCanceled)
            {
                source.Cancel();
                StartBtn.IsEnabled = true;
                StopBtn.IsEnabled = false;
            }
        }

        private async void OfficeBtn_Click(object sender, RoutedEventArgs e)
        {
            string srcFullFileName = await GetFilePathAsync();
            string srcFileName = Path.GetFileName(srcFullFileName);
            string pdfFileName = Path.ChangeExtension(srcFileName, "pdf");
            StorageFolder localFolder = ApplicationData.Current.LocalFolder;
            string dstFullFileName = localFolder.Path + Path.DirectorySeparatorChar + pdfFileName;

            Debug.WriteLine($"SRC: {srcFullFileName}\nDST: {dstFullFileName}");

            string[] paths = new string[] 
            {
                srcFullFileName,
                dstFullFileName
            };
            ValueSet valueSet = new ValueSet
            {
                { "convert", paths }
            };

            if (App.Connection != null)
            {
                Debug.WriteLine("[Convert Response]");
                AppServiceResponse response = await App.Connection.SendMessageAsync(valueSet);
                var success = response.Message["success"] as bool?;
                if (success == true)
                {
                    StorageFile pdfFile = await localFolder.GetFileAsync(pdfFileName);
                    Debug.WriteLine($"[Success]: {pdfFile.Path}");
                }
            }
        }

        private static byte[] Decompress(byte[] gzip)
        {
            using (GZipStream stream = new GZipStream(new MemoryStream(gzip), CompressionMode.Decompress))
            {
                const int size = 4096;
                byte[] buffer = new byte[size];
                using (MemoryStream memory = new MemoryStream())
                {
                    int count = 0;
                    do
                    {
                        count = stream.Read(buffer, 0, size);
                        if (count > 0)
                        {
                            memory.Write(buffer, 0, count);
                        }
                    }
                    while (count > 0);
                    return memory.ToArray();
                }
            }
        }

        //TODO
        private static void DrawScreen()
        {

        }

        private async Task<string> GetFilePathAsync()
        {
            FileOpenPicker openPicker = new FileOpenPicker
            {
                ViewMode = PickerViewMode.Thumbnail,
                SuggestedStartLocation = PickerLocationId.DocumentsLibrary
            };
            openPicker.FileTypeFilter.Add(".docx");
            openPicker.FileTypeFilter.Add(".xlsx");
            openPicker.FileTypeFilter.Add(".ppt");

            StorageFile file = await openPicker.PickSingleFileAsync();
            if (file != null)
            {
                return file.Path;
            }

            return null;
        }
    }
}
