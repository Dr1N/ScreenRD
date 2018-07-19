using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ScreenCapture
{
    public class ScreenGrabber
    {
        #region Fields

        private static byte[] buffer;
        
        #endregion

        #region Interop

        /// <summary>
        /// Helper class containing Gdi32 API functions
        /// </summary>
        private class GDI32
        {
            public const int SRCCOPY = 0x00CC0020; // BitBlt dwRop parameter

            [DllImport("gdi32.dll")]
            public static extern bool BitBlt(IntPtr hObject, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hObjectSource, int nXSrc, int nYSrc, int dwRop);
            [DllImport("gdi32.dll")]
            public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC, int nWidth, int nHeight);
            [DllImport("gdi32.dll")]
            public static extern IntPtr CreateCompatibleDC(IntPtr hDC);
            [DllImport("gdi32.dll")]
            public static extern bool DeleteDC(IntPtr hDC);
            [DllImport("gdi32.dll")]
            public static extern bool DeleteObject(IntPtr hObject);
            [DllImport("gdi32.dll")]
            public static extern IntPtr SelectObject(IntPtr hDC, IntPtr hObject);
        }

        /// <summary>
        /// Helper class containing User32 API functions
        /// </summary>
        private class User32
        {
            [StructLayout(LayoutKind.Sequential)]
            public struct RECT
            {
                public int left;
                public int top;
                public int right;
                public int bottom;
            }
            [DllImport("user32.dll")]
            public static extern IntPtr GetDesktopWindow();
            [DllImport("user32.dll")]
            public static extern IntPtr GetWindowDC(IntPtr hWnd);
            [DllImport("user32.dll")]
            public static extern IntPtr ReleaseDC(IntPtr hWnd, IntPtr hDC);
            [DllImport("user32.dll")]
            public static extern IntPtr GetWindowRect(IntPtr hWnd, ref RECT rect);
        }

        #endregion

        #region Public
        
        public static ScreenShot Win32ScreenShot(double dpi = 1.0)
        {
            using (Bitmap bmp = CaptureScreen(dpi))
            {
                MarshalCopy(bmp);
                return new ScreenShot(bmp.Width, bmp.Height, buffer);
            }
        }
        
        #endregion

        #region Private

        private static void MarshalCopy(Bitmap bmp)
        {
            BitmapData bmpdata = null;
            try
            {
                bmpdata = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.ReadOnly, bmp.PixelFormat);
                int numbytes = bmpdata.Stride * bmp.Height;
                if (buffer == null || buffer.Length != numbytes)
                {
                    System.Diagnostics.Debug.WriteLine("[[[ Create Bufer ]]]");
                    buffer = new byte[numbytes];
                }
                IntPtr ptr = bmpdata.Scan0;

                Marshal.Copy(ptr, buffer, 0, numbytes);
            }
            finally
            {
                if (bmpdata != null)
                {
                    bmp.UnlockBits(bmpdata);
                }
            }
        }

        /// <summary>
        /// Creates an Image object containing a screen shot of the entire desktop
        /// </summary>
        /// <returns></returns>
        private static Bitmap CaptureScreen(double dpi = 1.0)
        {
            return CaptureWindow(User32.GetDesktopWindow(), dpi);
        }
        
        /// <summary>
        /// Captures a screen shot of a specific window, and saves it to a file
        /// </summary>
        /// <param name="handle"></param>
        /// <param name="filename"></param>
        /// <param name="format"></param>
        private static void CaptureWindowToFile(IntPtr handle, string filename, ImageFormat format, double dpi = 1.0)
        {
            Image img = CaptureWindow(handle, dpi);
            img.Save(filename, format);
        }

        /// <summary>
        /// Captures a screen shot of the entire desktop, and saves it to a file
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="format"></param>
        private static void CaptureScreenToFile(string filename, ImageFormat format, double dpi = 1.0)
        {
            Image img = CaptureScreen(dpi);
            img.Save(filename, format);
        }

        /// <summary>
        /// Creates an Image object containing a screen shot of a specific window
        /// </summary>
        /// <param name="handle">The handle to the window. (In windows forms, this is obtained by the Handle property)</param>
        /// <returns></returns>
        private static Bitmap CaptureWindow(IntPtr handle, double dpi)
        {
            // get te hDC of the target window
            IntPtr hdcSrc = User32.GetWindowDC(handle);
            // get the size
            User32.RECT windowRect = new User32.RECT();
            User32.GetWindowRect(handle, ref windowRect);
            int width = (int)((windowRect.right - windowRect.left) * dpi);
            int height = (int)((windowRect.bottom - windowRect.top) * dpi);
            // create a device context we can copy to
            IntPtr hdcDest = GDI32.CreateCompatibleDC(hdcSrc);
            // create a bitmap we can copy it to,
            // using GetDeviceCaps to get the width/height
            IntPtr hBitmap = GDI32.CreateCompatibleBitmap(hdcSrc, width, height);
            // select the bitmap object
            IntPtr hOld = GDI32.SelectObject(hdcDest, hBitmap);
            // bitblt over
            bool success = GDI32.BitBlt(hdcDest, 0, 0, width, height, hdcSrc, 0, 0, GDI32.SRCCOPY);

            try
            {
                if (success)
                {
                    Bitmap bitmap = (Bitmap)Image.FromHbitmap(hBitmap);
                    bitmap.MakeTransparent(Color.Transparent);

                    return bitmap;
                }
                else
                {
                    return null;
                }
            }
            finally
            {
                // restore selection
                GDI32.SelectObject(hdcDest, hOld);
                // clean up 
                GDI32.DeleteDC(hdcDest);
                User32.ReleaseDC(handle, hdcSrc);
                // free up the Bitmap object
                GDI32.DeleteObject(hBitmap);
            }
        }

        #endregion
    }
}
