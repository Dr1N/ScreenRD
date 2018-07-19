using ScreenCapture;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FomsClient
{
    public partial class ScreenForm : Form
    {
        public ScreenForm()
        {
            InitializeComponent();
        }

        private void btnScreen_Click(object sender, EventArgs e)
        {
            ScreenShot screenShot = ScreenGrabber.Win32ScreenShot(2.0);
            Debug.WriteLine($"Size: {screenShot.Width} x {screenShot.Height} Length: {screenShot.Bytes.Length}");
            pbScreen.Image = CreateBitmapFromScreenShot(screenShot);
        }

        private Bitmap CreateBitmapFromScreenShot(ScreenShot screenShot)
        {
            BitmapData bmpdata = null;
            Bitmap bmp = new Bitmap(screenShot.Width, screenShot.Height);
            try
            {
                bmpdata = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height), ImageLockMode.ReadOnly, bmp.PixelFormat);
                int numbytes = bmpdata.Stride * bmp.Height;
                IntPtr ptr = bmpdata.Scan0;
                Marshal.Copy(screenShot.Bytes, 0, ptr, numbytes);

                return bmp;
            }
            finally
            {
                if (bmpdata != null)
                {
                    bmp.UnlockBits(bmpdata);
                }
            }
        }
    }
}
