using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;

namespace ScreenCapture
{
    public class ScreenShot
    {
        public int Width { get; private set; }
        public int Height { get; private set; }
        public byte[] Bytes { get; private set; }
        
        public ScreenShot(int width, int height, byte[] bytes)
        {
            Width = width;
            Height = height;
            Bytes = bytes;
        }
    }
}
