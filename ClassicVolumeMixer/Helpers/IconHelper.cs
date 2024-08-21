using System;
using System.Drawing;
using System.Runtime.InteropServices;
using Microsoft.Win32;

namespace ClassicVolumeMixer.Helpers
{
    public static class IconHelper
    {
        [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

        public static Icon ExtractIcon(string sFile, int iIndex, bool flipColors)
        {
            ExtractIconEx(sFile, iIndex, out IntPtr intPtr, out _, 1);
            Icon icon = Icon.FromHandle(intPtr);

            if (flipColors)
            {
                icon = FlipIconColors(icon);
            }
            return icon;
        }

        private static Icon FlipIconColors(Icon icon)
        {
            Bitmap bitmap = icon.ToBitmap();
            for (int y = 0; y < bitmap.Height; y++)
            {
                for (int x = 0; x < bitmap.Width; x++)
                {
                    Color pixelColor = bitmap.GetPixel(x, y);
                    Color flippedColor = Color.FromArgb(pixelColor.A, 255 - pixelColor.R, 255 - pixelColor.G, 255 - pixelColor.B);
                    bitmap.SetPixel(x, y, flippedColor);
                }
            }
            return Icon.FromHandle(bitmap.GetHicon());
        }
    }
}