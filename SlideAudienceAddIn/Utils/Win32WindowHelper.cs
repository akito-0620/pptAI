using System;
using System.Runtime.InteropServices;

namespace SlideAudienceAddIn.Utils
{
    public static class Win32WindowHelper
    {
        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        public static bool TryGetWindowRect(IntPtr hWnd, out WindowRect rect)
        {
            rect = default(WindowRect);
            if (hWnd == IntPtr.Zero || !GetWindowRect(hWnd, out var nativeRect))
            {
                return false;
            }

            rect = new WindowRect(
                nativeRect.Left,
                nativeRect.Top,
                nativeRect.Right - nativeRect.Left,
                nativeRect.Bottom - nativeRect.Top);
            return true;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }

    public struct WindowRect
    {
        public WindowRect(int left, int top, int width, int height)
        {
            Left = left;
            Top = top;
            Width = width;
            Height = height;
        }

        public int Left { get; }

        public int Top { get; }

        public int Width { get; }

        public int Height { get; }
    }
}
