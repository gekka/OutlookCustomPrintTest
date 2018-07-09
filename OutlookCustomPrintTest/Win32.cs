namespace OutlookCustomPrintTest
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.Runtime.InteropServices;
    class Win32
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern int GetWindowTextLength(IntPtr hWnd);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        public static string GetWindowText(IntPtr hWnd)
        {
            StringBuilder sb = new StringBuilder(101);
            GetWindowText(hWnd, sb, 100);
            return sb.ToString();
        }

        [DllImport("user32", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        public static IList<IntPtr> EnumWindows()
        {
            List<IntPtr> list = new List<IntPtr>();
            EnumWindows((hWnd, lParam) => { list.Add(hWnd); return true; }, IntPtr.Zero);
            return list;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        public static RECT GetWindowRect(IntPtr hwnd)
        {
            RECT rect = default(RECT);
            if (!GetWindowRect(hwnd, ref rect))
            {
                throw new System.ComponentModel.Win32Exception();
            }
            return rect;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);
    }
}
