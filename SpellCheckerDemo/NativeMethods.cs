using System;
using System.Runtime.InteropServices;
using System.Text;

public static class NativeMethods
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll" , CharSet = CharSet.Unicode)]
    public static extern IntPtr SendMessage(IntPtr hWnd , uint Msg , IntPtr wParam , StringBuilder lParam);

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll" , CharSet = CharSet.Unicode)]
    public static extern bool GetWindowText(IntPtr hWnd , StringBuilder text , int count);

    [DllImport("user32.dll" , CharSet = CharSet.Auto)]
    public static extern void keybd_event(byte bVk , byte bScan , uint dwFlags , UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern bool GetKeyboardState(byte[] pbKeyState);

    [DllImport("user32.dll")]
    public static extern uint MapVirtualKey(uint uCode , uint uMapType);

    [DllImport("user32.dll" , CharSet = CharSet.Unicode)]
    public static extern int ToUnicode(uint wVirtKey , uint wScanCode , byte[] pKeyState , StringBuilder pChar , int wCharSize , uint wFlags);

    [DllImport("user32.dll" , SetLastError = true)]
    public static extern IntPtr FindWindowEx(IntPtr parentHandle , IntPtr childAfter , string lpszClass , string lpszWindow);

    [DllImport("user32.dll")]
    public static extern IntPtr GetKeyboardLayout(uint idThread);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd , out RECT lpRect);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetClientRect(IntPtr hWnd , out RECT lpRect);

    [DllImport("user32.dll")]
    public static extern bool ClientToScreen(IntPtr hWnd , ref POINT lpPoint);

    [DllImport("user32.dll" , SetLastError = true , CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd , int Msg , IntPtr wParam , IntPtr lParam);

    [DllImport("user32.dll" , SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd , out uint processId);

    [DllImport("user32.dll")]
    public static extern bool GetGUIThreadInfo(uint idThread , ref GUITHREADINFO lpgui);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern int GetWindowLong(IntPtr hwnd , int index);

    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern int SetWindowLong(IntPtr hwnd , int index , int newStyle);
    [DllImport("user32.dll" , CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd , uint Msg , IntPtr wParam , IntPtr lParam);
    [DllImport("user32.dll" , CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd , uint Msg , IntPtr wParam , string lParam);
    [DllImport("gdi32.dll")]
    public static extern IntPtr SelectObject(IntPtr hdc , IntPtr hgdiobj);

    [DllImport("gdi32.dll")]
    public static extern bool GetTextExtentPoint32(IntPtr hdc , string lpString , int cbString , out SIZE lpSize);

    [DllImport("user32.dll")]
    public static extern IntPtr GetDC(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hWnd , IntPtr hDC);
    [DllImport("user32.dll" , CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd , uint Msg , int wParam , ref CHARRANGE lParam);

    [DllImport("user32.dll" , CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd , uint Msg , int wParam , string lParam);

    public const int WM_CHAR = 0x0102;
    public const int VK_CONTROL = 0x11;
    public const int VK_LEFT = 0x25;
    public const int VK_DELETE = 0x2E;
    public const int EM_POSFROMCHAR = 0xD6;
    public const uint WM_GETTEXT = 0x000D;
    public const uint WM_GETTEXTLENGTH = 0x000E;
    public const int WM_KEYDOWN = 0x0100;
    public const int WM_KEYUP = 0x0101;
    public const int KEYEVENTF_KEYUP = 0x0002;
    public const int WM_GETFONT = 0x0031;
    public const int GWL_EXSTYLE = -20;
    public const int WS_EX_TRANSPARENT = 0x00000020;
    public const int WS_EX_LAYERED = 0x00080000;
    public const int EM_GETSEL = 0x00B0;
    public const int EM_REPLACESEL = 0x00C2;
    public const int EM_EXGETSEL = 0x00B0;
    public const int EM_GETSELTEXT = 0x043E;
    public const int EM_SETSEL = 0x00B1;

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct GUITHREADINFO
    {
        public int cbSize;
        public int flags;
        public IntPtr hwndActive;
        public IntPtr hwndFocus;
        public IntPtr hwndCapture;
        public IntPtr hwndMenuOwner;
        public IntPtr hwndMoveSize;
        public IntPtr hwndCaret;
        public Rectangle rcCaret;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct Rectangle
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct SIZE
    {
        public int cx;
        public int cy;
    }


    [StructLayout(LayoutKind.Sequential)]
    public struct CHARRANGE
    {
        public int cpMin;
        public int cpMax;
    }
}
