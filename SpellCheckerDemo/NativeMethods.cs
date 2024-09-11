using System;
using System.Runtime.InteropServices;
using System.Text;

public static class NativeMethods
{
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern void keybd_event(byte bVk , byte bScan , uint dwFlags , UIntPtr dwExtraInfo);

    [DllImport("user32.dll")]
    public static extern bool AttachThreadInput(uint idAttach , uint idAttachTo , bool fAttach);

    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd , IntPtr ProcessId);

    [DllImport("user32.dll")]
    public static extern bool GetKeyboardState(byte[] lpKeyState);

    [DllImport("user32.dll")]
    public static extern int ToUnicode(uint wVirtKey , uint wScanCode , byte[] lpKeyState ,
        [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff , int cchBuff , uint wFlags);

    [DllImport("user32.dll")]
    public static extern uint MapVirtualKey(uint uCode , uint uMapType);

    [DllImport("kernel32.dll")]
    public static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    public static extern IntPtr SetWinEventHook(uint eventMin , uint eventMax , IntPtr hmodWinEventProc ,
        WinEventDelegate lpfnWinEventProc , uint idProcess , uint idThread , uint dwFlags);

    [DllImport("user32.dll")]
    public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    [DllImport("user32.dll")]
    public static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    public static extern int GetWindowText(IntPtr hWnd , StringBuilder lpString , int nMaxCount);

    [DllImport("user32.dll" , CharSet = CharSet.Auto)]
    public static extern IntPtr SendMessage(IntPtr hWnd , uint Msg , IntPtr wParam , [Out] StringBuilder lParam);

    [DllImport("user32.dll" , SetLastError = true)]
    public static extern IntPtr FindWindowEx(IntPtr hwndParent , IntPtr hwndChildAfter , string lpszClass , string lpszWindow);

    [DllImport("kernel32.dll" , CharSet = CharSet.Auto , SetLastError = true)]
    public static extern IntPtr GetModuleHandle(string lpModuleName);

    [DllImport("user32.dll" , CharSet = CharSet.Auto , SetLastError = true)]
    public static extern IntPtr SetWindowsHookEx(int idHook , LowLevelKeyboardProc lpfn , IntPtr hMod , uint dwThreadId);

    [DllImport("user32.dll" , CharSet = CharSet.Auto , SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll" , CharSet = CharSet.Auto , SetLastError = true)]
    public static extern IntPtr CallNextHookEx(IntPtr hhk , int nCode , IntPtr wParam , IntPtr lParam);

    public delegate IntPtr LowLevelKeyboardProc(int nCode , IntPtr wParam , IntPtr lParam);
    public delegate void WinEventDelegate(IntPtr hWinEventHook , uint eventType , IntPtr hwnd , int idObject , int idChild , uint dwEventThread , uint dwmsEventTime);
}
