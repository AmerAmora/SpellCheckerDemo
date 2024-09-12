using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;

namespace SpellCheckerDemo
{
    public partial class FloatingPointWindow : Window
    {
        public FloatingPointWindow()
        {
            InitializeComponent();
            this.Topmost = true;

            // Allow the window to be dragged
            this.MouseLeftButtonDown += (s , e) => this.DragMove();

            // Toggle main window visibility on click
            this.MouseLeftButtonUp += ToggleMainWindowVisibility;

            SetupCaretTracking(); // Start caret tracking
        }

        private void ToggleMainWindowVisibility(object sender , MouseButtonEventArgs e)
        {
            if (Application.Current.MainWindow.IsVisible)
            {
                Application.Current.MainWindow.Hide();
            }
            else
            {
                Application.Current.MainWindow.Show();
                Application.Current.MainWindow.WindowState = WindowState.Normal;
                Application.Current.MainWindow.Activate();
            }
        }

        // Set up caret tracking
        private void SetupCaretTracking()
        {
            DispatcherTimer timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100) // Adjust the interval as needed
            };
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        // Track the caret position in external applications
        private void Timer_Tick(object sender , EventArgs e)
        {
            IntPtr foregroundWindow = NativeMethods.GetForegroundWindow();
            uint processId;
            uint threadId = NativeMethods.GetWindowThreadProcessId(foregroundWindow , out processId);
            NativeMethods.GUITHREADINFO guiThreadInfo = new NativeMethods.GUITHREADINFO();
            guiThreadInfo.cbSize = Marshal.SizeOf(guiThreadInfo);
            if (NativeMethods.GetGUIThreadInfo(threadId , ref guiThreadInfo))
            {
                NativeMethods.Rectangle caretRect = guiThreadInfo.rcCaret;
                // Convert to screen coordinates
                NativeMethods.POINT screenPoint;
                screenPoint.X = caretRect.Left;
                screenPoint.Y = caretRect.Top;
                NativeMethods.ClientToScreen(guiThreadInfo.hwndCaret , ref screenPoint);
                // Set the floating window position
                this.Left = screenPoint.X + 10; // Adjust offset based on icon size and DPI
                this.Top = screenPoint.Y + 10;  // Adjust offset based on icon size and DPI
            }
        }
    }

    public static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll" , SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd , out uint processId);

        [DllImport("user32.dll")]
        public static extern bool GetGUIThreadInfo(uint idThread , ref GUITHREADINFO lpgui);

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

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ClientToScreen(IntPtr hWnd , ref POINT lpPoint);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }
    }
}
