﻿using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Media3D;
using System.Windows.Media;
using System.Windows.Threading;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TrayNotify;
using System.Diagnostics;

namespace SpellCheckerDemo
{
    public partial class FloatingPointWindow : Window
    {
        private bool isUserDragging = false;
        private static bool isAuthenticated = false;
        private const string LoginUrl = "https://qalam.ai/auth/sign-in";

        public FloatingPointWindow()
        {
            InitializeComponent();
            this.Topmost = true;

            this.MouseLeftButtonDown += (s , e) =>
            {
                isUserDragging = true;
                this.DragMove();
            };

            this.MouseDoubleClick += (s , e) =>
            {
                isUserDragging = false;
                HandleDoubleClick(s , e);
            };

            SetupCaretTracking();
        }

        private void HandleDoubleClick(object sender , MouseButtonEventArgs e)
        {
            if (!isAuthenticated)
            {
                OpenLoginPage();
                isAuthenticated = true; // In a real application, set this after successful authentication
            }
            else
            {
                ToggleMainWindowVisibility();
            }
        }

        private void OpenLoginPage()
        {
            try
            {
                Process.Start(new ProcessStartInfo(LoginUrl) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open login page: {ex.Message}" , "Error" , MessageBoxButton.OK , MessageBoxImage.Error);
            }
        }

        private void ToggleMainWindowVisibility()
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

        private void SetupCaretTracking()
        {
            DispatcherTimer timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            timer.Tick += Timer_Tick;
            timer.Start();
        }

        private void Timer_Tick(object sender , EventArgs e)
        {
            if (isUserDragging) return;

            IntPtr foregroundWindow = NativeMethods.GetForegroundWindow();
            uint processId;
            uint threadId = NativeMethods.GetWindowThreadProcessId(foregroundWindow , out processId);
            NativeMethods.GUITHREADINFO guiThreadInfo = new NativeMethods.GUITHREADINFO();
            guiThreadInfo.cbSize = Marshal.SizeOf(guiThreadInfo);
            if (NativeMethods.GetGUIThreadInfo(threadId , ref guiThreadInfo))
            {
                NativeMethods.Rectangle caretRect = guiThreadInfo.rcCaret;
                NativeMethods.POINT screenPoint;
                screenPoint.X = caretRect.Left;
                screenPoint.Y = caretRect.Top;
                NativeMethods.ClientToScreen(guiThreadInfo.hwndCaret , ref screenPoint);
                this.Left = screenPoint.X + 10;
                this.Top = screenPoint.Y + 10;
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