using KeyboardTrackingApp;
using Newtonsoft.Json;
using SpellCheckerDemo.Models;
using System;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text.Json.Serialization;
using System.Web;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;

namespace SpellCheckerDemo
{
    public partial class FloatingPointWindow : Window
    {
        private bool isUserDragging = false;
        private DateTime lastClickTime;
        private HttpListener listener;
        private readonly AuthenticationService _authenticationService;
        
        public FloatingPointWindow(AuthenticationService authenticationService)
        {
            InitializeComponent();
            this.Topmost = true;

            this.MouseLeftButtonDown += (s , e) =>
            {
                // Check if the left mouse button is actually pressed before calling DragMove.
                if (e.LeftButton == MouseButtonState.Pressed)
                {
                    isUserDragging = true;
                    this.DragMove();
                }
            };

            this.MouseDoubleClick += HandleDoubleClick;

            ErrorCountGrid.MouseLeftButtonDown += ErrorCountGrid_MouseLeftButtonDown;

            SetupCaretTracking();
            _authenticationService = authenticationService;
        }

        private void HandleDoubleClick(object sender , MouseButtonEventArgs e)
        {
            isUserDragging = false;
            var currentToken = SecureTokenStorage.RetrieveToken();
            if (!string.IsNullOrEmpty(currentToken)) 
                _authenticationService.IsAuthenticated=true;
            
            if (!_authenticationService.IsAuthenticated)
            {
               _authenticationService.OpenLoginPage();
            }
            else
            {
                //ToggleMainWindowVisibility();
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

        public void UpdateErrorCount(int errorCount)
        {
            Dispatcher.Invoke(() =>
            {
                if (errorCount > 0)
                {
                    // Show error count
                    LogoImage.Visibility = Visibility.Collapsed;
                    ErrorCountGrid.Visibility = Visibility.Visible;
                    ErrorCountText.Text = errorCount.ToString();
                    AllErrorsCount.Text = $"عدد الاخطاء المكتشفة في النص : {errorCount} اخطاء";
                }
                else
                {
                    if (FixAllGrid.Visibility == Visibility.Visible)
                        LogoImage.Visibility = Visibility.Collapsed;
                    else 
                        LogoImage.Visibility = Visibility.Visible;

                    ErrorCountGrid.Visibility = Visibility.Collapsed;
                }
            });
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
                NativeMethods.POINT screenPoint = new NativeMethods.POINT
                {
                    X = caretRect.Left ,
                    Y = caretRect.Top
                };
                NativeMethods.ClientToScreen(guiThreadInfo.hwndCaret , ref screenPoint);

                this.Left = screenPoint.X + 10;
                this.Top = screenPoint.Y + 10;
            }
        }

        private void ErrorCountGrid_MouseLeftButtonDown(object sender , MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                var currentTime = DateTime.Now;

                // Check if the time since the last click is less than the double-click threshold
                if (( currentTime - lastClickTime ).TotalMilliseconds <= System.Windows.Forms.SystemInformation.DoubleClickTime)
                {
                    // Handle the double-click action here
                    ErrorCountGrid.Visibility = Visibility.Collapsed;
                    LogoImage.Visibility = Visibility.Collapsed;

                    FixAllGrid.Visibility = Visibility.Visible;
                    MainWindow.Width = 290;
                    MainWindow.Height = 150;
                }

                // Update the last click time
                lastClickTime = currentTime;
            }
        }

        private void FixAllErrorsButton_Click(object sender , RoutedEventArgs e)
        {
            // Call method in MainWindow to apply all suggestions
            ( (MainWindow)Application.Current.MainWindow ).ApplyAllSuggestions();

            // Hide the button and show the error count grid again
            FixAllGrid.Visibility = Visibility.Collapsed;
            ErrorCountGrid.Visibility = Visibility.Visible;
            MainWindow.Width = 70;
            MainWindow.Height = 70;
        }

        private void CloseButton_Click(object sender , RoutedEventArgs e)
        {
            FixAllGrid.Visibility = Visibility.Collapsed;
            ErrorCountGrid.Visibility = Visibility.Visible;
        }
    }
}
