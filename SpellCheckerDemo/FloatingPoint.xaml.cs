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
        private static bool isAuthenticated = false;
        //private const string LoginUrl = "https://qalam.ai/auth/sign-in";
        private const string LoginUrl = "https://localhost:7025/Home/Login";
        private DateTime lastClickTime;
        private const string RedirectUrl = "http://localhost";
        private HttpListener listener;
        private int port;

        public FloatingPointWindow()
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
            InitializeListener();
        }

        private void InitializeListener()
        {
            listener = new HttpListener();
            port = FindAvailablePort();
            listener.Prefixes.Add($"{RedirectUrl}:{port}/");
        }

        private int FindAvailablePort()
        {
            TcpListener l = new TcpListener(IPAddress.Loopback , 0);
            l.Start();
            int port = ( (IPEndPoint)l.LocalEndpoint ).Port;
            l.Stop();
            return port;
        }

        private async void ListenForCallback()
        {
            listener.Start();
            HttpListenerContext context = await listener.GetContextAsync();
            HttpListenerRequest request = context.Request;

            // Extract token information
            var query = HttpUtility.ParseQueryString(request.Url.Query);
            string data = query["data"];
            var apiResult = JsonConvert.DeserializeObject<AuthResponse>(data);
            var accessToken = apiResult.token;

            // Send a response to close the browser tab
            string responseData = "Login successful! You can close this tab and return to the application.";
            byte[] buffer = System.Text.Encoding.UTF8.GetBytes(responseData);
            HttpListenerResponse response = context.Response;
            response.ContentLength64 = buffer.Length;
            System.IO.Stream output = response.OutputStream;
            await output.WriteAsync(buffer , 0 , buffer.Length);
            output.Close();
            listener.Stop();

            // Handle successful login in your WPF application
            await Dispatcher.InvokeAsync(() =>
            {
                SecureTokenStorage.StoreToken(accessToken);

                // Update the bearer token in MainWindow
                if (Application.Current.MainWindow is MainWindow mainWindow)
                {
                    mainWindow.UpdateBearerToken(accessToken);
                }
                MessageBox.Show($"Login successful!");
                isAuthenticated = true;
            });
        }

        private void HandleDoubleClick(object sender , MouseButtonEventArgs e)
        {
            isUserDragging = false;

            if (!isAuthenticated)
            {
                OpenLoginPage();
                isAuthenticated = true; // In a real application, set this after successful authentication
            }
            else
            {
                //ToggleMainWindowVisibility();
            }
        }

        private void OpenLoginPage()
        {
            try
            {
                string fullLoginUrl = $"{LoginUrl}?redirect_uri={RedirectUrl}:{port}/";
                Process.Start(new ProcessStartInfo(fullLoginUrl) { UseShellExecute = true });

                ListenForCallback();
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
