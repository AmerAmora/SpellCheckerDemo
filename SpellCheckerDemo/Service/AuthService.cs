using Newtonsoft.Json;
using SpellCheckerDemo.Models;
using System.Diagnostics;
using System.Net.Sockets;
using System.Net;
using System.Web;
using System.Windows;
using System.Windows.Threading;
using SpellCheckerDemo;
using KeyboardTrackingApp;

public class AuthenticationService
{
    public bool IsAuthenticated { get; private set; } = false;

    public string Token { get; private set; }

    private const string LoginUrl = "https://localhost:7025/Home/Login";
    //private const string LoginUrl = "https://qalam.ai/auth/sign-in";
    private const string RedirectUrl = "http://localhost";
    private HttpListener listener;
    private int port;

    public AuthenticationService()
    {
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

    public async void ListenForCallback()
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
        await Application.Current.Dispatcher.InvokeAsync(() =>
        {
            SecureTokenStorage.StoreToken(accessToken);

            // Update the bearer token in MainWindow
            if (Application.Current.MainWindow is MainWindow mainWindow)
            {
                mainWindow.UpdateBearerToken(accessToken);
            }
            MessageBox.Show($"Login successful!");
            IsAuthenticated = true;
        });
    }

    public void OpenLoginPage()
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

    public void LogOut()
    {
        IsAuthenticated = false;
        SecureTokenStorage.StoreToken(string.Empty);
        Token = null;
    }
}
