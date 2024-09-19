using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace SpellCheckerDemo
{
    public partial class OverlayWindow : Window
    {
        public OverlayWindow()
        {
            InitializeComponent();
            Topmost = true;
            ShowInTaskbar = false;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;

            var hwnd = new WindowInteropHelper(this).Handle;
            var extendedStyle = Native.GetWindowLong(hwnd , Native.GWL_EXSTYLE);
            Native.SetWindowLong(hwnd , Native.GWL_EXSTYLE , extendedStyle | Native.WS_EX_TRANSPARENT | Native.WS_EX_LAYERED);
        }

        public void DrawUnderline(Point screenPosition , double width)
        {
            canvas.Children.Clear();

            var line = new System.Windows.Shapes.Line
            {
                X1 = 0 ,
                Y1 = 2 ,
                X2 = width ,
                Y2 = 2 ,
                Stroke = Brushes.Red ,
                StrokeThickness = 2
            };

            canvas.Children.Add(line);

            // Position the overlay window
            Left = screenPosition.X;
            Top = screenPosition.Y;
            Width = width;
            Height = 4; // Height of the underline
        }
    }

    public static class Native
    {
        public const int GWL_EXSTYLE = -20;
        public const int WS_EX_TRANSPARENT = 0x00000020;
        public const int WS_EX_LAYERED = 0x00080000;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hwnd , int index);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SetWindowLong(IntPtr hwnd , int index , int newStyle);
    }
}