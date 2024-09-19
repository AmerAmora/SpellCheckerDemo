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

        public void DrawUnderlines(List<(Point screenPosition, double width)> underlines)
        {
            canvas.Children.Clear();

            foreach (var (screenPosition, width) in underlines)
            {
                var line = new System.Windows.Shapes.Line
                {
                    X1 = screenPosition.X - Left ,
                    Y1 = screenPosition.Y - Top + 2 ,
                    X2 = screenPosition.X - Left + width ,
                    Y2 = screenPosition.Y - Top + 2 ,
                    Stroke = Brushes.Red ,
                    StrokeThickness = 2
                };

                canvas.Children.Add(line);
            }
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