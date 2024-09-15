using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SpellCheckerDemo
{
    public partial class OverlayWindow : Window
    {
        public OverlayWindow()
        {
            InitializeComponent();

            // Make the window transparent and click-through
            Topmost = true;
            ShowInTaskbar = false;

            // Make the window click-through
            var hwnd = new WindowInteropHelper(this).Handle;
            var extendedStyle = Native.GetWindowLong(hwnd , Native.GWL_EXSTYLE);
            Native.SetWindowLong(hwnd , Native.GWL_EXSTYLE , extendedStyle | Native.WS_EX_TRANSPARENT);
        }

        public void DrawUnderline(Rect textRect)
        {
            // Clear previous drawings
            canvas.Children.Clear();

            // Create a red line
            var line = new System.Windows.Shapes.Line
            {
                X1 = textRect.Left ,
                Y1 = textRect.Bottom ,
                X2 = textRect.Right ,
                Y2 = textRect.Bottom ,
                Stroke = Brushes.Red ,
                StrokeThickness = 2
            };

            canvas.Children.Add(line);
        }
    }

    public static class Native
    {
        public const int GWL_EXSTYLE = -20;
        public const int WS_EX_TRANSPARENT = 0x00000020;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int GetWindowLong(IntPtr hwnd , int index);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SetWindowLong(IntPtr hwnd , int index , int newStyle);
    }
}
