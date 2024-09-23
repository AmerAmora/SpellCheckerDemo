using SpellCheckerDemo.Models;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Shapes;

namespace SpellCheckerDemo
{
    public partial class OverlayWindow : Window
    {
        private List<(Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> _underlines;
        private Popup _suggestionPopup;
        private ListBox _suggestionListBox;

        public event EventHandler<(string suggestion, int startIndex, int endIndex)> SuggestionAccepted;

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

            InitializeSuggestionPopup();
        }

        private void InitializeSuggestionPopup()
        {
            _suggestionPopup = new Popup
            {
                AllowsTransparency = true ,
                Placement = PlacementMode.MousePoint
            };

            _suggestionListBox = new ListBox
            {
                Margin = new Thickness(5) ,
                BorderThickness = new Thickness(1) ,
                BorderBrush = Brushes.Gray ,
                Background = Brushes.White
            };
            _suggestionListBox.SelectionChanged += SuggestionListBox_SelectionChanged;

            _suggestionPopup.Child = _suggestionListBox;
        }

        public void DrawUnderlines(ErrorsUnderlines errors)
        {
            canvas.Children.Clear();
            _underlines = new List<(System.Windows.Point, double, string, List<string>, int, int)>();

            DrawErrorType(errors.SpellingErrors , Brushes.Red);
            DrawErrorType(errors.GrammarError , Brushes.LightBlue);
            DrawErrorType(errors.PhrasingErrors , Brushes.Yellow);
            DrawErrorType(errors.TafqitErrors , Brushes.LightGreen);
            DrawErrorType(errors.TermErrors , Brushes.Purple);
        }

        private void DrawErrorType(List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> errorList , Brush color)
        {
            foreach (var error in errorList)
            {
                var (screenPosition, width, incorrectWord, suggestions, startIndex, endIndex) = error;
                _underlines.Add(error);

                var line = new Line
                {
                    X1 = screenPosition.X - Left ,
                    Y1 = screenPosition.Y - Top + 2 ,
                    X2 = screenPosition.X - Left + width ,
                    Y2 = screenPosition.Y - Top + 2 ,
                    Stroke = color ,
                    StrokeThickness = 2 ,
                    Tag = (incorrectWord, suggestions, startIndex, endIndex)
                };

                line.MouseLeftButtonDown += Line_MouseLeftButtonDown;
                canvas.Children.Add(line);
            }
        }
        private void Line_MouseLeftButtonDown(object sender , MouseButtonEventArgs e)
        {
            if (sender is Line line && line.Tag is (string incorrectWord, List<string> suggestions, int startIndex, int endIndex))
            {
                _suggestionListBox.ItemsSource = suggestions;
                _suggestionPopup.Tag = (incorrectWord, suggestions, startIndex, endIndex);
                _suggestionPopup.IsOpen = true;
            }
        }

        private void SuggestionListBox_SelectionChanged(object sender , SelectionChangedEventArgs e)
        {
            if (_suggestionListBox.SelectedItem is string selectedSuggestion &&
                _suggestionPopup.Tag is (string incorrectWord, List<string> suggestions, int startIndex, int endIndex))
            {
                SuggestionAccepted?.Invoke(this , (selectedSuggestion, startIndex, endIndex));
                _suggestionPopup.IsOpen = false;
            }
        }

        private void AcceptButton_Click(object sender , RoutedEventArgs e)
        {
            if (_suggestionPopup.Tag is (string incorrectWord, List<string> suggestions, int startIndex, int endIndex) &&
                _suggestionListBox.SelectedItem is string selectedSuggestion)
            {
                SuggestionAccepted?.Invoke(this , (selectedSuggestion, startIndex, endIndex));
            }
            _suggestionPopup.IsOpen = false;
        }

        private void CancelButton_Click(object sender , RoutedEventArgs e)
        {
            _suggestionPopup.IsOpen = false;
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