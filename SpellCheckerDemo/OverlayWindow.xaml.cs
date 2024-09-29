using SpellCheckerDemo.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
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
        private StackPanel _suggestionPanel;
        private IntPtr _hookID = IntPtr.Zero;
        private NativeMethods.WinEventDelegate _winEventDelegate;
        private bool _isPopupJustOpened = false;

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
            var extendedStyle = NativeMethods.GetWindowLong(hwnd , NativeMethods.GWL_EXSTYLE);
            NativeMethods.SetWindowLong(hwnd , NativeMethods.GWL_EXSTYLE , extendedStyle | NativeMethods.WS_EX_TRANSPARENT | NativeMethods.WS_EX_LAYERED);

            InitializeSuggestionPopup();
            SetupWindowsHook();
        }

        private void SetupWindowsHook()
        {
            _winEventDelegate = new NativeMethods.WinEventDelegate(WinEventProc);
            _hookID = NativeMethods.SetWinEventHook(NativeMethods.EVENT_SYSTEM_FOREGROUND , NativeMethods.EVENT_SYSTEM_FOREGROUND , IntPtr.Zero ,
                _winEventDelegate , 0 , 0 , NativeMethods.WINEVENT_OUTOFCONTEXT);
            Debug.WriteLine($"Windows hook set up. Hook ID: {_hookID}");
        }

        private void WinEventProc(IntPtr hWinEventHook , uint eventType , IntPtr hwnd , int idObject , int idChild , uint dwEventThread , uint dwmsEventTime)
        {
            if (eventType == NativeMethods.EVENT_SYSTEM_FOREGROUND)
            {
                Dispatcher.Invoke(() =>
                {
                    if (_suggestionPopup.IsOpen && !_isPopupJustOpened)
                    {
                        Debug.WriteLine("Closing suggestion popup due to foreground window change.");
                        _suggestionPopup.IsOpen = false;
                    }
                    _isPopupJustOpened = false;
                });
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            NativeMethods.UnhookWinEvent(_hookID);
            Debug.WriteLine("Windows hook unhooked.");
            base.OnClosed(e);
        }

        private void InitializeSuggestionPopup()
        {
            _suggestionPopup = new Popup
            {
                AllowsTransparency = true ,
                Placement = PlacementMode.MousePoint ,
                StaysOpen = true
            };

            var border = new Border
            {
                Background = new SolidColorBrush(Color.FromArgb(240 , 255 , 255 , 255)) ,
                BorderBrush = new SolidColorBrush(Color.FromRgb(200 , 200 , 200)) ,
                BorderThickness = new Thickness(1) ,
                CornerRadius = new CornerRadius(4) ,
                Padding = new Thickness(5) ,
                Effect = new System.Windows.Media.Effects.DropShadowEffect
                {
                    Color = Colors.Gray ,
                    Direction = 315 ,
                    ShadowDepth = 5 ,
                    BlurRadius = 10 ,
                    Opacity = 0.3
                }
            };

            _suggestionPanel = new StackPanel
            {
                Orientation = Orientation.Vertical
            };

            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto ,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled ,
                MaxHeight = 200 ,
                Content = _suggestionPanel
            };

            border.Child = scrollViewer;
            _suggestionPopup.Child = border;
        }

        private UIElement CreateSuggestionRectangle(string text)
        {
            var border = new Border
            {
                Background = Brushes.White ,
                BorderBrush = new SolidColorBrush(Color.FromRgb(200 , 200 , 200)) ,
                BorderThickness = new Thickness(1) ,
                CornerRadius = new CornerRadius(4) ,
                Padding = new Thickness(10 , 5 , 10 , 5) ,
                Margin = new Thickness(0 , 0 , 0 , 5)
            };

            var textBlock = new TextBlock
            {
                Text = text ,
                Foreground = Brushes.Green ,
                FontWeight = FontWeights.SemiBold
            };

            border.Child = textBlock;
            border.MouseLeftButtonDown += Suggestion_MouseLeftButtonDown;
            return border;
        }

        private void Suggestion_MouseLeftButtonDown(object sender , MouseButtonEventArgs e)
        {
            if (sender is Border border &&
                border.Child is TextBlock textBlock &&
                _suggestionPopup.Tag is (string incorrectWord, List<string> suggestions, int startIndex, int endIndex))
            {
                string selectedSuggestion = textBlock.Text;
                Debug.WriteLine($"Suggestion selected: {selectedSuggestion}");
                SuggestionAccepted?.Invoke(this , (selectedSuggestion, startIndex, endIndex));
                _suggestionPopup.IsOpen = false;
            }
        }

        private void Line_MouseLeftButtonDown(object sender , MouseButtonEventArgs e)
        {
            if (sender is Line line && line.Tag is (string incorrectWord, List<string> suggestions, int startIndex, int endIndex))
            {
                Debug.WriteLine($"Underline clicked. Incorrect word: {incorrectWord}");
                _suggestionPanel.Children.Clear();
                foreach (var suggestion in suggestions)
                {
                    _suggestionPanel.Children.Add(CreateSuggestionRectangle(suggestion));
                }
                _suggestionPopup.Tag = (incorrectWord, suggestions, startIndex, endIndex);
                _isPopupJustOpened = true;
                _suggestionPopup.IsOpen = true;
                Debug.WriteLine("Suggestion popup opened.");

                e.Handled = true;
            }
        }

        public void DrawUnderlines(ErrorsUnderlines errors)
        {
            canvas.Children.Clear();
            _underlines = new List<(Point, double, string, List<string>, int, int)>();

            DrawErrorType(errors.SpellingErrors , Brushes.Red);
            DrawErrorType(errors.GrammarError , Brushes.LightBlue);
            DrawErrorType(errors.PhrasingErrors , Brushes.Yellow);
            DrawErrorType(errors.TafqitErrors , Brushes.LightGreen);
            DrawErrorType(errors.TermErrors , Brushes.Purple);

            Debug.WriteLine($"Total underlines drawn: {_underlines.Count}");
        }

        private void DrawErrorType(List<(Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> errorList , Brush color)
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
                Debug.WriteLine($"Underline drawn for '{incorrectWord}' at ({line.X1}, {line.Y1})");
            }
        }
    }
}