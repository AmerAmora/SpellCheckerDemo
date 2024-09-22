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
        private List<(Point screenPosition, double width, string incorrectWord, string suggestion)> _underlines;
        private Popup _suggestionPopup;
        private TextBlock _suggestionTextBlock;
        private Button _acceptButton;
        private Button _cancelButton;

        public event EventHandler<string> SuggestionAccepted;

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
        }

        private void InitializeSuggestionPopup()
        {
            _suggestionPopup = new Popup
            {
                AllowsTransparency = true ,
                Placement = System.Windows.Controls.Primitives.PlacementMode.MousePoint
            };

            var popupContent = new StackPanel
            {
                Background = Brushes.White ,
                Orientation = Orientation.Vertical
            };

            _suggestionTextBlock = new TextBlock
            {
                Margin = new Thickness(5)
            };

            _acceptButton = new Button
            {
                Content = "Accept" ,
                Margin = new Thickness(5)
            };
            _acceptButton.Click += AcceptButton_Click;

            _cancelButton = new Button
            {
                Content = "Cancel" ,
                Margin = new Thickness(5)
            };
            _cancelButton.Click += CancelButton_Click;

            popupContent.Children.Add(_suggestionTextBlock);
            popupContent.Children.Add(_acceptButton);
            popupContent.Children.Add(_cancelButton);

            _suggestionPopup.Child = popupContent;
        }

        public void DrawUnderlines(List<(Point screenPosition, double width, string incorrectWord, string suggestion)> underlines)
        {
            _underlines = underlines;
            canvas.Children.Clear();

            foreach (var (screenPosition, width, incorrectWord, suggestion) in underlines)
            {
                var line = new Line
                {
                    X1 = screenPosition.X - Left ,
                    Y1 = screenPosition.Y - Top + 2 ,
                    X2 = screenPosition.X - Left + width ,
                    Y2 = screenPosition.Y - Top + 2 ,
                    Stroke = Brushes.Red ,
                    StrokeThickness = 2 ,
                    Tag = (incorrectWord, suggestion)
                };

                line.MouseLeftButtonDown += Line_MouseLeftButtonDown;
                canvas.Children.Add(line);
            }
        }

        private void Line_MouseLeftButtonDown(object sender , MouseButtonEventArgs e)
        {
            if (sender is Line line && line.Tag is (string incorrectWord, string suggestion))
            {
                _suggestionTextBlock.Text = $"Suggestion: {suggestion}";
                _suggestionPopup.Tag = (incorrectWord, suggestion);
                _suggestionPopup.IsOpen = true;
            }
        }

        private void AcceptButton_Click(object sender , RoutedEventArgs e)
        {
            if (_suggestionPopup.Tag is (string incorrectWord, string suggestion))
            {
                SuggestionAccepted?.Invoke(this , suggestion);
            }
            _suggestionPopup.IsOpen = false;
        }

        private void CancelButton_Click(object sender , RoutedEventArgs e)
        {
            _suggestionPopup.IsOpen = false;
        }
    }
}