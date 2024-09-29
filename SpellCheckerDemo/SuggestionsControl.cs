using System.Windows.Controls;
using System.Windows;
using System.Windows.Media;

namespace SpellCheckerDemo;

public class SuggestionsControl : StackPanel
{
    public event EventHandler SuggestionAccepted;
    public event EventHandler SuggestionCancelled;

    private TextBlock _suggestionTextBlock;
    private Button _acceptButton;
    private Button _cancelButton;

    public SuggestionsControl()
    {
        _suggestionTextBlock = new TextBlock
        {
            Padding = new Thickness(5) ,
            Background = Brushes.LightYellow
        };

        _acceptButton = new Button
        {
            Content = "Accept" ,
            Margin = new Thickness(5) ,
            Padding = new Thickness(5)
        };
        _acceptButton.Click += (s , e) => SuggestionAccepted?.Invoke(this , EventArgs.Empty);

        _cancelButton = new Button
        {
            Content = "Cancel" ,
            Margin = new Thickness(5) ,
            Padding = new Thickness(5)
        };
        _cancelButton.Click += (s , e) => SuggestionCancelled?.Invoke(this , EventArgs.Empty);

        Children.Add(_suggestionTextBlock);
        Children.Add(_acceptButton);
        Children.Add(_cancelButton);
    }

    public void SetSuggestion(string suggestion)
    {
        if (string.IsNullOrEmpty(suggestion))
        {
            Visibility = Visibility.Collapsed;
        }
        else
        {
            _suggestionTextBlock.Text = $"Suggested: {suggestion}";
            Visibility = Visibility.Visible;
        }
    }
}
