namespace SpellCheckerDemo.Models;

public class ErrorsUnderlines
{
    public List<(System.Windows.Point point, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> SpellingErrors { get; set; } = new();

    public List<(System.Windows.Point point, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> GrammarError { get; set; } = new();

    public List<(System.Windows.Point point, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> PhrasingErrors { get; set; } = new();

    public List<(System.Windows.Point point, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> TafqitErrors { get; set; } = new();

    public List<(System.Windows.Point point, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> OtherErrors { get; set; } = new();

    public List<(System.Windows.Point point, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> TermErrors { get; set; } = new();
}
