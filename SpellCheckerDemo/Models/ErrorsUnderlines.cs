namespace SpellCheckerDemo.Models;

public class ErrorsUnderlines
{
    public List<(System.Windows.Point, double, string, List<string>, int, int)> SpellingErrors { get; set; } = new();

    public List<(System.Windows.Point, double, string, List<string>, int, int)> GrammarErrors { get; set; } = new();

    public List<(System.Windows.Point, double, string, List<string>, int, int)> PhrasingErrors { get; set; } = new();

    public List<(System.Windows.Point, double, string, List<string>, int, int)> TafqitErrors { get; set; } = new();

    public List<(System.Windows.Point, double, string, List<string>, int, int)> OtherErrors { get; set; } = new();

    public List<(System.Windows.Point, double, string, List<string>, int, int)> TermErrors { get; set; } = new();

    public int TotalErrorCount =>
    SpellingErrors.Count + GrammarErrors.Count + PhrasingErrors.Count + TafqitErrors.Count + TermErrors.Count;
}
