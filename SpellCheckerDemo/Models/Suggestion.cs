namespace SpellCheckerDemo;

public class Suggestion
{
    public string Text { get; set; }
    public string Confidence { get; set; }
    public List<string> Reasons { get; set; }
}
