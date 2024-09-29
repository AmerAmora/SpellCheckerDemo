namespace SpellCheckerDemo;

public class Suggestion
{
    public string text { get; set; }
    public string confidence { get; set; }
    public List<string> reasons { get; set; }
}
