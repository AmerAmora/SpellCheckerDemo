namespace SpellCheckerDemo.Models;

public class FlaggedToken
{
    public string OriginalWord { get; set; }
    public int StartIndex { get; set; }
    public int EndIndex { get; set; }
    public List<Suggestion> Suggestions { get; set; }
    public string ArabicReason { get; set; }
    public string Lang { get; set; }
}
