namespace SpellCheckerDemo.Models;

public class FlaggedToken
{
    public string OriginalWord { get; set; }
    public int start_index { get; set; }
    public int end_index { get; set; }
    public List<Suggestion> suggestions { get; set; }
    public string ArabicReason { get; set; }
    public string Lang { get; set; }
}
