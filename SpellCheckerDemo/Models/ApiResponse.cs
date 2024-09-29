namespace SpellCheckerDemo.Models;

public class ApiResponse
{
    public SpellCheckResponse spellCheckResponse { get; set; }
    public GrammarResponse grammarResponse { get; set; }
    public PhrasingResponse phrasingResponse { get; set; }
    public List<object> termSuggestions { get; set; }
    public TafqitResponse tafqitResponse { get; set; }
    public OtherSuggestions otherSuggestions { get; set; }
    public int teaserBalance { get; set; }
}
