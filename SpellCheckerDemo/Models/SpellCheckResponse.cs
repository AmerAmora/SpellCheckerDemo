using SpellCheckerDemo.Models;

namespace SpellCheckerDemo;

public class SpellCheckResponse
{
    public string Version { get; set; }
    public Results results { get; set; }
}
