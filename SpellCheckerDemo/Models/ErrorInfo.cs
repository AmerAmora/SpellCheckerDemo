using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace SpellCheckerDemo.Models
{
    public class ErrorInfo
    {
        public Point ScreenPosition { get; }
        public double Width { get; }
        public string IncorrectWord { get; }
        public List<string> Suggestions { get; }
        public int StartIndex { get; }
        public int EndIndex { get; }

        public ErrorInfo(Point screenPosition , double width , string incorrectWord , List<string> suggestions , int startIndex , int endIndex)
        {
            ScreenPosition = screenPosition;
            Width = width;
            IncorrectWord = incorrectWord;
            Suggestions = suggestions;
            StartIndex = startIndex;
            EndIndex = endIndex;
        }
    }
}
