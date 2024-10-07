using System.Diagnostics;
using System.Text;
using System.Windows;
using Keys = System.Windows.Forms.Keys;
using System.Windows.Threading;
using SpellCheckerDemo;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using SpellCheckerDemo.Models;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Word.Application;
using Document = Microsoft.Office.Interop.Word.Document;

namespace KeyboardTrackingApp
{
    public partial class MainWindow : Window
    {
        private GlobalKeyboardHook _keyboardHook;
        private StringBuilder _currentWord = new StringBuilder();
        private IntPtr _lastActiveWindowHandle;
        private StringBuilder _allText = new StringBuilder();
        private string _lastActiveWindowTitle = string.Empty;
        private DispatcherTimer _windowCheckTimer;
        private DispatcherTimer _processCheckTimer;
        private int? _notepadProcessId;
        private SuggestionsControl _suggestionsControl;
        private DispatcherTimer _contentSyncTimer;
        private FloatingPointWindow _floatingPoint;
        private OverlayWindow _overlay;
        private readonly HttpClient _httpClient;
        private readonly string _apiUrl = "https://api-stg.qalam.ai/test/go";
        private string _bearerToken = "";
        private string _documentId;
        private Screen _currentScreen;
        private AuthenticationService _authService;
        private int? _microsoftWordProcessId;

        public MainWindow()
        {
            InitializeComponent();
            _keyboardHook = new GlobalKeyboardHook();
            _keyboardHook.KeyPressed += OnKeyPressed;

            // Initialize SuggestionsControl
            _suggestionsControl = new SuggestionsControl();
            SuggestionsPopup.Child = _suggestionsControl;

            _windowCheckTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(100)
            };
            _windowCheckTimer.Tick += WindowCheckTimer_Tick;
            _windowCheckTimer.Start();

            _processCheckTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _processCheckTimer.Tick += ProcessCheckTimer_Tick;
            _processCheckTimer.Start();

            _contentSyncTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _contentSyncTimer.Tick += ContentSyncTimer_Tick;
            _contentSyncTimer.Start();
            _authService = new AuthenticationService();

            _floatingPoint = new FloatingPointWindow(_authService);
            _floatingPoint.Hide();

            _overlay = new OverlayWindow();
            _overlay.Show();
            _overlay.SuggestionAccepted += Overlay_SuggestionAccepted;

            this.Hide();
            _httpClient = new HttpClient();
            _documentId = Guid.NewGuid().ToString(); // Generate a unique document ID
        }

        public void UpdateBearerToken(string newToken)
        {
            _bearerToken = newToken;
            // Optionally, you can reinitialize the HttpClient with the new token
            InitializeHttpClient();
        }

        private void InitializeHttpClient()
        {
            _httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer" , _bearerToken);
        }
        private void Overlay_SuggestionAccepted(object sender , (string suggestion, int startIndex, int endIndex) e)
        {
            ReplaceWord(e.suggestion , e.startIndex , e.endIndex);
            CheckForIncorrectWords(); // Recheck for incorrect words after replacement
        }

        public async void ApplyAllSuggestions()
        {
            string text;
            ErrorsUnderlines errors = new ErrorsUnderlines();

            if (_microsoftWordProcessId.HasValue)
            {
                try
                {
                    Application wordApp = (Application)SpellCheckerDemo.ReplaceMarshall.GetActiveObject("Word.Application");
                    if (wordApp?.ActiveDocument == null)
                    {
                        System.Windows.MessageBox.Show("Failed to find Word document.");
                        return;
                    }

                    text = GetMicrosoftWordContent();
                    if (string.IsNullOrEmpty(text))
                    {
                        return;
                    }

                    var apiResponse = await GetSpellCheckResultsAsync(text);
                    if (apiResponse == null)
                    {
                        return;
                    }

                    // Get the Word window handle
                    IntPtr wordHandle = new IntPtr(wordApp.ActiveWindow.Hwnd);

                    errors.SpellingErrors = GetWordSpellingErrors(wordHandle , apiResponse , text);
                    //errors.GrammarError = GetWordGrammarErrors(wordHandle, apiResponse, text);
                    //errors.PhrasingErrors = GetWordPhrasingErrors(wordHandle, apiResponse, text);
                    //errors.TafqitErrors = GetWordTafqitErrors(wordHandle, apiResponse, text);
                    //errors.TermErrors = GetWordTermErrors(wordHandle, apiResponse, text);

                    Marshal.ReleaseComObject(wordApp);
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Error accessing Word: {ex.Message}");
                    return;
                }
            }
            else if (_notepadProcessId.HasValue)
            {
                IntPtr notepadHandle = FindNotepadWindow();
                if (notepadHandle == IntPtr.Zero)
                {
                    System.Windows.MessageBox.Show("Failed to find Notepad window.");
                    return;
                }

                IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);
                if (editHandle == IntPtr.Zero)
                {
                    System.Windows.MessageBox.Show("Failed to find Edit control in Notepad.");
                    return;
                }

                text = GetNotepadContent(notepadHandle);
                var apiResponse = await GetSpellCheckResultsAsync(text);

                if (editHandle != IntPtr.Zero)
                {
                    errors.SpellingErrors = GetSpellingErrors(editHandle , apiResponse , text);
                    errors.GrammarError = GetGrammarErrors(editHandle , apiResponse , text);
                    errors.PhrasingErrors = GetPhrasingErrors(editHandle , apiResponse , text);
                    errors.TafqitErrors = GetTafqitErrors(editHandle , apiResponse , text);
                    errors.TermErrors = GetTermErrors(editHandle , apiResponse , text);
                }
            }
            else
            {
                System.Windows.MessageBox.Show("No supported text editor is currently active.");
                return;
            }

            // Sort all errors by their start index in descending order to avoid position shifting
            var allErrors = new List<(string suggestion, int startIndex, int endIndex)>();

            void AddErrors(IEnumerable<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> errorList)
            {
                if (errorList != null)
                {
                    foreach (var error in errorList)
                    {
                        if (error.suggestions != null && error.suggestions.Any())
                        {
                            allErrors.Add((error.suggestions.First(), error.startIndex, error.endIndex));
                        }
                    }
                }
            }

            AddErrors(errors.SpellingErrors);
            AddErrors(errors.GrammarError);
            AddErrors(errors.PhrasingErrors);
            AddErrors(errors.TafqitErrors);
            AddErrors(errors.TermErrors);

            // Sort errors by start index in descending order to handle replacements from end to start
            allErrors = allErrors.OrderByDescending(e => e.startIndex).ToList();

            // Apply all replacements
            foreach (var error in allErrors)
            {
                if (!string.IsNullOrEmpty(error.suggestion))
                {
                    ReplaceWord(error.suggestion , error.startIndex , error.endIndex);
                    // Add a small delay between replacements to ensure proper processing
                    await Task.Delay(50);
                }
            }

            CheckForIncorrectWords();
        }

        private void ReplaceWord(string suggestion , int startIndex , int endIndex)
        {
            if (_microsoftWordProcessId.HasValue)
            {
                try
                {
                    Application wordApp = (Application)SpellCheckerDemo.ReplaceMarshall.GetActiveObject("Word.Application");
                    if (wordApp?.ActiveDocument != null)
                    {
                        Document doc = wordApp.ActiveDocument;

                        // Create a range that spans the text to be replaced
                        Microsoft.Office.Interop.Word.Range range = doc.Range(startIndex , endIndex + 1);

                        // Store the current selection
                        Microsoft.Office.Interop.Word.Selection currentSelection = wordApp.Selection;

                        // Select the text to be replaced
                        range.Select();

                        // Replace the text
                        wordApp.Selection.TypeText(suggestion);

                        // Restore the original selection
                        if (currentSelection != null)
                        {
                            currentSelection.Select();
                            Marshal.ReleaseComObject(currentSelection);
                        }

                        // Clean up COM objects
                        Marshal.ReleaseComObject(range);
                        Marshal.ReleaseComObject(doc);
                        Marshal.ReleaseComObject(wordApp);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error replacing text in Word: {ex.Message}");
                }
            }
            else if (_notepadProcessId.HasValue)
            {
                // Existing Notepad replacement code remains unchanged
                IntPtr notepadHandle = FindNotepadWindow();
                if (notepadHandle == IntPtr.Zero)
                {
                    System.Windows.MessageBox.Show("Failed to find Notepad window.");
                    return;
                }

                IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);
                if (editHandle == IntPtr.Zero)
                {
                    System.Windows.MessageBox.Show("Failed to find Edit control in Notepad.");
                    return;
                }

                // Select the incorrect word
                NativeMethods.SendMessage(editHandle , NativeMethods.EM_SETSEL , (IntPtr)startIndex , (IntPtr)( endIndex + 1 ));

                // Replace the selection with the suggestion
                foreach (char c in suggestion)
                {
                    NativeMethods.SendMessage(editHandle , NativeMethods.WM_CHAR , (IntPtr)c , IntPtr.Zero);
                }
            }
        }

        private IntPtr FindNotepadWindow()
        {
            IntPtr result = IntPtr.Zero;
            NativeMethods.EnumWindows((hWnd , lParam) =>
            {
                int processId;
                NativeMethods.GetWindowThreadProcessId(hWnd , out processId);
                if (processId == _notepadProcessId)
                {
                    result = hWnd;
                    return false; // Stop enumeration
                }
                return true; // Continue enumeration
            } , IntPtr.Zero);
            return result;
        }

        private async void CheckForIncorrectWords()
        {
            string text;

            // Get content based on active application
            if (_microsoftWordProcessId.HasValue)
            {
                text = GetMicrosoftWordContent();
            }
            else if (_notepadProcessId.HasValue)
            {
                text = _allText.ToString();
            }
            else
            {
                return;
            }

            if (string.IsNullOrEmpty(text))
            {
                UpdateErrorCount(0);
                _overlay.DrawUnderlines(new ErrorsUnderlines());
                return;
            }

            var apiResponse = await GetSpellCheckResultsAsync(text);
            if (apiResponse is null)
                return;
            if (apiResponse?.spellCheckResponse?.results?.flagged_tokens == null)
            {
                UpdateErrorCount(0);
                _overlay.DrawUnderlines(new ErrorsUnderlines()); // Clear all underlines
                Console.WriteLine("No flagged tokens found or API response was null.");
                return;
            }

            ErrorsUnderlines errors = new ErrorsUnderlines();

            if (_microsoftWordProcessId.HasValue)
            {
                Application wordApp = null;
                try
                {
                    wordApp = (Application)SpellCheckerDemo.ReplaceMarshall.GetActiveObject("Word.Application");
                    if (wordApp != null && wordApp.ActiveDocument != null)
                    {
                        var activeWindow = wordApp.ActiveWindow;
                        if (activeWindow != null)
                        {
                            // Get the window handle for Word
                            IntPtr wordHandle = new IntPtr(activeWindow.Hwnd);

                            errors.SpellingErrors = GetWordSpellingErrors(wordHandle , apiResponse , text);
                            //errors.GrammarError = GetWordGrammarErrors(wordHandle , apiResponse , text);
                            //errors.PhrasingErrors = GetWordPhrasingErrors(wordHandle , apiResponse , text);
                            //errors.TafqitErrors = GetWordTafqitErrors(wordHandle , apiResponse , text);
                            //errors.TermErrors = GetWordTermErrors(wordHandle , apiResponse , text);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error accessing Word: {ex.Message}");
                }
                finally
                {
                    if (wordApp != null)
                    {
                        Marshal.ReleaseComObject(wordApp);
                    }
                }
            }
            else if (_notepadProcessId.HasValue)
            {
                IntPtr notepadHandle = NativeMethods.GetForegroundWindow();
                IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);

                if (editHandle != IntPtr.Zero)
                {
                    errors.SpellingErrors = GetSpellingErrors(editHandle , apiResponse , text);
                    errors.GrammarError = GetGrammarErrors(editHandle , apiResponse , text);
                    errors.PhrasingErrors = GetPhrasingErrors(editHandle , apiResponse , text);
                    errors.TafqitErrors = GetTafqitErrors(editHandle , apiResponse , text);
                    errors.TermErrors = GetTermErrors(editHandle , apiResponse , text);
                }
            }

            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                var count = errors.SpellingErrors.Count + errors.GrammarError.Count +
                            errors.PhrasingErrors.Count + errors.TafqitErrors.Count +
                            errors.TermErrors.Count;
                _overlay.DrawUnderlines(errors);
                UpdateErrorCount(count);
            });
        }

        #region GetErrors
        private List<(System.Windows.Point screenPosition, double width, string incorrectWord,
            List<string> suggestions, int startIndex, int endIndex)> GetWordSpellingErrors(
            IntPtr wordHandle , ApiResponse apiResponse , string text)
        {
            var errors = new List<(System.Windows.Point, double, string, List<string>, int, int)>();

            try
            {
                Application wordApp = (Application)SpellCheckerDemo.ReplaceMarshall.GetActiveObject("Word.Application");
                if (wordApp?.ActiveDocument == null) return errors;

                foreach (var flaggedToken in apiResponse.spellCheckResponse.results.flagged_tokens)
                {
                    int startIndex = flaggedToken.start_index;
                    int endIndex = flaggedToken.end_index;

                    if (startIndex < 0 || startIndex >= text.Length || endIndex > text.Length || endIndex <= startIndex)
                        continue;

                    string incorrectWord = text.Substring(startIndex , endIndex - startIndex);
                    var range = wordApp.ActiveDocument.Range(startIndex , endIndex);

                    // Get screen coordinates for the range
                    var pointsLeft = range.Information[Microsoft.Office.Interop.Word.WdInformation.wdHorizontalPositionRelativeToPage];
                    var pointsTop = range.Information[Microsoft.Office.Interop.Word.WdInformation.wdVerticalPositionRelativeToPage];

                    // Convert points to pixels
                    double pixelsLeft = PointsToPixels(Convert.ToDouble(pointsLeft));
                    double pixelsTop = PointsToPixels(Convert.ToDouble(pointsTop));

                    // Calculate width using the font size and text length
                    float fontSize = range.Font.Size;
                    // Approximate width calculation based on font size and character count
                    double approximateWidth = PointsToPixels(fontSize * incorrectWord.Length * 0.6);  // 0.6 is an average character width factor

                    // Get the window position
                    NativeMethods.POINT clientPoint = new NativeMethods.POINT
                    {
                        X = (int)pixelsLeft ,
                        Y = (int)pixelsTop
                    };
                    NativeMethods.ClientToScreen(wordHandle , ref clientPoint);

                    List<string> suggestions = flaggedToken.suggestions.Select(s => s.text).ToList();
                    Point wpfPoint = ScreenToWpf(new Point(clientPoint.X , clientPoint.Y));

                    errors.Add((
                        new System.Windows.Point(wpfPoint.X , wpfPoint.Y + 20),
                        approximateWidth,
                        incorrectWord,
                        suggestions,
                        startIndex,
                        endIndex
                    ));

                    Marshal.ReleaseComObject(range);
                }

                Marshal.ReleaseComObject(wordApp.ActiveDocument);
                Marshal.ReleaseComObject(wordApp);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error in GetWordSpellingErrors: {ex.Message}");
            }

            return errors;
        }

        // Helper method to convert points to pixels
        private double PointsToPixels(double points)
        {
            // 1 point = 1/72 inch
            // Assuming 96 DPI (standard Windows resolution)
            return points * ( 96.0 / 72.0 );
        }

        private List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> GetSpellingErrors(
    IntPtr editHandle ,
    ApiResponse apiResponse,
    string text)
        {
            List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> spellingErrors = new List<(System.Windows.Point, double, string, List<string>, int, int)>();

            IntPtr hdc = NativeMethods.GetDC(editHandle);
            IntPtr hFont = NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETFONT , IntPtr.Zero , IntPtr.Zero);
            IntPtr oldFont = NativeMethods.SelectObject(hdc , hFont);

            foreach (var flaggedToken in apiResponse.spellCheckResponse.results.flagged_tokens)
            {
                int startIndex = flaggedToken.start_index;
                int endIndex = flaggedToken.end_index;
                string incorrectWord = "";
                if (startIndex >= 0 && startIndex < text.Length && endIndex <= text.Length && endIndex > startIndex)
                {
                    incorrectWord = text.Substring(startIndex , endIndex - startIndex);
                }
                else
                {
                    return spellingErrors; 
                }

                // Measure the text width
                NativeMethods.SIZE size;
                NativeMethods.GetTextExtentPoint32(hdc , incorrectWord , incorrectWord.Length , out size);

                // Get the position of the start of the word
                IntPtr charPos = NativeMethods.SendMessage(editHandle , NativeMethods.EM_POSFROMCHAR , (IntPtr)startIndex , IntPtr.Zero);
                int x = ( charPos.ToInt32() & 0xFFFF );
                int y = ( ( charPos.ToInt32() >> 16 ) & 0xFFFF );

                // Adjust for RTL text
                x = x - size.cx; // Move the starting point to the right edge of the word

                NativeMethods.POINT clientPoint = new NativeMethods.POINT { X = x , Y = y };
                NativeMethods.ClientToScreen(editHandle , ref clientPoint);

                List<string> suggestions = flaggedToken.suggestions.Select(s => s.text).ToList();
                Point wpfPoint = ScreenToWpf(new Point(clientPoint.X , clientPoint.Y));
                spellingErrors.Add((new System.Windows.Point(wpfPoint.X , wpfPoint.Y + 20), size.cx, incorrectWord, suggestions, startIndex, endIndex));
            }

            NativeMethods.SelectObject(hdc , oldFont);
            NativeMethods.ReleaseDC(editHandle , hdc);
            return spellingErrors;
        }

        private List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> GetGrammarErrors(
            IntPtr editHandle ,
            ApiResponse apiResponse,
            string text)
        {
            List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> spellingErrors = new List<(System.Windows.Point, double, string, List<string>, int, int)>();

            IntPtr hdc = NativeMethods.GetDC(editHandle);
            IntPtr hFont = NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETFONT , IntPtr.Zero , IntPtr.Zero);
            IntPtr oldFont = NativeMethods.SelectObject(hdc , hFont);
            if (apiResponse.grammarResponse.results.flagged_tokens is null)
                return new();

            foreach (var flaggedToken in apiResponse.grammarResponse.results.flagged_tokens)
            {
                int startIndex = flaggedToken.start_index;
                int endIndex = flaggedToken.end_index;
                string incorrectWord = "";
                if (startIndex >= 0 && startIndex < text.Length && endIndex <= text.Length && endIndex > startIndex)
                {
                    incorrectWord = text.Substring(startIndex , endIndex - startIndex);
                }
                else
                {
                    return spellingErrors;
                }
                // Measure the text width
                NativeMethods.SIZE size;
                NativeMethods.GetTextExtentPoint32(hdc , incorrectWord , incorrectWord.Length , out size);

                // Get the position of the start of the word
                IntPtr charPos = NativeMethods.SendMessage(editHandle , NativeMethods.EM_POSFROMCHAR , (IntPtr)startIndex , IntPtr.Zero);
                int x = ( charPos.ToInt32() & 0xFFFF );
                int y = ( ( charPos.ToInt32() >> 16 ) & 0xFFFF );

                // Adjust for RTL text
                x -= size.cx; // Move the starting point to the right edge of the word

                NativeMethods.POINT clientPoint = new NativeMethods.POINT { X = x , Y = y };
                NativeMethods.ClientToScreen(editHandle , ref clientPoint);

                List<string> suggestions = flaggedToken.suggestions.Select(s => s.text).ToList();
                Point wpfPoint = ScreenToWpf(new Point(clientPoint.X , clientPoint.Y));
                spellingErrors.Add((new System.Windows.Point(wpfPoint.X , wpfPoint.Y + 20), size.cx, incorrectWord, suggestions, startIndex, endIndex));
            }

            NativeMethods.SelectObject(hdc , oldFont);
            NativeMethods.ReleaseDC(editHandle , hdc);
            return spellingErrors;
        }

        private List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> GetPhrasingErrors(
            IntPtr editHandle ,
            ApiResponse apiResponse,
            string text)
        {
            List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> spellingErrors = new List<(System.Windows.Point, double, string, List<string>, int, int)>();

            IntPtr hdc = NativeMethods.GetDC(editHandle);
            IntPtr hFont = NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETFONT , IntPtr.Zero , IntPtr.Zero);
            IntPtr oldFont = NativeMethods.SelectObject(hdc , hFont);
            if (apiResponse.phrasingResponse?.results?.flagged_tokens is null)
                return new();
            foreach (var flaggedToken in apiResponse.phrasingResponse.results.flagged_tokens)
            {
                int startIndex = flaggedToken.start_index;
                int endIndex = flaggedToken.end_index;
                string incorrectWord = "";
                if (startIndex >= 0 && startIndex < text.Length && endIndex <= text.Length && endIndex > startIndex)
                {
                    incorrectWord = text.Substring(startIndex , endIndex - startIndex);
                }
                else
                {
                    return spellingErrors;
                }
                // Measure the text width
                NativeMethods.SIZE size;
                NativeMethods.GetTextExtentPoint32(hdc , incorrectWord , incorrectWord.Length , out size);

                // Get the position of the start of the word
                IntPtr charPos = NativeMethods.SendMessage(editHandle , NativeMethods.EM_POSFROMCHAR , (IntPtr)startIndex , IntPtr.Zero);
                int x = ( charPos.ToInt32() & 0xFFFF );
                int y = ( ( charPos.ToInt32() >> 16 ) & 0xFFFF );

                // Adjust for RTL text
                x -= size.cx; // Move the starting point to the right edge of the word

                NativeMethods.POINT clientPoint = new NativeMethods.POINT { X = x , Y = y };
                NativeMethods.ClientToScreen(editHandle , ref clientPoint);

                List<string> suggestions = flaggedToken.suggestions.Select(s => s.text).ToList();
                Point wpfPoint = ScreenToWpf(new Point(clientPoint.X , clientPoint.Y));
                spellingErrors.Add((new System.Windows.Point(wpfPoint.X , wpfPoint.Y + 20), size.cx, incorrectWord, suggestions, startIndex, endIndex));
            }

            NativeMethods.SelectObject(hdc , oldFont);
            NativeMethods.ReleaseDC(editHandle , hdc);
            return spellingErrors;
        }

        private List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> GetTafqitErrors(
            IntPtr editHandle ,
            ApiResponse apiResponse,
            string text)
        {
            List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> spellingErrors = new List<(System.Windows.Point, double, string, List<string>, int, int)>();

            IntPtr hdc = NativeMethods.GetDC(editHandle);
            IntPtr hFont = NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETFONT , IntPtr.Zero , IntPtr.Zero);
            IntPtr oldFont = NativeMethods.SelectObject(hdc , hFont);
            if (apiResponse.tafqitResponse?.results?.flagged_tokens is null)
                return new();
            foreach (var flaggedToken in apiResponse.tafqitResponse?.results?.flagged_tokens)
            {
                int startIndex = flaggedToken.start_index;
                int endIndex = flaggedToken.end_index;
                string incorrectWord = "";
                if (startIndex >= 0 && startIndex < text.Length && endIndex <= text.Length && endIndex > startIndex)
                {
                    incorrectWord = text.Substring(startIndex , endIndex - startIndex);
                }
                else
                {
                    return spellingErrors;
                }
                // Measure the text width
                NativeMethods.SIZE size;
                NativeMethods.GetTextExtentPoint32(hdc , incorrectWord , incorrectWord.Length , out size);

                // Get the position of the start of the word
                IntPtr charPos = NativeMethods.SendMessage(editHandle , NativeMethods.EM_POSFROMCHAR , (IntPtr)startIndex , IntPtr.Zero);
                int x = ( charPos.ToInt32() & 0xFFFF );
                int y = ( ( charPos.ToInt32() >> 16 ) & 0xFFFF );

                // Adjust for RTL text
                x -= size.cx; // Move the starting point to the right edge of the word

                NativeMethods.POINT clientPoint = new NativeMethods.POINT { X = x , Y = y };
                NativeMethods.ClientToScreen(editHandle , ref clientPoint);

                List<string> suggestions = flaggedToken.suggestions.Select(s => s.text).ToList();
                Point wpfPoint = ScreenToWpf(new Point(clientPoint.X , clientPoint.Y));
                spellingErrors.Add((new System.Windows.Point(wpfPoint.X , wpfPoint.Y + 20), size.cx, incorrectWord, suggestions, startIndex, endIndex));
            }

            NativeMethods.SelectObject(hdc , oldFont);
            NativeMethods.ReleaseDC(editHandle , hdc);
            return spellingErrors;
        }
        private List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> GetTermErrors(
            IntPtr editHandle ,
            ApiResponse apiResponse,
            string text)
        {
            List<(System.Windows.Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> spellingErrors = new List<(System.Windows.Point, double, string, List<string>, int, int)>();

            IntPtr hdc = NativeMethods.GetDC(editHandle);
            IntPtr hFont = NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETFONT , IntPtr.Zero , IntPtr.Zero);
            IntPtr oldFont = NativeMethods.SelectObject(hdc , hFont);
            var termSuggestions = new List<Term>();

            if (apiResponse.termSuggestions is not null && apiResponse.termSuggestions.Any())
            {
                // Create a list to store the deserialized terms

                // Loop through each object and cast/deserialize it to a Term object
                foreach (var suggestion in apiResponse.termSuggestions)
                {
                    // Assuming the suggestion is a JSON string or JObject
                    var term = JsonConvert.DeserializeObject<Term>(suggestion.ToString());
                    termSuggestions.Add(term);
                }
            }

            foreach (var term in termSuggestions)
            {
                int startIndex = term.from;
                int endIndex = term.to - 1;
                if (termSuggestions is null || !termSuggestions.Any())
                    return new();
                string incorrectWord = "";
                if (startIndex >= 0 && startIndex < text.Length && endIndex <= text.Length && endIndex > startIndex)
                {
                    incorrectWord = text.Substring(startIndex , endIndex - startIndex);
                }
                else
                {
                    return spellingErrors;
                }
                // Measure the text width
                NativeMethods.SIZE size;
                NativeMethods.GetTextExtentPoint32(hdc , incorrectWord , incorrectWord.Length , out size);

                // Get the position of the start of the word
                IntPtr charPos = NativeMethods.SendMessage(editHandle , NativeMethods.EM_POSFROMCHAR , (IntPtr)startIndex , IntPtr.Zero);
                int x = ( charPos.ToInt32() & 0xFFFF );
                int y = ( ( charPos.ToInt32() >> 16 ) & 0xFFFF );

                // Adjust for RTL text
                x -= size.cx; // Move the starting point to the right edge of the word

                NativeMethods.POINT clientPoint = new NativeMethods.POINT { X = x , Y = y };
                NativeMethods.ClientToScreen(editHandle , ref clientPoint);

                List<string> suggestions = new List<string>();
                suggestions.Add(term.replacement);
                Point wpfPoint = ScreenToWpf(new Point(clientPoint.X , clientPoint.Y));
                spellingErrors.Add((new System.Windows.Point(wpfPoint.X , wpfPoint.Y + 20), size.cx, incorrectWord, suggestions, startIndex, endIndex));
            }

            NativeMethods.SelectObject(hdc , oldFont);
            NativeMethods.ReleaseDC(editHandle , hdc);
            return spellingErrors;
        } 
        #endregion

        private async Task<ApiResponse> GetSpellCheckResultsAsync(string text)
        {
            _bearerToken = SecureTokenStorage.RetrieveToken();
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(_bearerToken))
                return null;
            var request = new SpellCheckRequest
            {
                text = text.Replace("\n" , "/n") ,
                docId = _documentId
            };

            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json , Encoding.UTF8 , "application/json");
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer" , _bearerToken);
            var response = await _httpClient.PostAsync(_apiUrl , content);
            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
            {
                _authService.LogOut();
                return null;
            }
            response.EnsureSuccessStatusCode();
            var responseString = await response.Content.ReadAsStringAsync();
            var result = JsonConvert.DeserializeObject<ApiResponse>(responseString);
            return result;
        }

        private void WindowCheckTimer_Tick(object sender , EventArgs e)
        {
            IntPtr foregroundWindow = NativeMethods.GetForegroundWindow();
            if (foregroundWindow != _lastActiveWindowHandle)
            {
                _lastActiveWindowHandle = foregroundWindow;
                StringBuilder sb = new StringBuilder(256);
                NativeMethods.GetWindowText(foregroundWindow , sb , 256);
                string newWindowTitle = sb.ToString();

                if (newWindowTitle != _lastActiveWindowTitle)
                {
                    _floatingPoint.Show();
                    _lastActiveWindowTitle = newWindowTitle;
                    if (IsNotepad(newWindowTitle))
                    {
                        _notepadProcessId = GetProcessId(foregroundWindow);
                        ReadWindowContent(foregroundWindow);
                    }
                    else if (IsMicroSoftWord(newWindowTitle))
                    {
                        _microsoftWordProcessId = GetProcessId(foregroundWindow);
                        ReadWordContent();
                    }
                }

                else
                {
                    _floatingPoint.Hide();
                }

                UpdateOverlayPosition();
                CheckForIncorrectWords();

                // Update the current screen
                _currentScreen = Screen.FromHandle(foregroundWindow);
            }
        }
        private void ProcessCheckTimer_Tick(object sender , EventArgs e)
        {
            if (_notepadProcessId.HasValue)
            {
                var process = Process.GetProcesses().FirstOrDefault(p => p.Id == _notepadProcessId.Value);
                if (process == null)
                {
                    _allText.Clear();
                    _notepadProcessId = null;
                }
            }
            else if (_microsoftWordProcessId.HasValue)
            {

                var wordProcess = Process.GetProcesses().FirstOrDefault(p => p.Id == _microsoftWordProcessId.Value);
                if (wordProcess == null)
                {
                    _allText.Clear();
                    _microsoftWordProcessId = null;
                }
            }
        }

        private bool IsMicroSoftWord(string windowTitle)
        {
            return windowTitle.EndsWith(" - Word");
        }

        private void ReadWordContent()
        {
            Application wordApp = null;
            Document activeDocument = null;

            try
            {
                // Attempt to get the existing instance of Word
                try
                {
                    wordApp = (Application)SpellCheckerDemo.ReplaceMarshall.GetActiveObject("Word.Application");
                }
                catch (COMException)
                {
                    Console.WriteLine("Microsoft Word is not running.");
                    return;
                }

                // Access the active Word document
                if (wordApp != null && wordApp.Documents.Count > 0)
                {
                    activeDocument = wordApp.ActiveDocument;
                    StringBuilder sb = new StringBuilder();

                    // Loop through each paragraph in the document
                    foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in activeDocument.Paragraphs)
                    {
                        sb.AppendLine(paragraph.Range.Text);
                    }

                    _allText.Clear();
                    _allText.Append(sb.ToString());
                    _currentWord.Clear();

                    Dispatcher.Invoke(() =>
                    {
                        CheckForIncorrectWords(); // Custom method to highlight specific words
                    });

                    Console.WriteLine("Read content from Word: " + _allText.ToString());
                }
                else
                {
                    Console.WriteLine("No active Word document found.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading Word content: {ex.Message}");
            }
        }

        private string GetMicrosoftWordContent()
        {
            try
            {
                Application wordApp = null;
                Document activeDocument = null;
                try
                {
                    wordApp = (Application)SpellCheckerDemo.ReplaceMarshall.GetActiveObject("Word.Application");
                }
                catch (COMException)
                {
                    Console.WriteLine("Microsoft Word is not running.");
                    return null;
                }

                // Access the active Word document
                if (wordApp != null && wordApp.Documents.Count > 0)
                {
                    activeDocument = wordApp.ActiveDocument;
                    StringBuilder sb = new StringBuilder();

                    // Loop through each paragraph in the document
                    foreach (Microsoft.Office.Interop.Word.Paragraph paragraph in activeDocument.Paragraphs)
                    {
                        sb.AppendLine(paragraph.Range.Text);
                    }
                    return sb.ToString();
                }
            }
            catch (Exception)
            {
                Console.WriteLine("Microsoft Word is not running.");
            }
            return null;

        }

        private void UpdateOverlayPosition()
        {
            NativeMethods.RECT rect;
            if (NativeMethods.GetWindowRect(_lastActiveWindowHandle , out rect))
            {
                // Convert screen coordinates to DPI-aware WPF coordinates
                System.Windows.Point topLeft = ScreenToWpf(new System.Windows.Point(rect.Left , rect.Top));
                System.Windows.Point bottomRight = ScreenToWpf(new System.Windows.Point(rect.Right , rect.Bottom));

                _overlay.Left = topLeft.X;
                _overlay.Top = topLeft.Y;
                _overlay.Width = bottomRight.X - topLeft.X;
                _overlay.Height = bottomRight.Y - topLeft.Y;
            }
        }

        private System.Windows.Point ScreenToWpf(System.Windows.Point screenPoint)
        {
            PresentationSource source = PresentationSource.FromVisual(this);
            if (source == null) return screenPoint;

            return source.CompositionTarget.TransformFromDevice.Transform(screenPoint);
        }

        private int? GetProcessId(IntPtr windowHandle)
        {
            try
            {
                Process process = Process.GetProcesses().FirstOrDefault(p => p.MainWindowHandle == windowHandle);
                return process?.Id;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error retrieving process ID: {ex.Message}");
                return null;
            }
        }

        private bool IsNotepad(string windowTitle)
        {
            return windowTitle.EndsWith(".txt - Notepad") ||
                   windowTitle.Equals("Untitled - Notepad") ||
                   windowTitle.EndsWith("Notepad");
        }

        private void ReadWindowContent(IntPtr notepadHandle)
        {
            try
            {
                IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);
                if (editHandle != IntPtr.Zero)
                {
                    int length = (int)NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETTEXTLENGTH , IntPtr.Zero , null);
                    if (length > 0)
                    {
                        StringBuilder sb = new StringBuilder(length + 1);
                        NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETTEXT , (IntPtr)sb.Capacity , sb);

                        _allText.Clear();
                        _allText.Append(sb.ToString());
                        _currentWord.Clear();

                        Dispatcher.Invoke(() =>
                        {
                            CheckForIncorrectWords();
                        });

                        Console.WriteLine("Read content from Notepad: " + _allText.ToString());
                    }
                }
                else
                {
                    Console.WriteLine("Failed to find the edit control.");
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error reading Notepad content: {ex.Message}");
            }
        }

        private void OnKeyPressed(object sender , Keys e)
        {
            _lastActiveWindowHandle = NativeMethods.GetForegroundWindow();

            char keyChar = GetCharFromKey(e);

            if (char.IsLetter(keyChar))
            {
                _currentWord.Append(keyChar.ToString());
                _allText.Append(keyChar.ToString());
            }
            else if (e == Keys.Space || e == Keys.OemPeriod)
            {
                _currentWord.Clear();
                _allText.Append((char)e);
            }
            else if (e == Keys.Back && _currentWord.Length > 0)
            {
                _currentWord.Length--;
                _allText.Length--;
            }

            Dispatcher.Invoke(() =>
            {
                CheckForIncorrectWords();
            });
        }

        private char GetCharFromKey(Keys key)
        {
            byte[] keyboardState = new byte[256];
            NativeMethods.GetKeyboardState(keyboardState);

            uint scanCode = NativeMethods.MapVirtualKey((uint)key , 0);
            StringBuilder stringBuilder = new StringBuilder(2);

            int result = NativeMethods.ToUnicode((uint)key , scanCode , keyboardState , stringBuilder , stringBuilder.Capacity , 0);
            if (result > 0)
            {
                return stringBuilder[0];
            }
            else
            {
                return '\0';
            }
        }

        private void UpdateTextAndClosePopup(string word)
        {
            _currentWord.Clear();

            Dispatcher.Invoke(() =>
            {
                CheckForIncorrectWords();
                _suggestionsControl.SetSuggestion(null);
                SuggestionsPopup.IsOpen = false;
            });
        }

        private void UpdateErrorCount(int errorCount)
        {
            _floatingPoint.UpdateErrorCount(errorCount);
        }

        private void SendKeys(Keys key)
        {
            NativeMethods.keybd_event((byte)key , 0 , 0 , UIntPtr.Zero);
            NativeMethods.keybd_event((byte)key , 0 , NativeMethods.KEYEVENTF_KEYUP , UIntPtr.Zero);
        }

        private void ContentSyncTimer_Tick(object sender , EventArgs e)
        {
            if (_notepadProcessId.HasValue)
            {
                string notepadContent = GetNotepadContent();
                if (notepadContent != _allText.ToString())
                {
                    _allText.Clear();
                    _allText.Append(notepadContent);
                    _currentWord.Clear();
                    Dispatcher.Invoke(() =>
                    {
                        CheckForIncorrectWords();
                    });
                }
            }
            else if (_microsoftWordProcessId.HasValue)
            {
                string wordContent = GetMicrosoftWordContent();
                if (wordContent != _allText.ToString())
                {
                    _allText.Clear();
                    _allText.Append(wordContent);
                    _currentWord.Clear();
                    Dispatcher.Invoke(() =>
                    {
                        CheckForIncorrectWords();
                    });
                }
            }
        }

        private string GetNotepadContent(IntPtr? notepadHandle = null)
        {
            if (notepadHandle == null)
            {
                notepadHandle = NativeMethods.GetForegroundWindow();
            }
            IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle.Value , IntPtr.Zero , "Edit" , null);
            if (editHandle != IntPtr.Zero)
            {
                int length = (int)NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETTEXTLENGTH , IntPtr.Zero , null);
                if (length > 0)
                {
                    StringBuilder sb = new StringBuilder(length + 1);
                    NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETTEXT , (IntPtr)sb.Capacity , sb);
                    return sb.ToString();
                }
            }
            return string.Empty;
        }
        private void CheckKeyboardLayout()
        {
            IntPtr layout = NativeMethods.GetKeyboardLayout(0); // 0 means to get the layout of the current thread.
            int languageCode = layout.ToInt32() & 0xFFFF; // The low word contains the language identifier.
            if (languageCode == 0x0C01) // 0x0C01 is the hexadecimal code for Arabic (Saudi Arabia).
            {
                Console.WriteLine("Arabic keyboard layout is active.");
            }
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            base.OnClosing(e);
            _floatingPoint.Close();
            _overlay.Close();
            _keyboardHook.Dispose();
        }
    }
}
