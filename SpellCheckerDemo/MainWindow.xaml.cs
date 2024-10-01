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
using MessageBox = System.Windows.MessageBox;
using static System.Net.Mime.MediaTypeNames;

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

            _floatingPoint = new FloatingPointWindow();
            _floatingPoint.Hide();

            _overlay = new OverlayWindow();
            _overlay.Show();
            _overlay.SuggestionAccepted += Overlay_SuggestionAccepted;

            this.Hide();
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer" , _bearerToken);
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
            string text = GetNotepadContent(notepadHandle);
            var apiResponse = await GetSpellCheckResultsAsync(text);

            ErrorsUnderlines errors = new ErrorsUnderlines();

            if (editHandle != IntPtr.Zero)
            {
                errors.SpellingErrors = GetSpellingErrors(editHandle , apiResponse, text);
                errors.GrammarError = GetGrammarErrors(editHandle , apiResponse, text);
                errors.PhrasingErrors = GetPhrasingErrors(editHandle , apiResponse, text);
                errors.TafqitErrors = GetTafqitErrors(editHandle , apiResponse, text);
                errors.TermErrors = GetTermErrors(editHandle , apiResponse, text);
            }
            foreach (var error in errors.SpellingErrors)
            {
                ReplaceWord(error.suggestions.FirstOrDefault() , error.startIndex , error.endIndex);
            }
            foreach (var error in errors.GrammarError)
            {
                ReplaceWord(error.suggestions.FirstOrDefault() , error.startIndex , error.endIndex);
            }
            foreach (var error in errors.PhrasingErrors)
            {
                ReplaceWord(error.suggestions.FirstOrDefault() , error.startIndex , error.endIndex);
            }
            foreach (var error in errors.TafqitErrors)
            {
                ReplaceWord(error.suggestions.FirstOrDefault() , error.startIndex , error.endIndex);
            }
            foreach (var error in errors.TermErrors)
            {
                ReplaceWord(error.suggestions.FirstOrDefault() , error.startIndex , error.endIndex);
            }
            CheckForIncorrectWords();
        }

        private void ReplaceWord(string suggestion , int startIndex , int endIndex)
        {
            if (_notepadProcessId == null)
            {
                System.Windows.MessageBox.Show("Notepad process ID is not set.");
                return;
            }

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

            // Bring Notepad to the foreground
            NativeMethods.SetForegroundWindow(notepadHandle);

            // Select the incorrect word
            NativeMethods.SendMessage(editHandle , NativeMethods.EM_SETSEL , (IntPtr)startIndex , (IntPtr)endIndex+1);

            // Replace the selection with the suggestion
            foreach (char c in suggestion)
            {
                NativeMethods.SendMessage(editHandle , NativeMethods.WM_CHAR , (IntPtr)c , IntPtr.Zero);
            }

            // Update our internal text representation
            //_allText.Remove(startIndex , endIndex - startIndex);
            //_allText.Insert(startIndex , suggestion);
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
            string text = _allText.ToString();
            var apiResponse = await GetSpellCheckResultsAsync(text);

            if (apiResponse?.spellCheckResponse?.results?.flagged_tokens == null)
            {
                UpdateErrorCount(0);
                _overlay.DrawUnderlines(new ErrorsUnderlines()); // Clear all underlines
                Console.WriteLine("No flagged tokens found or API response was null.");
                return;
            }
            ErrorsUnderlines errors = new ErrorsUnderlines();

            IntPtr notepadHandle = NativeMethods.GetForegroundWindow();
            IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);

            if (editHandle != IntPtr.Zero)
            {
                errors.SpellingErrors = GetSpellingErrors(editHandle,apiResponse,text);
                errors.GrammarError = GetGrammarErrors(editHandle,apiResponse, text);
                errors.PhrasingErrors = GetPhrasingErrors(editHandle, apiResponse, text);
                errors.TafqitErrors = GetTafqitErrors(editHandle,apiResponse, text);
                errors.TermErrors = GetTermErrors(editHandle,apiResponse, text);
            }

            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                var count = errors.SpellingErrors.Count + errors.GrammarError.Count + errors.PhrasingErrors.Count + errors.TafqitErrors.Count + errors.TermErrors.Count;
                _overlay.DrawUnderlines(errors);
                UpdateErrorCount(count);
            });
        }

        #region GetErrors
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
            if (string.IsNullOrEmpty(text))
                return null;
            if (string.IsNullOrEmpty(_bearerToken))
                return null;
            var request = new SpellCheckRequest
            {
                text = text.Replace("\n" , "/n") ,
                docId = _documentId
            };

            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json , Encoding.UTF8 , "application/json");

            var response = await _httpClient.PostAsync(_apiUrl , content);
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
