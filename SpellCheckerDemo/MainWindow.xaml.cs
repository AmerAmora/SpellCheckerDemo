using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using Keys = System.Windows.Forms.Keys;
using System.Windows.Threading;
using SpellCheckerDemo;
using Brushes = System.Windows.Media.Brushes;
using Point = System.Drawing.Point;
using System.Drawing;
using System.Windows.Interop;
using System.Windows.Media;
using static NativeMethods;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using SpellCheckerDemo.Models;

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
        private readonly string _bearerToken = "eyJraWQiOiJCSHhSWWpqenV6N1JpKzM4dVlCWkJcLzYwR3FIcVhqQjI2bHAxOVd6dTIwaz0iLCJhbGciOiJSUzI1NiJ9.eyJzdWIiOiI2MmU3ODA0Yi03N2E0LTRjMTQtOGEwYi1iMWJmNjhkODhmZWEiLCJjb2duaXRvOmdyb3VwcyI6WyJhZG1pbiJdLCJlbWFpbF92ZXJpZmllZCI6dHJ1ZSwiY3VzdG9tOnV0bV9zcmMiOiJOQSIsImlzcyI6Imh0dHBzOlwvXC9jb2duaXRvLWlkcC5ldS13ZXN0LTEuYW1hem9uYXdzLmNvbVwvZXUtd2VzdC0xX2xNc0lHNmQ3ZyIsImNvZ25pdG86dXNlcm5hbWUiOiI2MmU3ODA0Yi03N2E0LTRjMTQtOGEwYi1iMWJmNjhkODhmZWEiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiI4NDZmMzAzNi1lNjg2LTRlNDMtYjUyYy00NmFmYTE4OTM4YTciLCJnaXZlbl9uYW1lIjoidGVzdCIsImF1ZCI6IjU5cW0xbDRnamlnczc2bzVqOTJucDQ2MGp0IiwiZXZlbnRfaWQiOiI4YzA1ODEyOC05ZDIxLTQzY2ItYWM3Ny0xMTAwZjEyMWY1NDkiLCJ0b2tlbl91c2UiOiJpZCIsImF1dGhfdGltZSI6MTcyNjc1MDc3OCwiZXhwIjoxNzI2NzU0Mzc4LCJpYXQiOjE3MjY3NTA3NzgsImZhbWlseV9uYW1lIjoidGVzdCIsImVtYWlsIjoicGVkaWZhYzI4M0BjZXRub2IuY29tIn0.XGll6gXeZyc4riSoF26qFQXgSndWjTbRdlCYr_NpdzSnWJzYnEweny4tuAcK2ya8mH55jNUa1h3ThN1F9-LnKhbID45vhzB3Y0IyxiAXIZMS5iRZSeJRzlZBpTiMAlGsbN0XZTq7guY11UjtBc-p3BuexAVqVYmDFrW80V24kSEJJ1g3-6_RKVC1v3CNOUT8tx_W6hTW0_LDKXpLQ5iIuNhxPqJ3FQKiMH5o73wUouhSZv51eOzQu9D33ywjmkhYCakUsPfjQixDIFn0MapEm6yi4AEEG3DxySZPhE14ihXedH4XCQua2hhB2LZVsHJm7_ygiNv9_dsvrpB9mf-xNQ";
        private string _documentId;

        public MainWindow()
        {
            InitializeComponent();
            _keyboardHook = new GlobalKeyboardHook();
            _keyboardHook.KeyPressed += OnKeyPressed;

            // Initialize SuggestionsControl
            _suggestionsControl = new SuggestionsControl();
            _suggestionsControl.SuggestionAccepted += OnSuggestionAccepted;
            _suggestionsControl.SuggestionCancelled += OnSuggestionCancelled;
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
            _floatingPoint.Show();

            _overlay = new OverlayWindow();
            _overlay.Show();

            this.Hide();
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer" , _bearerToken);
            _documentId = Guid.NewGuid().ToString(); // Generate a unique document ID
        }

        private async Task<ApiResponse> GetSpellCheckResultsAsync(string text)
        {
            if (string.IsNullOrEmpty(text))
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
                    _lastActiveWindowTitle = newWindowTitle;
                    if (IsNotepad(newWindowTitle))
                    {
                        _notepadProcessId = GetProcessId(foregroundWindow);
                        ReadWindowContent(foregroundWindow);
                    }
                    else
                    {
                        _notepadProcessId = null;
                    }
                }

                UpdateOverlayPosition();
                CheckForIncorrectWords();
            }
        }

        private void UpdateOverlayPosition()
        {
            NativeMethods.RECT rect;
            if (NativeMethods.GetWindowRect(_lastActiveWindowHandle , out rect))
            {
                _overlay.Left = rect.Left;
                _overlay.Top = rect.Top;
                _overlay.Width = rect.Right - rect.Left;
                _overlay.Height = rect.Bottom - rect.Top;
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
                MessageBox.Show($"Error retrieving process ID: {ex.Message}");
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
                            HighlightTeh();
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
                MessageBox.Show($"Error reading Notepad content: {ex.Message}");
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
                CheckWord();
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

        private void CheckWord()
        {
            string word = _currentWord.ToString();
            if (word.Equals("teh" , StringComparison.OrdinalIgnoreCase))
            {
                Dispatcher.Invoke(() =>
                {
                    _suggestionsControl.SetSuggestion("The");
                    SuggestionsPopup.IsOpen = true;
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    _suggestionsControl.SetSuggestion(null);
                    SuggestionsPopup.IsOpen = false;
                });
            }
        }

        private void OnSuggestionAccepted(object sender , EventArgs e)
        {
            ReplaceWordInActiveApplication("The");
            UpdateTextAndClosePopup("The");
        }

        private void OnSuggestionCancelled(object sender , EventArgs e)
        {
            UpdateTextAndClosePopup("teh");
        }

        private void UpdateTextAndClosePopup(string word)
        {
            ReplaceWordInActiveApplication(word);

            _currentWord.Clear();

            Dispatcher.Invoke(() =>
            {
                HighlightTeh();
                CheckForIncorrectWords();
                _suggestionsControl.SetSuggestion(null);
                SuggestionsPopup.IsOpen = false;
            });
        }

        private void HighlightTeh()
        {
            string fullText = _allText.ToString();
            CurrentWordTextBlock.Inlines.Clear();

            string[] parts = fullText.Split(new[] { "teh" } , StringSplitOptions.None);

            for (int i = 0 ; i < parts.Length ; i++)
            {
                CurrentWordTextBlock.Inlines.Add(new Run(parts[i]));

                if (i < parts.Length - 1)
                {
                    var incorrectWord = new Run("teh")
                    {
                        Background = Brushes.Red ,
                        Foreground = Brushes.White
                    };
                    CurrentWordTextBlock.Inlines.Add(incorrectWord);
                }
            }
        }

        private async void CheckForIncorrectWords()
        {
            string text = _allText.ToString();
            var apiResponse = await GetSpellCheckResultsAsync(text);

            if (apiResponse?.spellCheckResponse?.results?.flagged_tokens == null)
            {
                Console.WriteLine("No flagged tokens found or API response was null.");
                return;
            }

            List<(System.Windows.Point screenPosition, double width, string incorrectWord, string suggestion)> underlines = new List<(System.Windows.Point, double, string, string)>();

            foreach (var flaggedToken in apiResponse.spellCheckResponse.results.flagged_tokens)
            {
                IntPtr notepadHandle = NativeMethods.GetForegroundWindow();
                IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);

                if (editHandle != IntPtr.Zero)
                {
                    IntPtr charPos = NativeMethods.SendMessage(editHandle , NativeMethods.EM_POSFROMCHAR , (IntPtr)flaggedToken.start_index , IntPtr.Zero);
                    int x = ( charPos.ToInt32() & 0xFFFF );
                    int y = ( ( charPos.ToInt32() >> 16 ) & 0xFFFF );

                    POINT clientPoint = new POINT { X = x , Y = y };
                    NativeMethods.ClientToScreen(editHandle , ref clientPoint);

                    double width = ( flaggedToken.end_index - flaggedToken.start_index ) * 8; // Approximate width based on character count
                    string incorrectWord = text.Substring(flaggedToken.start_index , flaggedToken.end_index - flaggedToken.start_index);
                    string suggestion = flaggedToken.suggestions.FirstOrDefault().text ?? "";
                    underlines.Add((new System.Windows.Point(clientPoint.X , clientPoint.Y + 20), width, incorrectWord, suggestion));
                }
            }

            Application.Current.Dispatcher.Invoke(() =>
            {
                if (underlines.Count > 0)
                {
                    _overlay.DrawUnderlines(underlines);
                    Console.WriteLine($"Drew {underlines.Count} underlines.");
                }
                else
                {
                    _overlay.DrawUnderlines(new List<(System.Windows.Point, double, string, string)>()); // Clear all underlines
                    Console.WriteLine("No errors found in text");
                }
                UpdateErrorCount(apiResponse.spellCheckResponse.results.flagged_tokens.Count);
            });
        }

        private void InitializeOverlay()
        {
            _overlay = new OverlayWindow();
            _overlay.Show();
            _overlay.SuggestionAccepted += Overlay_SuggestionAccepted;
        }

        private void Overlay_SuggestionAccepted(object sender , string suggestion)
        {
            ReplaceWordInActiveApplication(suggestion);
            CheckForIncorrectWords(); // Recheck for incorrect words after replacement
        }

        private void UpdateErrorCount(int errorCount)
        {
            _floatingPoint.UpdateErrorCount(errorCount);
        }

        private Point GetPositionOfWord(string word)
        {
            IntPtr notepadHandle = NativeMethods.GetForegroundWindow();
            IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);
            if (editHandle != IntPtr.Zero)
            {
                string notepadText = GetNotepadContent();
                int index = notepadText.IndexOf(word , StringComparison.OrdinalIgnoreCase);
                if (index != -1)
                {
                    // Use SendMessage to get the position of the word
                    IntPtr pos = NativeMethods.SendMessage(editHandle , NativeMethods.EM_POSFROMCHAR , (IntPtr)index , IntPtr.Zero);
                    if (pos != IntPtr.Zero)
                    {
                        int x = pos.ToInt32() & 0xFFFF; // X position is in low-order word
                        int y = ( pos.ToInt32() >> 16 ) & 0xFFFF; // Y position is in high-order word
                        return new Point(x , y);
                    }
                }
            }
            return Point.Empty;
        }

        private void ReplaceWordInActiveApplication(string replacementWord)
        {
            IntPtr hWnd = NativeMethods.GetForegroundWindow();
            if (hWnd == IntPtr.Zero)
                return;

            NativeMethods.SetForegroundWindow(hWnd);

            // Get the current content
            string currentContent = _allText.ToString();

            // Find the last occurrence of "teh"
            int lastIndex = currentContent.LastIndexOf("teh" , StringComparison.OrdinalIgnoreCase);
            if (lastIndex == -1)
                return; // "teh" not found, no replacement needed

            // Calculate how many backspaces are needed
            int backspaces = currentContent.Length - lastIndex - 3;

            // Send backspace keys to remove "teh"
            for (int i = 0 ; i < backspaces + 3 ; i++)
            {
                SendKeys(Keys.Back);
            }

            // Send the replacement word
            foreach (char c in replacementWord)
            {
                SendKeys((Keys)char.ToUpper(c));
            }

            // Update our internal text representation
            _allText.Remove(lastIndex , 3);
            _allText.Insert(lastIndex , replacementWord);
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
                        HighlightTeh();
                        CheckForIncorrectWords();
                    });
                }
            }
        }

        private string GetNotepadContent()
        {
            IntPtr notepadHandle = NativeMethods.GetForegroundWindow();
            IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);
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

    public class SpellCheckRequest
    {
        public string text { get; set; }
        public string docId { get; set; }
    }

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

    public class SpellCheckResponse
    {
        public string version { get; set; }
        public SpellCheckResults results { get; set; }
    }

    public class SpellCheckResults
    {
        public List<FlaggedToken> flagged_tokens { get; set; }
    }

    public class FlaggedToken
    {
        public string original_word { get; set; }
        public int start_index { get; set; }
        public int end_index { get; set; }
        public List<Suggestion> suggestions { get; set; }
        public string arabic_reason { get; set; }
        public string lang { get; set; }
    }

    public class Suggestion
    {
        public string text { get; set; }
        public string confidence { get; set; }
        public List<string> reasons { get; set; }
    }

    public class GrammarResponse
    {
        public string version { get; set; }
        public object results { get; set; }
    }

    public class PhrasingResponse
    {
        public string version { get; set; }
        public object results { get; set; }
    }

    public class TafqitResponse
    {
        public object results { get; set; }
    }

    public class OtherSuggestions
    {
        public string version { get; set; }
        public object results { get; set; }
    }
}
