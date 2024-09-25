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
using System.Windows.Interop;

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
        private readonly string _bearerToken = "eyJraWQiOiJCSHhSWWpqenV6N1JpKzM4dVlCWkJcLzYwR3FIcVhqQjI2bHAxOVd6dTIwaz0iLCJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJiY2RjNmEwMi0wYzEyLTQzNjItYjcxZS04MjY4MzkyYTI5YWUiLCJjb2duaXRvOmdyb3VwcyI6WyJhZG1pbiJdLCJlbWFpbF92ZXJpZmllZCI6dHJ1ZSwiY3VzdG9tOnV0bV9zcmMiOiJOQSIsImlzcyI6Imh0dHBzOlwvXC9jb2duaXRvLWlkcC5ldS13ZXN0LTEuYW1hem9uYXdzLmNvbVwvZXUtd2VzdC0xX2xNc0lHNmQ3ZyIsImNvZ25pdG86dXNlcm5hbWUiOiJiY2RjNmEwMi0wYzEyLTQzNjItYjcxZS04MjY4MzkyYTI5YWUiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiJhOTBmZGQwOC01MWE3LTRmY2MtOWI5NC02MzIxZTI5Y2FiZjYiLCJnaXZlbl9uYW1lIjoiV29yZC1QYWNrYWdlIiwiYXVkIjoiNTlxbTFsNGdqaWdzNzZvNWo5Mm5wNDYwanQiLCJldmVudF9pZCI6ImFhODQ5ZmU4LWE5ZjEtNDMwNS04ZGY2LWY3OTZjMWEyNzg0ZSIsInRva2VuX3VzZSI6ImlkIiwiYXV0aF90aW1lIjoxNzI3MjQ3NTc1LCJleHAiOjE3MjcyNTExNzUsImlhdCI6MTcyNzI0NzU3NSwiZmFtaWx5X25hbWUiOiJVc2VyIiwiZW1haWwiOiJ3aXJlamVmOTAwQGhld2Vlay5jb20ifQ.jBIt4OpUbznF8fk4OWR_a3NCm6PvVjC2scBozLNwchrgnXJtpbgJRcMjTeZ-m5kCecULVKwkqwK9D982wRDxxlQtTl7Du2Xr2GUBujaQNnPRcrpisyosERAr_SpgPKXTf_QDLexqtvUTMCi8TAcHMknSuUxcaskgUgMjtL8aNJePXM7hxivBTdHqx9bbese42PypPP07dRwRhPtCkLdHQDp1-zMRnEXhRofyMSuoW4JG8Mdxi8nNqQ_xRZzx63xj_rueLq9pbky5Uy_A3mNVzvPfaPUiwYIwVTerHmueD3QiMRqOf-_m8rvL3torhrm616LymLbYjgXF5B774rDyew";
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
            this.Loaded += MainWindow_Loaded;

            this.Hide();
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer" , _bearerToken);
            _documentId = Guid.NewGuid().ToString(); // Generate a unique document ID
        }

        private void MainWindow_Loaded(object sender , RoutedEventArgs e)
        {
            HwndSource.FromHwnd(new WindowInteropHelper(this).Handle).AddHook(new HwndSourceHook(WndProc));
        }

        private IntPtr WndProc(IntPtr hwnd , int msg , IntPtr wParam , IntPtr lParam , ref bool handled)
        {
            const int WM_ACTIVATEAPP = 0x001C;

            if (msg == WM_ACTIVATEAPP)
            {
                if (wParam.ToInt32() == 0) // Another window is being activated
                {
                    _overlay.CloseSuggestionPopup();
                }
            }

            return IntPtr.Zero;
        }

        private void Overlay_SuggestionAccepted(object sender , (string suggestion, int startIndex, int endIndex) e)
        {
            ReplaceWord(e.suggestion , e.startIndex , e.endIndex);
            CheckForIncorrectWords(); // Recheck for incorrect words after replacement
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

        public async Task CheckForIncorrectWords()
        {
            string text = _allText.ToString();
            var apiResponse = await GetSpellCheckResultsAsync(text);

            if (apiResponse?.spellCheckResponse?.results?.flagged_tokens == null)
            {
                UpdateErrorCount(0);
                _overlay.DrawUnderlines(new ErrorsUnderlines());
                Console.WriteLine("No flagged tokens found or API response was null.");
                return;
            }

            IntPtr notepadHandle = NativeMethods.GetForegroundWindow();
            IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);

            if (editHandle == IntPtr.Zero)
            {
                Console.WriteLine("Failed to find Notepad edit control.");
                return;
            }

            var errors = new ErrorsUnderlines
            {
                SpellingErrors = GetErrors(editHandle , apiResponse.spellCheckResponse.results.flagged_tokens) ,
                GrammarErrors = GetErrors(editHandle , apiResponse.grammarResponse?.results?.flagged_tokens) ,
                PhrasingErrors = GetErrors(editHandle , apiResponse.phrasingResponse?.results?.flagged_tokens) ,
                TafqitErrors = GetErrors(editHandle , apiResponse.tafqitResponse?.results?.flagged_tokens) ,
                TermErrors = GetTermErrors(editHandle , apiResponse.termSuggestions)
            };

            UpdateUI(errors);
        }

        private List<(Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> GetErrors(IntPtr editHandle , IEnumerable<FlaggedToken> flaggedTokens)
        {
            if (flaggedTokens == null) return new List<(Point, double, string, List<string>, int, int)>();

            return flaggedTokens.Select(token => CreateErrorInfo(editHandle , token)).ToList();
        }

        private List<(Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex)> GetTermErrors(IntPtr editHandle , IEnumerable<object> termSuggestions)
        {
            if (termSuggestions == null) return new List<(Point, double, string, List<string>, int, int)>();

            var terms = termSuggestions.Select(s => JsonConvert.DeserializeObject<Term>(s.ToString()));
            return terms.Select(term => CreateErrorInfo(editHandle , term)).ToList();
        }

        private (Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex) CreateErrorInfo(IntPtr editHandle , FlaggedToken token)
        {
            string incorrectWord = GetIncorrectWord(token.start_index , token.end_index);
            var (screenPosition, width) = GetWordPositionAndWidth(editHandle , incorrectWord , token.start_index);
            return (screenPosition, width, incorrectWord, token.suggestions.Select(s => s.text).ToList(), token.start_index, token.end_index);
        }

        private (Point screenPosition, double width, string incorrectWord, List<string> suggestions, int startIndex, int endIndex) CreateErrorInfo(IntPtr editHandle , Term term)
        {
            string incorrectWord = GetIncorrectWord(term.from , term.to - 1);
            var (screenPosition, width) = GetWordPositionAndWidth(editHandle , incorrectWord , term.from);
            return (screenPosition, width, incorrectWord, new List<string> { term.replacement }, term.from, term.to - 1);
        }

        private string GetIncorrectWord(int startIndex , int endIndex)
        {
            string text = _allText.ToString();
            return ( startIndex >= 0 && startIndex < text.Length && endIndex <= text.Length && endIndex > startIndex )
                ? text.Substring(startIndex , endIndex - startIndex)
                : string.Empty;
        }

        private (Point screenPosition, double width) GetWordPositionAndWidth(IntPtr editHandle , string word , int startIndex)
        {
            using var dc = new SafeDC(editHandle);
            using var font = new SafeFont(editHandle);

            NativeMethods.SIZE size;
            NativeMethods.GetTextExtentPoint32(dc.HDC , word , word.Length , out size);

            IntPtr charPos = NativeMethods.SendMessage(editHandle , NativeMethods.EM_POSFROMCHAR , (IntPtr)startIndex , IntPtr.Zero);
            int x = ( charPos.ToInt32() & 0xFFFF ) - size.cx;
            int y = ( ( charPos.ToInt32() >> 16 ) & 0xFFFF );

            NativeMethods.POINT clientPoint = new NativeMethods.POINT { X = x , Y = y };
            NativeMethods.ClientToScreen(editHandle , ref clientPoint);

            Point wpfPoint = ScreenToWpf(new Point(clientPoint.X , clientPoint.Y));
            return (new Point(wpfPoint.X , wpfPoint.Y + 20), size.cx);
        }

        private void UpdateUI(ErrorsUnderlines errors)
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                int count = errors.TotalErrorCount;
                _overlay.DrawUnderlines(errors);
                UpdateErrorCount(count);
            });
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
                    _floatingPoint.Show();
                    _lastActiveWindowTitle = newWindowTitle;
                    if (IsNotepad(newWindowTitle))
                    {
                        _notepadProcessId = GetProcessId(foregroundWindow);
                        ReadWindowContent(foregroundWindow);
                    }

                    else
                    {
                        _floatingPoint.Hide();
                    }
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
}
