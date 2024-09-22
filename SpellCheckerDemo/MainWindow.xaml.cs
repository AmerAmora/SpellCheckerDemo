﻿using System.Diagnostics;
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
        private readonly string _bearerToken = "eyJraWQiOiJCSHhSWWpqenV6N1JpKzM4dVlCWkJcLzYwR3FIcVhqQjI2bHAxOVd6dTIwaz0iLCJhbGciOiJSUzI1NiJ9.eyJzdWIiOiI2MmU3ODA0Yi03N2E0LTRjMTQtOGEwYi1iMWJmNjhkODhmZWEiLCJjb2duaXRvOmdyb3VwcyI6WyJhZG1pbiJdLCJlbWFpbF92ZXJpZmllZCI6dHJ1ZSwiY3VzdG9tOnV0bV9zcmMiOiJOQSIsImlzcyI6Imh0dHBzOlwvXC9jb2duaXRvLWlkcC5ldS13ZXN0LTEuYW1hem9uYXdzLmNvbVwvZXUtd2VzdC0xX2xNc0lHNmQ3ZyIsImNvZ25pdG86dXNlcm5hbWUiOiI2MmU3ODA0Yi03N2E0LTRjMTQtOGEwYi1iMWJmNjhkODhmZWEiLCJwcmVmZXJyZWRfdXNlcm5hbWUiOiI4NDZmMzAzNi1lNjg2LTRlNDMtYjUyYy00NmFmYTE4OTM4YTciLCJnaXZlbl9uYW1lIjoidGVzdCIsImF1ZCI6IjU5cW0xbDRnamlnczc2bzVqOTJucDQ2MGp0IiwiZXZlbnRfaWQiOiI1OTdkYTg1My1hYWVlLTQ4ODgtODI4Ni1hMDRlMjY5OTExYTEiLCJ0b2tlbl91c2UiOiJpZCIsImF1dGhfdGltZSI6MTcyNzAwMzczOSwiZXhwIjoxNzI3MDA3MzM5LCJpYXQiOjE3MjcwMDM3MzksImZhbWlseV9uYW1lIjoidGVzdCIsImVtYWlsIjoicGVkaWZhYzI4M0BjZXRub2IuY29tIn0.QkKgRlMZWEVgX4OyXE5iSB3vGCCL9MKOAEJJL_MVOGmW5P4dEstE7yAwx2XnP5IVkE006UNnD3i_OsVRwZhbii2ekZ23Qs6lIJRmIfCjyBf_9WlWk9TGc3Kew2N9VTLvK1StkSjxr2DwGare1T2dPRlYL-4lua1EWEgHIZBSxwkDe7d2JJkILmxo8gTzSpgwTpQBU_D1sUKzsQsBU-wcV670HDnAW7ECz5c2T8J0n9kDYOwnetR2PnudnBveIWffzk3VmvDFYO_SvDn_b3z-nqI_RkcMsgZubfo5M0BJvhQwRD4Rb-wmrkWaCy7xGJbMBgr9DZb0fM_m5WFY3-hXiQ";
        private string _documentId;

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
            _floatingPoint.Show();

            _overlay = new OverlayWindow();
            _overlay.Show();
            _overlay.SuggestionAccepted += Overlay_SuggestionAccepted;

            this.Hide();
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("*/*"));
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer" , _bearerToken);
            _documentId = Guid.NewGuid().ToString(); // Generate a unique document ID
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
                MessageBox.Show("Notepad process ID is not set.");
                return;
            }

            IntPtr notepadHandle = FindNotepadWindow();
            if (notepadHandle == IntPtr.Zero)
            {
                MessageBox.Show("Failed to find Notepad window.");
                return;
            }

            IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);
            if (editHandle == IntPtr.Zero)
            {
                MessageBox.Show("Failed to find Edit control in Notepad.");
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
                Console.WriteLine("No flagged tokens found or API response was null.");
                return;
            }

            List<(System.Windows.Point screenPosition, double width, string incorrectWord, string suggestion, int startIndex, int endIndex)> underlines = new List<(System.Windows.Point, double, string, string, int, int)>();

            IntPtr notepadHandle = NativeMethods.GetForegroundWindow();
            IntPtr editHandle = NativeMethods.FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);

            if (editHandle != IntPtr.Zero)
            {
                IntPtr hdc = NativeMethods.GetDC(editHandle);
                IntPtr hFont = NativeMethods.SendMessage(editHandle , NativeMethods.WM_GETFONT , IntPtr.Zero , IntPtr.Zero);
                IntPtr oldFont = NativeMethods.SelectObject(hdc , hFont);

                foreach (var flaggedToken in apiResponse.spellCheckResponse.results.flagged_tokens)
                {
                    int startIndex = flaggedToken.start_index;
                    int endIndex = flaggedToken.end_index;
                    string incorrectWord = text.Substring(startIndex , endIndex - startIndex);

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

                    string suggestion = flaggedToken.suggestions.FirstOrDefault()?.text ?? "";
                    underlines.Add((new System.Windows.Point(clientPoint.X , clientPoint.Y + 20), size.cx, incorrectWord, suggestion, startIndex, endIndex));
                }

                NativeMethods.SelectObject(hdc , oldFont);
                NativeMethods.ReleaseDC(editHandle , hdc);
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
                    _overlay.DrawUnderlines(new List<(System.Windows.Point, double, string, string, int, int)>()); // Clear all underlines
                    Console.WriteLine("No errors found in text");
                }
                UpdateErrorCount(apiResponse.spellCheckResponse.results.flagged_tokens.Count);
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
                    _lastActiveWindowTitle = newWindowTitle;
                    if (IsNotepad(newWindowTitle))
                    {
                        _notepadProcessId = GetProcessId(foregroundWindow);
                        ReadWindowContent(foregroundWindow);
                    }
                    //else
                    //{
                    //    _notepadProcessId = null;
                    //}
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
