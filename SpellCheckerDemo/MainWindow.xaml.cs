using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Keys = System.Windows.Forms.Keys;
using System.Windows.Threading;

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

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern IntPtr SendMessage(IntPtr hWnd , uint Msg , IntPtr wParam , StringBuilder lParam);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern bool GetWindowText(IntPtr hWnd , StringBuilder text , int count);

        [DllImport("user32.dll" , CharSet = CharSet.Auto)]
        private static extern void keybd_event(byte bVk , byte bScan , uint dwFlags , UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern bool GetKeyboardState(byte[] pbKeyState);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode , uint uMapType);

        [DllImport("user32.dll")]
        private static extern int ToUnicode(uint wVirtKey , uint wScanCode , byte[] pKeyState , StringBuilder pChar , int wCharSize , uint wFlags);

        [DllImport("user32.dll" , SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr parentHandle , IntPtr childAfter , string lpszClass , string lpszWindow);

        private const uint WM_GETTEXT = 0x000D;
        private const uint WM_GETTEXTLENGTH = 0x000E;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x0101;
        private const int KEYEVENTF_KEYUP = 0x0002;

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
                Interval = TimeSpan.FromMilliseconds(500)
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
        }

        private void WindowCheckTimer_Tick(object sender , EventArgs e)
        {
            IntPtr foregroundWindow = GetForegroundWindow();
            if (foregroundWindow != _lastActiveWindowHandle)
            {
                _lastActiveWindowHandle = foregroundWindow;
                StringBuilder sb = new StringBuilder(256);
                GetWindowText(foregroundWindow , sb , 256);
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
                IntPtr editHandle = FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);
                if (editHandle != IntPtr.Zero)
                {
                    int length = (int)SendMessage(editHandle , WM_GETTEXTLENGTH , IntPtr.Zero , null);
                    if (length > 0)
                    {
                        StringBuilder sb = new StringBuilder(length + 1);
                        SendMessage(editHandle , WM_GETTEXT , (IntPtr)sb.Capacity , sb);

                        _allText.Clear();
                        _allText.Append(sb.ToString());
                        _currentWord.Clear();

                        Dispatcher.Invoke(() =>
                        {
                            HighlightTeh();
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
            _lastActiveWindowHandle = GetForegroundWindow();

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
                HighlightTeh();
            });
        }

        private char GetCharFromKey(Keys key)
        {
            byte[] keyboardState = new byte[256];
            GetKeyboardState(keyboardState);

            uint scanCode = MapVirtualKey((uint)key , 0);
            StringBuilder stringBuilder = new StringBuilder(2);

            int result = ToUnicode((uint)key , scanCode , keyboardState , stringBuilder , stringBuilder.Capacity , 0);
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

        private void ReplaceWordInActiveApplication(string replacementWord)
        {
            IntPtr hWnd = GetForegroundWindow();
            if (hWnd == IntPtr.Zero)
                return;

            SetForegroundWindow(hWnd);

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
            keybd_event((byte)key , 0 , 0 , UIntPtr.Zero);
            keybd_event((byte)key , 0 , KEYEVENTF_KEYUP , UIntPtr.Zero);
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
                    });
                }
            }
        }

        private string GetNotepadContent()
        {
            IntPtr notepadHandle = GetForegroundWindow();
            IntPtr editHandle = FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);
            if (editHandle != IntPtr.Zero)
            {
                int length = (int)SendMessage(editHandle , WM_GETTEXTLENGTH , IntPtr.Zero , null);
                if (length > 0)
                {
                    StringBuilder sb = new StringBuilder(length + 1);
                    SendMessage(editHandle , WM_GETTEXT , (IntPtr)sb.Capacity , sb);
                    return sb.ToString();
                }
            }
            return string.Empty;
        }
    }

    public class SuggestionsControl : StackPanel
    {
        public event EventHandler SuggestionAccepted;
        public event EventHandler SuggestionCancelled;

        private TextBlock _suggestionTextBlock;
        private Button _acceptButton;
        private Button _cancelButton;

        public SuggestionsControl()
        {
            _suggestionTextBlock = new TextBlock
            {
                Padding = new Thickness(5) ,
                Background = Brushes.LightYellow
            };

            _acceptButton = new Button
            {
                Content = "Accept" ,
                Margin = new Thickness(5) ,
                Padding = new Thickness(5)
            };
            _acceptButton.Click += (s , e) => SuggestionAccepted?.Invoke(this , EventArgs.Empty);

            _cancelButton = new Button
            {
                Content = "Cancel" ,
                Margin = new Thickness(5) ,
                Padding = new Thickness(5)
            };
            _cancelButton.Click += (s , e) => SuggestionCancelled?.Invoke(this , EventArgs.Empty);

            Children.Add(_suggestionTextBlock);
            Children.Add(_acceptButton);
            Children.Add(_cancelButton);
        }

        public void SetSuggestion(string suggestion)
        {
            if (string.IsNullOrEmpty(suggestion))
            {
                Visibility = Visibility.Collapsed;
            }
            else
            {
                _suggestionTextBlock.Text = $"Suggested: {suggestion}";
                Visibility = Visibility.Visible;
            }
        }
    }
}