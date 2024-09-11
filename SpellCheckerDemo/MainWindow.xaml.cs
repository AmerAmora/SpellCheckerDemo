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
        private SuggestionsControl SuggestionsControl;

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern void keybd_event(byte bVk , byte bScan , uint dwFlags , UIntPtr dwExtraInfo);

        [DllImport("user32.dll")]
        public static extern bool AttachThreadInput(uint idAttach , uint idAttachTo , bool fAttach);

        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd , IntPtr ProcessId);

        [DllImport("user32.dll")]
        static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        static extern int ToUnicode(uint wVirtKey , uint wScanCode , byte[] lpKeyState ,
                                    [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff , int cchBuff , uint wFlags);

        [DllImport("user32.dll")]
        static extern uint MapVirtualKey(uint uCode , uint uMapType);

        [DllImport("kernel32.dll")]
        static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd , StringBuilder text , int count);

        private const int KEYEVENTF_KEYUP = 0x0002;

        public MainWindow()
        {
            InitializeComponent();
            _keyboardHook = new GlobalKeyboardHook();
            _keyboardHook.KeyPressed += OnKeyPressed;

            // Initialize SuggestionsControl
            SuggestionsControl = new SuggestionsControl();
            SuggestionsControl.SuggestionSelected += OnSuggestionSelected;
            SuggestionsPopup.Child = SuggestionsControl;

            _windowCheckTimer = new DispatcherTimer();
            _windowCheckTimer.Interval = TimeSpan.FromMilliseconds(500);
            _windowCheckTimer.Tick += WindowCheckTimer_Tick;
            _windowCheckTimer.Start();
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
                        ReadWindowContent();
                    }
                }
            }
        }

        private bool IsNotepad(string windowTitle)
        {
            return windowTitle.EndsWith(".txt - Notepad") || windowTitle.Equals("Untitled - Notepad") || windowTitle.EndsWith("Notepad");
        }

        private void ReadWindowContent()
        {
            try
            {
                // Save current clipboard content
                string oldClipboardContent = Clipboard.GetText();

                // Select all text and copy to clipboard
                SendKeys(Keys.ControlKey , Keys.A);
                Thread.Sleep(100);
                SendKeys(Keys.ControlKey , Keys.C);
                Thread.Sleep(100);

                // Read from clipboard
                string content = Clipboard.GetText();

                // Restore old clipboard content
                Clipboard.SetText(oldClipboardContent);

                _allText.Clear();
                _allText.Append(content);
                _currentWord.Clear();

                Dispatcher.Invoke(() =>
                {
                    HighlightTeh();
                });

                // Log or display the read content (for debugging purposes)
                Console.WriteLine("Read content from Notepad: " + content);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Notepad content: {ex.Message}");
            }
        }

        private void SendKeys(Keys modifierKey , Keys key)
        {
            keybd_event((byte)modifierKey , 0 , 0 , UIntPtr.Zero);
            keybd_event((byte)key , 0 , 0 , UIntPtr.Zero);
            keybd_event((byte)key , 0 , KEYEVENTF_KEYUP , UIntPtr.Zero);
            keybd_event((byte)modifierKey , 0 , KEYEVENTF_KEYUP , UIntPtr.Zero);
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
                    SuggestionsControl.SetSuggestions(new[] { "the" , "thee" });
                    SuggestionsPopup.IsOpen = true;
                });
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    SuggestionsControl.SetSuggestions(null);
                    SuggestionsPopup.IsOpen = false;
                });
            }
        }

        private void OnSuggestionSelected(object sender , string suggestion)
        {
            string currentText = _allText.ToString();
            string currentWord = _currentWord.ToString();

            currentText = currentText.Replace("teh" , suggestion);
            _allText.Clear();
            _allText.Append(currentText);

            ReplaceWordInActiveApplication(suggestion);

            _currentWord.Clear();

            Dispatcher.Invoke(() =>
            {
                HighlightTeh();
                SuggestionsControl.SetSuggestions(null);
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
                    var incorrectWord = new Run("teh") { Background = Brushes.Red , Foreground = Brushes.White };
                    CurrentWordTextBlock.Inlines.Add(incorrectWord);

                    var arrow = new Run(" => ");
                    CurrentWordTextBlock.Inlines.Add(arrow);

                    var correctWord = new Run("the") { Foreground = Brushes.Green };
                    CurrentWordTextBlock.Inlines.Add(correctWord);

                    var acceptButton = new Button { Content = "Accept" , Margin = new Thickness(5) };
                    acceptButton.Click += (s , e) => OnSuggestionSelected(this , "the");
                    CurrentWordTextBlock.Inlines.Add(acceptButton);

                    var rejectButton = new Button { Content = "Reject" , Margin = new Thickness(5) };
                    rejectButton.Click += (s , e) => SuggestionsPopup.IsOpen = false;
                    CurrentWordTextBlock.Inlines.Add(rejectButton);
                }
            }
        }

        private void ReplaceWordInActiveApplication(string replacement)
        {
            if (_lastActiveWindowHandle != IntPtr.Zero)
            {
                uint currentThreadId = GetCurrentThreadId();
                uint foregroundThreadId = GetWindowThreadProcessId(_lastActiveWindowHandle , IntPtr.Zero);

                AttachThreadInput(currentThreadId , foregroundThreadId , true);
                SetForegroundWindow(_lastActiveWindowHandle);
                AttachThreadInput(currentThreadId , foregroundThreadId , false);

                keybd_event((byte)Keys.ControlKey , 0 , 0 , UIntPtr.Zero);
                keybd_event((byte)Keys.A , 0 , 0 , UIntPtr.Zero);
                keybd_event((byte)Keys.A , 0 , KEYEVENTF_KEYUP , UIntPtr.Zero);
                keybd_event((byte)Keys.ControlKey , 0 , KEYEVENTF_KEYUP , UIntPtr.Zero);

                keybd_event((byte)Keys.Delete , 0 , 0 , UIntPtr.Zero);
                keybd_event((byte)Keys.Delete , 0 , KEYEVENTF_KEYUP , UIntPtr.Zero);
                Thread.Sleep(10);

                foreach (char c in replacement)
                {
                    byte vk = ToVirtualKey(c);
                    keybd_event(vk , 0 , 0 , UIntPtr.Zero);
                    keybd_event(vk , 0 , KEYEVENTF_KEYUP , UIntPtr.Zero);
                }
            }
        }

        private byte ToVirtualKey(char c)
        {
            return (byte)MapVirtualKey((uint)char.ToUpper(c) , 2);
        }

        protected override void OnClosed(EventArgs e)
        {
            _keyboardHook.Dispose();
            base.OnClosed(e);
        }
    }

    public class SuggestionsControl : StackPanel
    {
        public event EventHandler<string> SuggestionSelected;

        public void SetSuggestions(IEnumerable<string> suggestions)
        {
            Children.Clear();
            if (suggestions != null)
            {
                foreach (var suggestion in suggestions)
                {
                    var textBlock = new TextBlock
                    {
                        Text = suggestion ,
                        Padding = new Thickness(5) ,
                        Background = Brushes.LightGray
                    };
                    textBlock.MouseEnter += (s , e) =>
                    {
                        textBlock.Background = Brushes.Gray;
                        SuggestionSelected?.Invoke(this , suggestion);
                    };
                    textBlock.MouseLeave += (s , e) => textBlock.Background = Brushes.LightGray;
                    Children.Add(textBlock);
                }
            }
        }
    }
}