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

namespace KeyboardTrackingApp;

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
        _suggestionsControl.SuggestionSelected += OnSuggestionSelected;
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
            // Get the handle of the Notepad's edit control
            IntPtr editHandle = FindWindowEx(notepadHandle , IntPtr.Zero , "Edit" , null);
            if (editHandle != IntPtr.Zero)
            {
                // Get the length of the text in the edit control
                int length = (int)SendMessage(editHandle , WM_GETTEXTLENGTH , IntPtr.Zero , null);
                if (length > 0)
                {
                    // Allocate a StringBuilder with the text length
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
                _suggestionsControl.SetSuggestions(new[] { "the" , "thee" });
                SuggestionsPopup.IsOpen = true;
            });
        }
        else
        {
            Dispatcher.Invoke(() =>
            {
                _suggestionsControl.SetSuggestions(null);
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
            _suggestionsControl.SetSuggestions(null);
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
        SendKeys(Keys.ControlKey , Keys.A);
        Thread.Sleep(100);
        SendKeys(Keys.ControlKey , Keys.V);
    }

    private void SendKeys(Keys modifierKey , Keys key)
    {
        keybd_event((byte)modifierKey , 0 , 0 , UIntPtr.Zero);
        keybd_event((byte)key , 0 , 0 , UIntPtr.Zero);
        keybd_event((byte)key , 0 , KEYEVENTF_KEYUP , UIntPtr.Zero);
        keybd_event((byte)modifierKey , 0 , KEYEVENTF_KEYUP , UIntPtr.Zero);
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