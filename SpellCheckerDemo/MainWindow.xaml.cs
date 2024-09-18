using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Documents;
using Keys = System.Windows.Forms.Keys;
using System.Windows.Threading;
using SpellCheckerDemo;
using Brushes = System.Windows.Media.Brushes;
using Point = System.Drawing.Point;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;
using System.Runtime.InteropServices;

namespace KeyboardTrackingApp;

public partial class MainWindow : System.Windows.Window
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
    private int? _microsoftWordProcessId;

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
                else if (IsMicroSoftWord(newWindowTitle))
                {
                    _microsoftWordProcessId = GetProcessId(foregroundWindow);
                    ReadWordContent();
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
                    HighlightTeh(); // Custom method to highlight specific words
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
            HighlightTeh();
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

    private void CheckForIncorrectWords()
    {
        string text = _allText.ToString();
        int index = text.LastIndexOf("teh" , StringComparison.OrdinalIgnoreCase);

        if (index != -1)
        {
            NativeMethods.RECT clientRect;
            NativeMethods.GetClientRect(_lastActiveWindowHandle , out clientRect);

            NativeMethods.POINT point = new NativeMethods.POINT { X = 0 , Y = 0 };
            NativeMethods.ClientToScreen(_lastActiveWindowHandle , ref point);

            // Get the position of the word "teh"
            Point tehPosition;
            if (IsMicroSoftWord(_lastActiveWindowTitle))
            {
                tehPosition = GetPositionOfWordInWord("teh");
            }
            else
            {
                tehPosition = GetPositionOfWord("teh");
            }

            if (tehPosition != Point.Empty)
            {
                // Calculate the position of "teh" based on the actual text position
                var textRect = new Rect(point.X + tehPosition.X , point.Y + tehPosition.Y , 26 , 20);

                _overlay.DrawUnderline(textRect);
            }
            else
            {
                _overlay.DrawUnderline(new Rect(0 , 0 , 0 , 0)); // Clear the underline
            }
        }
        else
        {
            _overlay.DrawUnderline(new Rect(0 , 0 , 0 , 0)); // Clear the underline
        }
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

    private Point GetPositionOfWordInWord(string word)
    {
        try
        {
            Application wordApp = (Application)ReplaceMarshall.GetActiveObject("Word.Application");
            Document doc = wordApp.ActiveDocument;
            Microsoft.Office.Interop.Word.Range rng = doc.Content;

            rng.Find.ClearFormatting();
            rng.Find.Text = word;
            rng.Find.Execute();

            if (rng.Find.Found)
            {
                int left = rng.Information[WdInformation.wdHorizontalPositionRelativeToPage];
                int top = rng.Information[WdInformation.wdVerticalPositionRelativeToPage];

                // Convert from points to pixels
                left = (int)( left * 96 / 72 );
                top = (int)( top * 96 / 72 );

                return new Point(left , top);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error getting word position in Word: {ex.Message}");
        }

        return Point.Empty;
    }

    private void ReplaceWordInActiveApplication(string replacementWord)
    {
        IntPtr hWnd = NativeMethods.GetForegroundWindow();
        if (hWnd == IntPtr.Zero)
            return;

        NativeMethods.SetForegroundWindow(hWnd);

        if (IsMicroSoftWord(_lastActiveWindowTitle))
        {
            ReplaceWordInWord("teh" , replacementWord);
        }
        else
        {
            // Existing code for Notepad
            string currentContent = _allText.ToString();
            int lastIndex = currentContent.LastIndexOf("teh" , StringComparison.OrdinalIgnoreCase);
            if (lastIndex == -1)
                return;

            int backspaces = currentContent.Length - lastIndex - 3;

            for (int i = 0 ; i < backspaces + 3 ; i++)
            {
                SendKeys(Keys.Back);
            }

            foreach (char c in replacementWord)
            {
                SendKeys((Keys)char.ToUpper(c));
            }

            _allText.Remove(lastIndex , 3);
            _allText.Insert(lastIndex , replacementWord);
        }
    }

    private void ReplaceWordInWord(string oldWord , string newWord)
    {
        try
        {
            Application wordApp = (Application)ReplaceMarshall.GetActiveObject("Word.Application");
            Document doc = wordApp.ActiveDocument;
            Microsoft.Office.Interop.Word.Range rng = doc.Content;

            rng.Find.ClearFormatting();
            rng.Find.Text = oldWord;
            rng.Find.Replacement.Text = newWord;
            rng.Find.Execute(Replace: WdReplace.wdReplaceAll);

            // Update our internal text representation
            string updatedContent = GetMicrosoftWordContent();
            _allText.Clear();
            _allText.Append(updatedContent);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error replacing word in Word: {ex.Message}");
        }
    }

    private void SendKeys(Keys key)
    {
        NativeMethods.keybd_event((byte)key , 0 , 0 , UIntPtr.Zero);
        NativeMethods.keybd_event((byte)key , 0 , NativeMethods.KEYEVENTF_KEYUP , UIntPtr.Zero);
    }

    private void ContentSyncTimer_Tick(object sender , EventArgs e)
    {
        IntPtr foregroundWindow = NativeMethods.GetForegroundWindow();

        _lastActiveWindowHandle = foregroundWindow;
        StringBuilder sb = new StringBuilder(256); // This should now properly handle Arabic text
        NativeMethods.GetWindowText(foregroundWindow , sb , 256);
        string newWindowTitle = sb.ToString();
        var isnotePad = IsNotepad(newWindowTitle);
        var isword = IsMicroSoftWord(newWindowTitle);
        if (isnotePad)
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

        if (isword)
        {
            string microsoftWord = GetMicrosoftWordContent();
            if (microsoftWord != _allText.ToString())
            {
                _allText.Clear();
                _allText.Append(microsoftWord);
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
