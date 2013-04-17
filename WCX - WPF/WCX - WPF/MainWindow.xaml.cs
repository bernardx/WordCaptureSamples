using System;
using System.Runtime.InteropServices;
using System.Windows;
using WordCapture;
using System.Windows.Input;
using System.ComponentModel;
using System.Windows.Interop;

namespace WCX
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id,
           int fsModifiers, Key virtualKey);

        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        private WCaptureX m_wCapture = ComFactory.Instance.NewWCaptureX();
        private WMonitorX m_wMonitor = ComFactory.Instance.NewWMonitorX();

        public const uint E_CAPTURE_NONE = 0x00;
        public const uint E_CAPTURE_MOUSE = 0x01;
        public const uint E_CAPTURE_MOUSE_GEST = 0x02;
        public const uint E_CAPTURE_CARET = 0x04;
        public const uint E_CAPTURE_CURSOR = 0x08;
        public const uint E_CAPTURE_HOVER = 0x10;
        public const uint E_CAPTURE_HOTKEY = 0x20;
        public const uint E_CAPTURE_SELECTED_TEXT = 0x40;

        public const int HOTKEYF_ALT = 0x04;
        public const int HOTKEYF_SHIFT = 0x01;
        public const int HOTKEYF_CONTROL = 0x02;
        public const int HOTKEYF_EXT = 0x08;

        private int m_nHotkeyCursorId;
        private int m_nHotkeyCaretId;
        private int m_nHotkeySelectedTextId;

        private string clickedWord;
        public string ClickedWord 
        {
            get
            {
                return clickedWord;
            }
            set
            {
                clickedWord = value;
                OnNotifyPropertyChanged("ClickedWord");
            }
        }

        private string leftContext;
        public string LeftContext 
        {
            get
            {
                return leftContext;
            }
            set
            {
                leftContext = value;
                OnNotifyPropertyChanged("LeftContext");
            }
        }

        private string rightContext;
        public string RightContext 
        {
            get
            {
                return rightContext;
            }
            set
            {
                rightContext = value;
                OnNotifyPropertyChanged("RightContext");
            }
        }

        private string paragraph;
        public string Paragraph 
        {
            get
            {
                return paragraph;
            }
            set
            {
                paragraph = value;
                OnNotifyPropertyChanged("Paragraph");
            }        
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OnNotifyPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public static ushort LowWord(uint value)
        {
            return (ushort)(value & 0xFFFF);
        }
        public static ushort HighWord(uint value)
        {
            return (ushort)(value >> 16);
        }
        public static byte LowByte(ushort value)
        {
            return (byte)(value & 0xFF);
        }
        public static byte HighByte(ushort value)
        {
            return (byte)(value >> 8);
        }

        public void Hotkey2MonitorParams(int dwHotkey, out int nModifier, out int nKey)
        {
            nModifier = 0;
            nKey = (int)LowByte(LowWord((uint)dwHotkey));

            int wModifiers = HighByte(LowWord((uint)dwHotkey));
            if ((wModifiers & HOTKEYF_ALT) != 0)
                nModifier |= (int)W_KEY.wmKeyAlt;
            if ((wModifiers & HOTKEYF_CONTROL) != 0)
                nModifier |= (int)W_KEY.wmKeyCtrl;
            if ((wModifiers & HOTKEYF_EXT) != 0)
                nModifier |= (int)W_KEY.wmKeyWin;
            if ((wModifiers & HOTKEYF_SHIFT) != 0)
                nModifier |= (int)W_KEY.wmKeyShift;
        }

        public void StartMonitoring()
        {

            int color = 61440;

            W_LINE_STYLE wLineStyle = W_LINE_STYLE.wmLineDot;
            W_KEY wKey = W_KEY.wmKeyCtrl;
            W_MOUSE wMouse = W_MOUSE.wmMouseRight;
            wMouse |= W_MOUSE.wmMouseDoubleClick;

            int dwHotkeyCursor = 845;
            int dwHotkeyCaret = 835;
            int dwIdleTime = 2;

            m_wMonitor.LineStyle = (int)wLineStyle;
            m_wMonitor.LineColor = (uint)color;

            int dwHotkeySelectedText = 852;
            m_wMonitor.Start((int)wKey, (int)wMouse, true);

            m_wMonitor.Start2(dwIdleTime * 1000);

            IntPtr windowHandle = new WindowInteropHelper(Application.Current.MainWindow).Handle;

            int nModifier, nKey;
            nModifier = nKey = 0;
            Hotkey2MonitorParams(dwHotkeyCursor, out nModifier, out nKey);
            m_wMonitor.Start3(nModifier, nKey, out m_nHotkeyCursorId);
            RegisterHotKey(windowHandle, m_nHotkeyCursorId, nModifier, (Key)nKey);

            nModifier = nKey = 0;
            Hotkey2MonitorParams(dwHotkeyCaret, out nModifier, out nKey);
            m_wMonitor.Start3(nModifier, nKey, out m_nHotkeyCaretId);
            RegisterHotKey(windowHandle, m_nHotkeyCaretId, nModifier, (Key)nKey);

            nModifier = nKey = 0;
            Hotkey2MonitorParams(dwHotkeySelectedText, out nModifier, out nKey);
            m_wMonitor.Start3(nModifier, nKey, out m_nHotkeySelectedTextId);
            RegisterHotKey(windowHandle, m_nHotkeySelectedTextId, nModifier, (Key)nKey);
        }

        public WResult PerformCapture(int hWnd, int x1, int y1, int x2, int y2)
        {
            WInput objInput = ComFactory.Instance.NewWInput();	// set capture options
            int wOptions = 0;

            // set the get paragraph flag
            wOptions |= (int)W_CAPTURE_OPTIONS.wCaptureOptionsGetParagraph;
            // set the highlight word flag
            wOptions |= (int)W_CAPTURE_OPTIONS.wCaptureOptionsHighlightWords;
            // set the getContext flag
            wOptions |= (int)W_CAPTURE_OPTIONS.wCaptureOptionsGetContext;
            //set capture parameters
            objInput.Hwnd = hWnd;
            objInput.StartX = x1;
            objInput.StartY = y1;
            objInput.EndX = x2;
            objInput.EndY = y2;
            objInput.Options = wOptions;


            // set the number of context words
            objInput.ContextWordsLeft = 1;
            objInput.ContextWordsRight = 1;
            // declare the string which will get the results
            string strResult;

            WResult objResult;
            objResult = m_wCapture.Capture(objInput);

            // get the text from the capture
            strResult = objResult.Text;


            //use OCR if native method fails
            if (strResult == string.Empty)
            {
                wOptions |= (int)W_CAPTURE_OPTIONS.wCaptureOptionsGetTessOCRText;
                objInput.Options = wOptions;
                objResult = m_wCapture.Capture(objInput);
                strResult = objResult.Text;
            }

            return objResult;
        }

        private void CaptureEvent(int hwnd, int x1, int y1, int x2, int y2)
        {
            WResult objResult = PerformCapture(hwnd, x1, y1, x2, y2);

            if (objResult == null)
            {
                return;
            }

            ClickedWord = objResult.Text;
            LeftContext = objResult.LeftContext;
            RightContext = objResult.RightContext;
            Paragraph = objResult.Paragraph;
        }

        private void CaptureFullText(int x, int y)
        {
            UIControl spUIC = ComFactory.Instance.NewUIControl();

            spUIC.CreateFromScreenPoint(x, y);
            string strRes = spUIC.Value;
            if (strRes == string.Empty)
            {
                strRes = spUIC.Name;
            }
            ClickedWord = strRes;
        }

        private void MainWindow_Initialized(object sender, EventArgs e)
        {
            StartMonitoring();
            m_wMonitor.WEvent += new _IWMonitorXEvents_WEventEventHandler(CaptureEvent);
        }

        private void WCXExample_Unloaded(object sender, RoutedEventArgs e)
        {
            m_wMonitor.Stop();
            m_wCapture.EndCaptureSession();
        }

        private void WCXExample_SourceInitialized(object sender, EventArgs e)
        {
            HwndSource source = PresentationSource.FromVisual(this) as HwndSource;
            source.AddHook(WndProc);
        }

        protected IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            const int WM_HOTKEY = 0x0312;
            WCaptureX capture = ComFactory.Instance.NewWCaptureX();
            int Hwnd, X, Y;
            switch (msg)
            {
                case WM_HOTKEY:
                    {
                        handled = true;
                        int wp = wParam.ToInt32();
                        if (wp == m_nHotkeyCursorId)
                        {
                            capture.GetCursorInfo(out Hwnd, out X, out Y);
                            CaptureEvent(Hwnd, X, Y, X, Y);
                        }
                        else if (wp == m_nHotkeyCaretId)
                        {
                            capture.GetCaretInfo(out Hwnd, out X, out Y);
                            CaptureEvent(Hwnd, X, Y, X, Y);
                            break;
                        }
                        else if (wp == m_nHotkeySelectedTextId)
                        {
                            capture.GetCursorInfo(out Hwnd, out X, out Y);
                            CaptureEvent(Hwnd, X, Y, X, Y);
                            break;
                        }
                        break;
                    }
                default:
                    break;
            }
            return hwnd;
        }

    }
}
