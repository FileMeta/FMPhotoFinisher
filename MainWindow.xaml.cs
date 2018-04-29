using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Interop;

namespace FMPhotoFinisher
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        internal delegate void StringDelegate(string text);
        internal delegate void VoidDelegate();

        public MainWindow()
        {
            InitializeComponent();
        }

        // TODO: Change TextBlock to FlowDocumentScrollViewer

        /// <summary>
        /// Write to the output control. Thread safe.
        /// </summary>
        /// <param name="text"></param>
        public void OutputWrite(string text)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new StringDelegate(OutputWriteInternal), text);
            }
            else
            {
                OutputWriteInternal(text);
            }
        }

        public void WriteLine(string text)
        {
            OutputWrite(text);
            OutputWrite("\r\n");
        }

        public void WriteLine()
        {
            OutputWrite("\r\n");
        }

        public void OutputWrite(string format, params object[] args)
        {
            OutputWrite(string.Format(format, args));
        }

        private void OutputWriteInternal(string text)
        {
            m_BodyDoc.ContentEnd.InsertTextInRun(text);
            ScrollToBottom();
        }

        public void SetProgress(string text)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.BeginInvoke(new StringDelegate(SetProgressInternal), text);
            }
            else
            {
                OutputWriteInternal(text);
            }
        }

        private void SetProgressInternal(string text)
        {
            m_Progress.Text = text;
        }

        IntPtr m_hwnd = IntPtr.Zero;

        /// <summary>
        /// Get window handle. Thread safe.
        /// </summary>
        public IntPtr GetWindowHandle()
        {
            if (m_hwnd == IntPtr.Zero)
            {
                if (!Dispatcher.CheckAccess())
                {
                    Dispatcher.Invoke(new VoidDelegate(GetWindowHandleInternal));
                }
                else
                {
                    GetWindowHandleInternal();
                }
            }
            return m_hwnd;
        }

        private void GetWindowHandleInternal()
        {
            m_hwnd = new WindowInteropHelper(this).Handle;
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var thread = new OperationsThread(this);
            thread.Start();
        }

        ScrollViewer m_bodyDocScrollViewer = null;

        private void ScrollToBottom()
        {
            if (m_bodyDocScrollViewer == null)
            {
                DependencyObject obj = m_BodyViewer;

                do
                {
                    if (VisualTreeHelper.GetChildrenCount(obj) > 0)
                    {
                        obj = VisualTreeHelper.GetChild(obj as Visual, 0);
                        m_bodyDocScrollViewer = obj as ScrollViewer;
                    }
                    else
                    {
                        break;
                    }
                }
                while (m_bodyDocScrollViewer == null);
            }
            if (m_bodyDocScrollViewer != null)
            {
                m_bodyDocScrollViewer.ScrollToBottom();
            }
        }

    }
}
