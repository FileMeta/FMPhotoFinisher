using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
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
using System.Diagnostics;

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
            m_Progress.Text = text??string.Empty;
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
            StartAppAndMonitor(null, null);
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

        #region Application Monitor Thread

        Process m_proc;
        Thread m_stdOutThread;
        Thread m_stdErrThread;

        void StartAppAndMonitor(string appName, string commandLine)
        {

            try
            {
                // Compose arguments
                string exe = @"C:\Users\brand\source\FileMeta\FMPhotoFinisher\FMPhotoFinish\bin\Debug\FMPhotoFinish.exe";
                string arguments = @"-s ""E:\SampleData\PhotoFinisherUnitTest"" -d ""E:\FMPhotoFinisherTestOutput"" -autorot -orderedNames -transcode";

                // Prepare process start
                var psi = new ProcessStartInfo(exe, arguments);
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true; // Set to false if you want to monitor
                psi.RedirectStandardOutput = true;
                psi.StandardOutputEncoding = Encoding.UTF8;
                psi.RedirectStandardError = true;
                psi.StandardErrorEncoding = Encoding.UTF8;

                m_proc = Process.Start(psi);

                m_stdOutThread = new Thread(StdOutThreadMain);
                m_stdOutThread.Start();

                m_stdErrThread = new Thread(StdErrThreadMain);
                m_stdErrThread.Start();

                // TODO: Set up an event on the process and clean up when it exits
            }
            catch (Exception err)
            {
                WriteLine(err.ToString());
            }
        }

        private void StdOutThreadMain()
        {
            try
            {
                for (; ; )
                {
                    string line = m_proc.StandardOutput.ReadLine();
                    if (line == null) break;
                    WriteLine(line);
                }

            }
            catch (Exception err)
            {
                WriteLine(err.ToString());
            }
        }

        private void StdErrThreadMain()
        {
            try
            {
                for (; ; )
                {
                    string line = m_proc.StandardOutput.ReadLine();
                    if (line == null) break;
                    SetProgress(line);
                }

            }
            catch (Exception err)
            {
                WriteLine(err.ToString());
            }
        }

        #endregion Application Monitor Thread

    } // Class ProgressWindow

}
