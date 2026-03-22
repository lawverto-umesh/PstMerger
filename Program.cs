using System;
using System.IO;
using System.Windows.Forms;

namespace PstMerger
{
    static class Program
    {
        private static string _crashLogFile;
        public static bool SkipDuplicateChecking { get; private set; }

        private static void ParseCommandLineArgs(string[] args)
        {
            SkipDuplicateChecking = false;
            
            foreach (string arg in args)
            {
                string lowerArg = arg.ToLowerInvariant();
                if (lowerArg == "--skip-duplicates" || lowerArg == "-s" || lowerArg == "/skip-duplicates")
                {
                    SkipDuplicateChecking = true;
                    break;
                }
            }
        }

        [STAThread]
        static void Main(string[] args)
        {
            // Parse command line arguments
            ParseCommandLineArgs(args);

            // Set up crash logging before anything else
            _crashLogFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, string.Format("PstMerge_CRASH_{0:yyyyMMdd_HHmmss}.log", DateTime.Now));
            // Set up crash logging before anything else
            _crashLogFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, string.Format("PstMerge_CRASH_{0:yyyyMMdd_HHmmss}.log", DateTime.Now));
            
            // Register global exception handlers
            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            Application.ThreadException += Application_ThreadException;
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MainForm(SkipDuplicateChecking));
            }
            catch (Exception ex)
            {
                LogCrash("UNHANDLED EXCEPTION IN MAIN", ex);
                MessageBox.Show(string.Format("CRITICAL ERROR: {0}\n\nCheck log file for details.", ex.Message),
                    "Application Crash", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Exception ex = e.ExceptionObject as Exception;
            if (ex != null)
                LogCrash("UNHANDLED APPDOMAIN EXCEPTION", ex);
            else
                LogCrash("UNHANDLED APPDOMAIN EXCEPTION", new Exception(e.ExceptionObject.ToString()));
        }

        private static void Application_ThreadException(object sender, System.Threading.ThreadExceptionEventArgs e)
        {
            LogCrash("UNHANDLED UI THREAD EXCEPTION", e.Exception);
            MessageBox.Show(string.Format("ERROR: {0}\n\nCheck log file for details.", e.Exception.Message),
                "Application Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private static void LogCrash(string title, Exception ex)
        {
            try
            {
                string message = string.Format("[{0:yyyy-MM-dd HH:mm:ss.fff}] {1}\n" +
                    "Message: {2}\n" +
                    "Type: {3}\n" +
                    "HResult: {4}\n" +
                    "Source: {5}\n" +
                    "StackTrace:\n{6}\n\n",
                    DateTime.Now, title, ex.Message, ex.GetType().FullName,
                    ex.HResult, ex.Source, ex.StackTrace);

                if (ex.InnerException != null)
                    message += string.Format("InnerException: {0}\nStackTrace: {1}\n\n",
                        ex.InnerException.Message, ex.InnerException.StackTrace);

                File.AppendAllText(_crashLogFile, message);
            }
            catch { }
        }
    }
}
