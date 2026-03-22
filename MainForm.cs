using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PstMerger
{
    public partial class MainForm : Form
    {
        private PstService _pstService;
        private System.Threading.CancellationTokenSource _cts;
        private string _logFile;
        private bool _skipDuplicateChecking;

        public MainForm(bool skipDuplicateChecking = false)
        {
            _skipDuplicateChecking = skipDuplicateChecking;
            
            try
            {
                InitializeComponent();
                string duplicateCheckStatus = _skipDuplicateChecking ? "DISABLED" : "ENABLED";
                this.Text = string.Format("PST Merge Tool v{0} (Duplicate checking: {1})", Application.ProductVersion, duplicateCheckStatus);
                _pstService = new PstService();
                _logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, string.Format("PstMerge_{0:yyyyMMdd_HHmmss}.log", DateTime.Now));
                Log("Tool initialized. Enterprise Log started: " + _logFile);
                Log("Duplicate checking: " + duplicateCheckStatus);
                
                // Set form closing handler to catch exceptions during shutdown
                this.FormClosing += MainForm_FormClosing;
                this.FormClosed += MainForm_FormClosed;
                
                // Kill Outlook process to prevent COM conflicts during merge
                KillOutlookProcess();
            }
            catch (Exception ex)
            {
                string msg = string.Format("CRITICAL ERROR during initialization: {0}\nStackTrace: {1}", ex.Message, ex.StackTrace);
                try { if (_logFile != null) File.AppendAllText(_logFile, msg + Environment.NewLine); } catch { }
                throw;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (_cts != null)
                {
                    _cts.Cancel();
                    _cts.Dispose();
                }
            }
            catch (Exception ex)
            {
                Log("ERROR during form closing: " + ex.Message);
            }
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            try
            {
                // Clean up the service instance
                if (_pstService != null)
                {
                    _pstService = null;
                }
            }
            catch (Exception ex)
            {
                Log("ERROR during form closed cleanup: " + ex.Message);
            }
        }

        private void KillOutlookProcess()
        {
            try
            {
                System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
                if (processes.Length > 0)
                {
                    Log(string.Format("Found {0} Outlook process(es). Terminating...", processes.Length));
                    
                    foreach (var process in processes)
                    {
                        try
                        {
                            process.Kill();
                            process.WaitForExit(5000); // Wait up to 5 seconds for graceful termination
                            Log("Outlook process terminated successfully.");
                        }
                        catch (Exception ex)
                        {
                            Log("Warning: Failed to terminate Outlook: " + ex.Message);
                        }
                    }
                    
                    System.Threading.Thread.Sleep(1000); // Give Outlook time to fully close
                }
            }
            catch (Exception ex)
            {
                Log("Warning: Error checking for Outlook process: " + ex.Message);
            }
        }

        private void btnBrowseSource_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtSourceFolder.Text = fbd.SelectedPath;
                }
            }
        }

        private void btnBrowseDest_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Outlook Data File (*.pst)|*.pst";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    txtDestPst.Text = sfd.FileName;
                }
            }
        }

        private void btnFixRegistry_Click(object sender, EventArgs e)
        {
            try
            {
                Log("Applying PST size limit fixes to registry...");
                
                // We target Outlook 15.0 and 16.0
                string[] versions = { "15.0", "16.0" };
                foreach (var v in versions)
                {
                    string keyPath = string.Format(@"Software\Microsoft\Office\{0}\Outlook\PST", v);
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath))
                    {
                        if (key != null)
                        {
                            // Values in MB. 2000000 MB = ~2 TB (effectively unlimited)
                            key.SetValue("MaxLargeFileSize", 2000000, RegistryValueKind.DWord);
                            key.SetValue("WarnLargeFileSize", 1900000, RegistryValueKind.DWord);
                        }
                    }
                }

                Log("SUCCESS: PST size limits increased to 2TB (effectively unlimited). Please restart Outlook if it's open.");
                MessageBox.Show("Registry updated. Please restart Outlook if it is currently running.", "Registry Fix Applied", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("ERROR applying registry fix: " + ex.Message);
                MessageBox.Show("Failed to update registry. You may need to run as Administrator.\n\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnStartMerge_Click(object sender, EventArgs e)
        {
            string sourceDir = txtSourceFolder.Text;
            string destFile = txtDestPst.Text;

            if (string.IsNullOrEmpty(sourceDir) || !Directory.Exists(sourceDir))
            {
                MessageBox.Show("Please select a valid source folder.");
                return;
            }

            if (string.IsNullOrEmpty(destFile))
            {
                MessageBox.Show("Please select a destination PST file.");
                return;
            }

            // Check Disk Space
            try
            {
                string drive = Path.GetPathRoot(Path.GetFullPath(destFile));
                DriveInfo di = new DriveInfo(drive);
                long totalSourceSize = 0;
                var pstFilesCheck = Directory.GetFiles(sourceDir, "*.pst", SearchOption.TopDirectoryOnly);
                foreach (var f in pstFilesCheck) totalSourceSize += new FileInfo(f).Length;
                
                if (di.AvailableFreeSpace < (totalSourceSize * 1.1)) // 10% buffer
                {
                    var msg = string.Format("Warning: You might not have enough disk space on {0}.\nAvailable: {1} GB\nRequired (est): {2} GB\n\nContinue anyway?", 
                        drive, di.AvailableFreeSpace / 1024 / 1024 / 1024, (totalSourceSize * 1.1) / 1024 / 1024 / 1024);
                    if (MessageBox.Show(msg, "Disk Space Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                        return;
                }
            }
            catch { }

            var pstFiles = Directory.GetFiles(sourceDir, "*.pst", SearchOption.TopDirectoryOnly);
            if (pstFiles.Length == 0)
            {
                MessageBox.Show("No PST files found in source folder.");
                return;
            }

            // Check for running Outlook before starting
            ShowOutlookWarning();

            btnStartMerge.Enabled = false;
            btnFixRegistry.Enabled = false;
            btnCancel.Visible = true;
            btnCancel.Enabled = true;
            progressBar.Value = 0;
            progressBar.Maximum = pstFiles.Length;

            Log(string.Format("Starting merge of {0} files...", pstFiles.Length));

            try
            {
                _cts = new System.Threading.CancellationTokenSource();
                
                // Run the merge operation asynchronously on STA thread to avoid COM threading issues
                var mergeTask = System.Threading.Tasks.Task.Run(async () =>
                {
                    await _pstService.MergeFilesAsync(pstFiles, destFile, _cts.Token, (progress, message) =>
                {
                    this.Invoke(new Action(() =>
                    {
                        if (progress == -2)
                        {
                            SetCurrentCopyStatus(message);
                            return;
                        }
                        if (progress == -3)
                        {
                            // Log skipped items to main log only, don't clutter current copy window
                            Log(message);
                            return;
                        }

                        Log(message);
                        if (progress > 0) progressBar.Value = Math.Min(progress, progressBar.Maximum);
                    }));
                }, _skipDuplicateChecking);
                });

                await mergeTask;

                if (_cts.Token.IsCancellationRequested)
                {
                    Log("STOPPED: Merge was cancelled by user.");
                    MessageBox.Show("Process was cancelled.", "Stopped", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    Log("COMPLETED: All PST files merged successfully.");
                    MessageBox.Show("Merge completed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (OperationCanceledException)
            {
                Log("STOPPED: Merge was cancelled.");
            }
            catch (Exception ex)
            {
                Log("FATAL ERROR: " + ex.Message);
                Log("ERROR DETAILS: " + ex.GetType().Name);
                if (!string.IsNullOrEmpty(ex.StackTrace))
                    Log("STACK TRACE: " + ex.StackTrace);
                
                // Log inner exceptions recursively
                Exception inner = ex.InnerException;
                int depth = 1;
                while (inner != null && depth <= 5)
                {
                    Log(string.Format("INNER EXCEPTION {0}: {1}", depth, inner.Message));
                    if (!string.IsNullOrEmpty(inner.StackTrace))
                        Log(string.Format("INNER STACK {0}: {1}", depth, inner.StackTrace));
                    inner = inner.InnerException;
                    depth++;
                }
                
                MessageBox.Show("An error occurred during the merge:\n\n" + ex.Message, "Merge Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnStartMerge.Enabled = true;
                btnFixRegistry.Enabled = true;
                btnCancel.Visible = false;
                if (_cts != null) _cts.Dispose();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (_cts != null)
            {
                Log("Cancellation requested... waiting for current item to finish...");
                _cts.Cancel();
                btnCancel.Enabled = false;
            }
        }

        private bool IsOutlookRunning()
        {
            try
            {
                System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
                return processes.Length > 0;
            }
            catch
            {
                return false;
            }
        }

        private void ShowOutlookWarning()
        {
            if (IsOutlookRunning())
            {
                var result = MessageBox.Show(
                    "Outlook appears to be running. This can cause COM errors during PST merging.\n\n" +
                    "For best results, please close Outlook before merging PST files.\n\n" +
                    "Continue anyway?",
                    "Outlook Running Warning",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning
                );
                
                if (result == DialogResult.No)
                {
                    return; // User cancelled
                }
            }
        }

        private void SetCurrentCopyStatus(string message)
        {
            try
            {
                if (txtCurrentCopy.InvokeRequired)
                {
                    txtCurrentCopy.Invoke(new Action(() => SetCurrentCopyStatus(message)));
                    return;
                }

                string line = string.Format("[{0:HH:mm:ss}] {1}", DateTime.Now, message);
                txtCurrentCopy.AppendText(line + Environment.NewLine);
                txtCurrentCopy.SelectionStart = txtCurrentCopy.Text.Length;
                txtCurrentCopy.ScrollToCaret();
            }
            catch
            {
                // keep going even if current copy update fails
            }
        }

        private void Log(string message)
        {
            try
            {
                if (txtLog.InvokeRequired)
                {
                    txtLog.Invoke(new Action(() => Log(message)));
                    return;
                }
                string line = string.Format("[{0:HH:mm:ss}] {1}", DateTime.Now, message);
                txtLog.AppendText(line + Environment.NewLine);
                
                // Persistent File Logging
                try { File.AppendAllText(_logFile, line + Environment.NewLine); } catch { }
            }
            catch (Exception ex)
            {
                // If UI logging fails, still try file logging
                try
                {
                    string line = string.Format("[{0:HH:mm:ss}] LOG ERROR: {1} - Message was: {2}", DateTime.Now, ex.Message, message);
                    File.AppendAllText(_logFile, line + Environment.NewLine);
                }
                catch { }
            }
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            Version v = new Version(Application.ProductVersion);
            string displayVersion = string.Format("{0}.{1}.{2}", v.Major, v.Minor, v.Build);

            string about = string.Format("PST Merge Tool v{0}\n\n", displayVersion) +
                           "Developed by: Mithun\n" +
                           "© DataGuardNXT 2026\n\n" +
                           "All Rights Reserved.\n\n" +
                           "Enterprise-grade PST merging solution\n" +
                           "for large mailbox recovery.";
            MessageBox.Show(about, "About PST Merge Tool", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
