using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.AccessControl;
using System.Security.Principal;
using System.Threading;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PstMerger
{
    public class PstService
    {
        [DllImport("ole32.dll")]
        private static extern int CoInitializeSecurity(IntPtr pVoid, int cAuthSvc, IntPtr asAuthSvc, IntPtr pReserved1, uint dwAuthnLevel, uint dwImpLevel, IntPtr pAuthList, uint dwCapabilities, IntPtr pReserved3);

        // Optimize: Use parallel processing with controlled concurrency
        private const int MaxConcurrentPstFiles = 3; // Process up to 3 PST files simultaneously
        private const int MaxConcurrentItemsPerFolder = 10; // Copy up to 10 items concurrently per folder

        // Performance tracking
        private DateTime _startTime;
        private int _totalItemsProcessed;
        private int _totalDuplicatesSkipped;
        private int _totalLargeItemsHandled;
        private int _totalItemsSkipped;

        private void EnsurePstOwnershipAndFullControl(string path, Action<int, string> onProgress)
        {
            try
            {
                if (!File.Exists(path))
                    return;

                var fileInfo = new FileInfo(path);
                if (fileInfo.IsReadOnly)
                {
                    fileInfo.IsReadOnly = false;
                    onProgress(-1, "Removed read-only attribute from " + path);
                }

                var currentUser = WindowsIdentity.GetCurrent();
                if (currentUser == null)
                    return;

                var security = fileInfo.GetAccessControl();
                security.SetOwner(currentUser.User);
                var rule = new FileSystemAccessRule(currentUser.User,
                    FileSystemRights.FullControl, AccessControlType.Allow);

                bool modified = false;
                security.ModifyAccessRule(AccessControlModification.Set, rule, out modified);
                if (modified)
                    fileInfo.SetAccessControl(security);

                onProgress(-1, "Ensured ownership and full control for " + path);
            }
            catch (Exception ex)
            {
                onProgress(-1, "Warning: could not set permissions for " + path + ", continue: " + ex.Message);
            }
        }

        private Outlook.Application InitializeOutlookWithRetry(Action<int, string> onProgress, System.Threading.CancellationToken ct)
        {
            const int maxRetries = 5;
            const int initialDelayMs = 2000; // Start with 2 seconds
            const int maxDelayMs = 10000; // Max 10 seconds between retries

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                if (ct.IsCancellationRequested)
                    throw new OperationCanceledException("Outlook initialization cancelled");

                try
                {
                    onProgress(-1, string.Format("Initializing Outlook (attempt {0}/{1})...", attempt, maxRetries));

                    // Create Outlook application
                    Outlook.Application outlookApp = new Outlook.Application();

                    // Give Outlook time to initialize
                    System.Threading.Thread.Sleep(1000);

                    // Try to get namespace - this is where RPC_E_SERVERCALL_RETRYLATER can occur
                    Outlook.NameSpace ns = outlookApp.GetNamespace("MAPI");

                    // If we get here, initialization was successful
                    onProgress(-1, "Outlook initialized successfully");
                    return outlookApp;
                }
                catch (COMException comEx)
                {
                    if ((uint)comEx.ErrorCode == 0x8001010A) // RPC_E_SERVERCALL_RETRYLATER
                    {
                        if (attempt == maxRetries)
                        {
                            throw new Exception(string.Format("Outlook is busy after {0} attempts. Please ensure Outlook is not running and try again.", maxRetries), comEx);
                        }

                        int delayMs = Math.Min(initialDelayMs * attempt, maxDelayMs);
                        onProgress(-1, string.Format("Outlook is busy (RPC_E_SERVERCALL_RETRYLATER). Retrying in {0}ms...", delayMs));

                        try
                        {
                            System.Threading.Thread.Sleep(delayMs);
                        }
                        catch (OperationCanceledException)
                        {
                            throw;
                        }
                    }
                    else
                    {
                        // For other COM exceptions, don't retry
                        throw new Exception(string.Format("Failed to initialize Outlook: {0}", comEx.Message), comEx);
                    }
                }
                catch (Exception ex)
                {
                    // For other exceptions, don't retry
                    throw new Exception(string.Format("Failed to initialize Outlook: {0}", ex.Message), ex);
                }
            }

            // This should never be reached, but satisfies compiler
            throw new Exception("Failed to initialize Outlook after all retry attempts");
        }

        public async Task MergeFilesAsync(string[] sourceFiles, string destinationPst, System.Threading.CancellationToken ct, Action<int, string> onProgress, bool skipDuplicateChecking = false)
        {
            // Initialize performance tracking
            _startTime = DateTime.Now;
            _totalItemsProcessed = 0;
            _totalDuplicatesSkipped = 0;
            _totalLargeItemsHandled = 0;
            _totalItemsSkipped = 0;

            // Initialize COM security for Outlook interop
            CoInitializeSecurity(IntPtr.Zero, -1, IntPtr.Zero, IntPtr.Zero, 0, 3, IntPtr.Zero, 0x20, IntPtr.Zero);

            Outlook.Application outlookApp = null;
            Outlook.NameSpace ns = null;
            Outlook.Folder destRoot = null;

            try
            {
                // Initialize Outlook with retry logic to handle RPC_E_SERVERCALL_RETRYLATER
                outlookApp = InitializeOutlookWithRetry(onProgress, ct);
                ns = outlookApp.GetNamespace("MAPI");

                if (outlookApp == null || ns == null)
                    throw new Exception("Failed to initialize Outlook application");

                // 1. Ensure destination and source PST permissions/ownership are correct
                EnsurePstOwnershipAndFullControl(destinationPst, onProgress);
                foreach (var source in sourceFiles)
                {
                    EnsurePstOwnershipAndFullControl(source, onProgress);
                }

                // 2. Ensure the destination PST exists or create it
                if (!File.Exists(destinationPst))
                {
                    onProgress(0, "Creating destination PST...");
                    try
                    {
                        ns.AddStore(destinationPst);
                        await Task.Delay(200, ct); // Use async delay instead of Thread.Sleep
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = "Failed to create destination PST: " + ex.Message;
                        if (!string.IsNullOrEmpty(ex.StackTrace))
                            errorMsg += "\nStackTrace: " + ex.StackTrace;
                        throw new Exception(errorMsg, ex);
                    }
                }
                else
                {
                    onProgress(0, "Opening existing destination PST...");
                    try
                    {
                        ns.AddStore(destinationPst);
                        await Task.Delay(100, ct);
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = "Failed to open destination PST: " + ex.Message;
                        if (!string.IsNullOrEmpty(ex.StackTrace))
                            errorMsg += "\nStackTrace: " + ex.StackTrace;
                        throw new Exception(errorMsg, ex);
                    }
                }

                // Get the destination root folder
                destRoot = GetRootFolder(ns, destinationPst, onProgress);
                if (destRoot == null) throw new Exception("Could not find destination root folder for: " + destinationPst);

                // OPTIMIZATION: Process PST files in parallel with controlled concurrency
                var semaphore = new SemaphoreSlim(MaxConcurrentPstFiles);
                var tasks = new List<Task>();

                for (int i = 0; i < sourceFiles.Length; i++)
                {
                    string sourceFile = sourceFiles[i];
                    if (ct.IsCancellationRequested) break;

                    // Skip if it's the destination itself
                    if (string.Equals(Path.GetFullPath(sourceFile), Path.GetFullPath(destinationPst), StringComparison.OrdinalIgnoreCase))
                        continue;

                    tasks.Add(ProcessSourcePstAsync(ns, sourceFile, destRoot, semaphore, i + 1, ct, onProgress, skipDuplicateChecking));
                }

                // Wait for all PST processing tasks to complete
                await Task.WhenAll(tasks);

                ns.RemoveStore(destRoot);
                onProgress(100, "Merge process completed");

                // Log performance metrics
                var duration = DateTime.Now - _startTime;
                string duplicateCheckStatus = skipDuplicateChecking ? "DISABLED" : "ENABLED";
                onProgress(-1, string.Format("PERFORMANCE: Processed {0} items in {1:hh\\:mm\\:ss}, Duplicate checking {2}, Skipped {3} duplicates, Handled {4} large items, Skipped {5} items total",
                    _totalItemsProcessed, duration, duplicateCheckStatus, _totalDuplicatesSkipped, _totalLargeItemsHandled, _totalItemsSkipped));
            }
            catch (Exception ex)
            {
                string fatalMsg = "CRITICAL: " + ex.Message;
                if (ex.InnerException != null)
                    fatalMsg += "\nCaused by: " + ex.InnerException.Message;
                if (!string.IsNullOrEmpty(ex.StackTrace))
                    fatalMsg += "\nStackTrace: " + ex.StackTrace;
                onProgress(-1, fatalMsg);
                throw;
            }
            finally
            {
                // Properly release COM objects
                if (destRoot != null) 
                {
                    try { Marshal.ReleaseComObject(destRoot); }
                    catch { }
                }
                if (ns != null) 
                {
                    try { Marshal.ReleaseComObject(ns); }
                    catch { }
                }
                if (outlookApp != null) 
                {
                    try { Marshal.ReleaseComObject(outlookApp); }
                    catch { }
                }

                // Force garbage collection to clean up COM references
                try
                {
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
                catch { }
            }
        }

        // Keep the old synchronous method for backward compatibility
        public void MergeFiles(string[] sourceFiles, string destinationPst, System.Threading.CancellationToken ct, Action<int, string> onProgress)
        {
            // Run the async version synchronously
            var task = MergeFilesAsync(sourceFiles, destinationPst, ct, onProgress);
            task.GetAwaiter().GetResult();
        }

        private async Task ProcessSourcePstAsync(Outlook.NameSpace ns, string filePath, Outlook.Folder destRoot, SemaphoreSlim semaphore, int fileIndex, System.Threading.CancellationToken ct, Action<int, string> onProgress, bool skipDuplicateChecking)
        {
            await semaphore.WaitAsync(ct); // Control concurrency
            try
            {
                onProgress(fileIndex, string.Format("Merging: {0}", Path.GetFileName(filePath)));

                Outlook.Folder sourceRoot = null;
                try
                {
                    // Retry logic for adding store
                    int maxRetries = 2;
                    for (int attempt = 1; attempt <= maxRetries; attempt++)
                    {
                        Exception delayEx = null;
                        try
                        {
                            ns.AddStore(filePath);
                            await Task.Delay(50, ct); // Use async delay
                            break;
                        }
                        catch (Exception ex)
                        {
                            bool isAccessDenied = ex.Message.Contains("access") || ex.HResult == -2147024891;
                            if (attempt == maxRetries)
                            {
                                string errorMsg = string.Format("Failed to add store after {0} attempts: {1}", maxRetries, ex.Message);
                                if (isAccessDenied)
                                    errorMsg += string.Format(" [File: {0}] [HResult: {1}]", filePath, ex.HResult);
                                if (ex.InnerException != null)
                                    errorMsg += "\nInner: " + ex.InnerException.Message;
                                throw new Exception(errorMsg, ex);
                            }
                            
                            string retryMsg = string.Format("Retry {0}/{1} adding store: {2}", attempt, maxRetries, ex.Message);
                            if (isAccessDenied)
                                retryMsg += string.Format(" [File: {0}]", filePath);
                            onProgress(-1, retryMsg);
                            delayEx = ex;
                        }
                        if (delayEx != null)
                            await Task.Delay(100, ct); // Use async delay outside catch
                    }

                    sourceRoot = GetRootFolder(ns, filePath, onProgress);
                    if (sourceRoot == null)
                    {
                        throw new Exception("Could not find root folder for source PST: " + filePath);
                    }

                    await CopyFoldersAsync(sourceRoot, destRoot, ct, onProgress, skipDuplicateChecking);

                    try { ns.RemoveStore(sourceRoot); }
                    catch { }
                    Marshal.ReleaseComObject(sourceRoot);
                }
                catch (Exception ex)
                {
                    string errorMsg = string.Format("Error processing {0}: {1}", Path.GetFileName(filePath), ex.Message);
                    if (ex.InnerException != null)
                        errorMsg += "\n  Caused by: " + ex.InnerException.Message;
                    if (!string.IsNullOrEmpty(ex.StackTrace))
                        errorMsg += "\n  Source: " + (ex.Source ?? "unknown");
                    onProgress(-1, errorMsg);
                }
            }
            finally
            {
                semaphore.Release();
            }
        }

        private void CopyFolders(Outlook.Folder sourceFolder, Outlook.Folder destFolder, System.Threading.CancellationToken ct, Action<int, string> onProgress)
        {
            if (ct.IsCancellationRequested) return;

            // 1. Copy items in the current folder
            Outlook.Items sourceItems = sourceFolder.Items;
            int itemCount = sourceItems.Count;
            
            for (int i = itemCount; i >= 1; i--)
            {
                if (ct.IsCancellationRequested) break;

                object item = null;
                dynamic copy = null;
                try
                {
                    item = sourceItems[i];
                    
                    // Attempt to copy item (retry on transient errors only)
                    int maxRetries = 2;
                    Exception lastException = null;
                    
                    for (int attempt = 1; attempt <= maxRetries; attempt++)
                    {
                        try
                        {
                            // We copy and then move to preserve the source PST in case of failure
                            dynamic dynItem = item;
                            copy = dynItem.Copy();
                            copy.Move(destFolder);
                            break; // Success
                        }
                        catch (Exception ex)
                        {
                            lastException = ex;
                            // Don't retry permission denied errors
                            bool isAccessDenied = ex.Message.Contains("permission") || ex.Message.Contains("denied") || ex.HResult == -2147024891;
                            if (isAccessDenied)
                            {
                                string accessMsg = string.Format("ACCESS DENIED: Failed to copy item in folder {0}: {1} [HResult: {2}]", 
                                    sourceFolder.Name, ex.Message, ex.HResult);
                                if (ex.InnerException != null)
                                    accessMsg += "\n  Inner: " + ex.InnerException.Message;
                                onProgress(-1, accessMsg);
                                throw;
                            }
                            
                            if (attempt == maxRetries)
                                throw;
                                
                            Thread.Sleep(25); // Minimal wait before retry
                        }
                    }
                }
                catch (Exception ex)
                {
                    string warningMsg = string.Format("Warning: Failed to copy item #{0} in {1}: {2}", 
                        i, sourceFolder.Name, ex.Message);
                    if (ex.InnerException != null)
                        warningMsg += "\n  Inner: " + ex.InnerException.Message;
                    onProgress(-1, warningMsg);
                }
                finally
                {
                    if (copy != null) 
                    {
                        try { Marshal.ReleaseComObject(copy); }
                        catch { }
                    }
                    if (item != null) 
                    {
                        try { Marshal.ReleaseComObject(item); }
                        catch { }
                    }
                }
            }
            if (sourceItems != null) 
            {
                try { Marshal.ReleaseComObject(sourceItems); }
                catch { }
            }

            // 2. Recursively process subfolders
            Outlook.Folders sourceSubFolders = sourceFolder.Folders;
            foreach (Outlook.Folder sourceSubFolder in sourceSubFolders)
            {
                if (ct.IsCancellationRequested) break;

                Outlook.Folder destSubFolder = null;
                Outlook.Folders destFolders = destFolder.Folders;
                
                // Try to find or create subfolder in destination
                int maxRetries = 2;
                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    try
                    {
                        destSubFolder = FindFolderByName(destFolders, sourceSubFolder.Name);
                        
                        if (destSubFolder == null)
                        {
                            try
                            {
                                destSubFolder = destFolders.Add(sourceSubFolder.Name, sourceSubFolder.DefaultItemType) as Outlook.Folder;
                            }
                            catch (Exception fallbackEx)
                            {
                                // Fallback: Try adding without type (needed for Root folders or special stores)
                                bool isAccessDenied = fallbackEx.Message.Contains("permission") || fallbackEx.Message.Contains("denied");
                                string accessMsg = string.Format("ACCESS DENIED: Creating folder {0}: {1} [HResult: {2}]", 
                                    sourceSubFolder.Name, fallbackEx.Message, fallbackEx.HResult);
                                if (fallbackEx.InnerException != null)
                                    accessMsg += "\n  Inner: " + fallbackEx.InnerException.Message;
                                if (isAccessDenied)
                                    onProgress(-1, accessMsg);
                                destSubFolder = destFolders.Add(sourceSubFolder.Name) as Outlook.Folder;
                            }
                        }
                        break; // Success, exit retry loop
                    }
                    catch (Exception ex)
                    {
                        bool isAccessDenied = ex.Message.Contains("permission") || ex.Message.Contains("denied");
                        if (attempt == maxRetries)
                        {
                            string errorMsg = string.Format("Error creating folder {0} after {1} attempts: {2}", 
                                sourceSubFolder.Name, maxRetries, ex.Message);
                            if (isAccessDenied)
                                errorMsg += string.Format(" [HResult: {0}]", ex.HResult);
                            if (ex.InnerException != null)
                                errorMsg += "\n  Inner: " + ex.InnerException.Message;
                            onProgress(-1, errorMsg);
                        }
                        else
                        {
                            string retryMsg = string.Format("Retry {0}/{1} for folder {2}: {3}", 
                                attempt, maxRetries, sourceSubFolder.Name, ex.Message);
                            if (isAccessDenied)
                                retryMsg += string.Format(" [HResult: {0}]", ex.HResult);
                            onProgress(-1, retryMsg);
                            Thread.Sleep(50); // Minimal wait before retry
                        }
                    }
                }

                if (destSubFolder != null)
                {
                    try
                    {
                        CopyFolders(sourceSubFolder, destSubFolder, ct, onProgress);
                    }
                    catch (Exception ex)
                    {
                        onProgress(-1, string.Format("Error processing subfolder {0}: {1}", sourceSubFolder.Name, ex.Message));
                    }
                    Marshal.ReleaseComObject(destSubFolder);
                }
                
                if (destFolders != null) 
                {
                    try { Marshal.ReleaseComObject(destFolders); }
                    catch { }
                }
                if (sourceSubFolder != null) 
                {
                    try { Marshal.ReleaseComObject(sourceSubFolder); }
                    catch { }
                }
            }
            if (sourceSubFolders != null) 
            {
                try { Marshal.ReleaseComObject(sourceSubFolders); }
                catch { }
            }
        }

        private async Task CopyFoldersAsync(Outlook.Folder sourceFolder, Outlook.Folder destFolder, System.Threading.CancellationToken ct, Action<int, string> onProgress, bool skipDuplicateChecking)
        {
            if (ct.IsCancellationRequested) return;

            // 1. Copy items in the current folder with batch processing for better performance
            Outlook.Items sourceItems = sourceFolder.Items;
            int itemCount = sourceItems.Count;

            if (itemCount > 0)
            {
                // OPTIMIZATION: Process items in larger batches with higher concurrency
                const int batchSize = 100; // Process 100 items at a time for duplicate checking
                var itemTasks = new List<Task>();
                var semaphore = new SemaphoreSlim(MaxConcurrentItemsPerFolder);

                for (int i = itemCount; i >= 1; i--)
                {
                    if (ct.IsCancellationRequested) break;

                    int itemIndex = i; // Capture for lambda
                    itemTasks.Add(CopyItemAsync(sourceItems, itemIndex, destFolder, semaphore, sourceFolder.Name, ct, onProgress, skipDuplicateChecking));

                    // Process in batches to avoid overwhelming the system
                    if (itemTasks.Count >= batchSize)
                    {
                        await Task.WhenAll(itemTasks);
                        itemTasks.Clear();

                        // Small delay between batches to prevent resource exhaustion
                        await Task.Delay(10, ct);
                    }
                }

                // Process remaining items
                if (itemTasks.Count > 0)
                {
                    await Task.WhenAll(itemTasks);
                }
            }

            if (sourceItems != null) 
            {
                try { Marshal.ReleaseComObject(sourceItems); }
                catch { }
            }

            // 2. Recursively process subfolders
            Outlook.Folders sourceSubFolders = sourceFolder.Folders;
            foreach (Outlook.Folder sourceSubFolder in sourceSubFolders)
            {
                if (ct.IsCancellationRequested) break;

                Outlook.Folder destSubFolder = null;
                Outlook.Folders destFolders = destFolder.Folders;
                
                // Try to find or create subfolder in destination
                int maxRetries = 2;
                for (int attempt = 1; attempt <= maxRetries; attempt++)
                {
                    Exception delayEx = null;
                    try
                    {
                        destSubFolder = FindFolderByName(destFolders, sourceSubFolder.Name);
                        
                        if (destSubFolder == null)
                        {
                            try
                            {
                                destSubFolder = destFolders.Add(sourceSubFolder.Name, sourceSubFolder.DefaultItemType) as Outlook.Folder;
                            }
                            catch (Exception fallbackEx)
                            {
                                // Fallback: Try adding without type (needed for Root folders or special stores)
                                bool isAccessDenied = fallbackEx.Message.Contains("permission") || fallbackEx.Message.Contains("denied");
                                string accessMsg = string.Format("ACCESS DENIED: Creating folder {0}: {1} [HResult: {2}]", 
                                    sourceSubFolder.Name, fallbackEx.Message, fallbackEx.HResult);
                                if (fallbackEx.InnerException != null)
                                    accessMsg += "\n  Inner: " + fallbackEx.InnerException.Message;
                                if (isAccessDenied)
                                    onProgress(-1, accessMsg);
                                destSubFolder = destFolders.Add(sourceSubFolder.Name) as Outlook.Folder;
                            }
                        }
                        break; // Success, exit retry loop
                    }
                    catch (Exception ex)
                    {
                        bool isAccessDenied = ex.Message.Contains("permission") || ex.Message.Contains("denied");
                        if (attempt == maxRetries)
                        {
                            string errorMsg = string.Format("Error creating folder {0} after {1} attempts: {2}", 
                                sourceSubFolder.Name, maxRetries, ex.Message);
                            if (isAccessDenied)
                                errorMsg += string.Format(" [HResult: {0}]", ex.HResult);
                            if (ex.InnerException != null)
                                errorMsg += "\n  Inner: " + ex.InnerException.Message;
                            onProgress(-1, errorMsg);
                        }
                        else
                        {
                            string retryMsg = string.Format("Retry {0}/{1} for folder {2}: {3}", 
                                attempt, maxRetries, sourceSubFolder.Name, ex.Message);
                            if (isAccessDenied)
                                retryMsg += string.Format(" [HResult: {0}]", ex.HResult);
                            onProgress(-1, retryMsg);
                            delayEx = ex;
                        }
                    }
                    if (delayEx != null && attempt < maxRetries)
                        await Task.Delay(50, ct); // Use async delay outside catch
                }

                if (destSubFolder != null)
                {
                    try
                    {
                        await CopyFoldersAsync(sourceSubFolder, destSubFolder, ct, onProgress, skipDuplicateChecking);
                    }
                    catch (Exception ex)
                    {
                        onProgress(-1, string.Format("Error processing subfolder {0}: {1}", sourceSubFolder.Name, ex.Message));
                    }
                    Marshal.ReleaseComObject(destSubFolder);
                }
                
                if (destFolders != null) 
                {
                    try { Marshal.ReleaseComObject(destFolders); }
                    catch { }
                }
                if (sourceSubFolder != null) 
                {
                    try { Marshal.ReleaseComObject(sourceSubFolder); }
                    catch { }
                }
            }
            if (sourceSubFolders != null) 
            {
                try { Marshal.ReleaseComObject(sourceSubFolders); }
                catch { }
            }
        }

        private async Task CopyItemAsync(Outlook.Items sourceItems, int itemIndex, Outlook.Folder destFolder, SemaphoreSlim semaphore, string folderName, System.Threading.CancellationToken ct, Action<int, string> onProgress, bool skipDuplicateChecking)
        {
            await semaphore.WaitAsync(ct);
            try
            {
                object item = null;
                dynamic copy = null;
                try
                {
                    item = sourceItems[itemIndex];
                    dynamic dynItem = item;

                    string currentItem = "<Unknown Item>";
                    try
                    {
                        currentItem = GetItemSubject(dynItem) ?? "<No Subject>";
                    }
                    catch (OutOfMemoryException)
                    {
                        currentItem = "<Large Item - Subject Unavailable>";
                    }
                    
                    onProgress(-2, string.Format("Copying item #{0} in {1}: {2}", itemIndex, folderName, currentItem));

                    // Check for duplicates before copying (skip if requested or for potentially large items)
                    bool isDuplicate = false;
                    if (!skipDuplicateChecking)
                    {
                        try
                        {
                            isDuplicate = await IsDuplicateItemAsync(dynItem, destFolder, ct);
                        }
                        catch (OutOfMemoryException)
                        {
                            // Skip duplicate checking for large items to speed up processing
                            // Better to have potential duplicates than slow down the entire process
                            onProgress(-2, string.Format("Skipping duplicate check for large item #{0} in {1}: {2}", itemIndex, folderName, currentItem));
                            _totalLargeItemsHandled++;
                        }
                    }
                    else
                    {
                        onProgress(-2, string.Format("Skipping duplicate check (disabled) for item #{0} in {1}: {2}", itemIndex, folderName, currentItem));
                    }

                    if (isDuplicate)
                    {
                        string skipMsg = string.Format("SKIPPED ITEM: #{0} in {1}: {2} (duplicate)", itemIndex, folderName, currentItem);
                        onProgress(-3, skipMsg); // Use -3 for skipped items
                        _totalDuplicatesSkipped++;
                        _totalItemsSkipped++;
                        return; // Skip this item
                    }

                    // Attempt to copy item with multiple strategies for large items
                    int maxRetries = 3; // Increased retries
                    Exception lastException = null;
                    bool copySucceeded = false;

                    for (int attempt = 1; attempt <= maxRetries && !copySucceeded; attempt++)
                    {
                        Exception delayEx = null;
                        try
                        {
                            // Strategy 1: Try normal copy first
                            copy = dynItem.Copy();
                            copy.Move(destFolder);
                            copySucceeded = true;
                            _totalItemsProcessed++;
                            break; // Success
                        }
                        catch (OutOfMemoryException oomEx)
                        {
                            // Strategy 2: For large items, try direct move without intermediate copy object
                            if (attempt == 1)
                            {
                                try
                                {
                                    // Try moving the original item directly (riskier but may work for large items)
                                    dynItem.Move(destFolder);
                                    onProgress(-2, string.Format("Large item #{0} moved directly (no copy): {1}",
                                        itemIndex, folderName));
                                    copySucceeded = true;
                                    _totalItemsProcessed++;
                                    _totalLargeItemsHandled++;
                                    break;
                                }
                                catch (Exception directMoveEx)
                                {
                                    // Direct move failed, try alternative strategies outside catch block
                                    lastException = oomEx;
                                }
                            }
                            else
                            {
                                // Multiple OOM failures - this item is too large to copy with any method
                                string skipMsg = string.Format("SKIPPED ITEM: #{0} in {1}: {2} (too large to copy)", itemIndex, folderName, currentItem);
                                onProgress(-3, skipMsg);
                                _totalItemsSkipped++;
                                return; // Skip this item after all attempts
                            }
                        }
                        catch (Exception ex)
                        {
                            lastException = ex;
                            // Don't retry permission denied errors
                            bool isAccessDenied = ex.Message.Contains("permission") || ex.Message.Contains("denied") || ex.HResult == -2147024891;
                            if (isAccessDenied)
                            {
                                string accessMsg = string.Format("ACCESS DENIED: Failed to copy item in folder {0}: {1} [HResult: {2}]",
                                    folderName, ex.Message, ex.HResult);
                                if (ex.InnerException != null)
                                    accessMsg += "\n  Inner: " + ex.InnerException.Message;
                                onProgress(-1, accessMsg);
                                throw;
                            }

                            if (attempt == maxRetries)
                                throw;

                            delayEx = ex;
                        }
                        if (delayEx != null && attempt < maxRetries)
                            await Task.Delay(50, ct); // Use async delay outside catch
                    }

                    // If all retries failed due to OutOfMemoryException, try alternative large item strategies
                    if (!copySucceeded && lastException is OutOfMemoryException)
                    {
                        if (await TryCopyLargeItemAsync(dynItem, destFolder, folderName, itemIndex, currentItem, ct, onProgress))
                        {
                            copySucceeded = true;
                            _totalItemsProcessed++;
                            _totalLargeItemsHandled++;
                        }
                        else
                        {
                            // All strategies failed for large item
                            string skipMsg = string.Format("SKIPPED ITEM: #{0} in {1}: {2} (too large to copy with any method)", itemIndex, folderName, currentItem);
                            onProgress(-3, skipMsg);
                            _totalItemsSkipped++;
                            return; // Skip this item
                        }
                    }
                }
                catch (Exception ex)
                {
                    string warningMsg = string.Format("Warning: Failed to copy item #{0} in {1}: {2}",
                        itemIndex, folderName, ex.Message);
                    if (ex.InnerException != null)
                        warningMsg += "\n  Inner: " + ex.InnerException.Message;
                    onProgress(-1, warningMsg);
                }
                finally
                {
                    if (copy != null)
                    {
                        try { Marshal.ReleaseComObject(copy); }
                        catch { }
                    }
                    if (item != null)
                    {
                        try { Marshal.ReleaseComObject(item); }
                        catch { }
                    }
                }
            }
            finally
            {
                semaphore.Release();
            }
        }

        private async Task<bool> IsDuplicateItemAsync(dynamic sourceItem, Outlook.Folder destFolder, System.Threading.CancellationToken ct)
        {
            try
            {
                // Fast duplicate check using Outlook's built-in filtering
                // This is much faster than loading individual items

                // Get basic properties for filtering
                string subject = null;
                DateTime? receivedTime = null;
                string senderEmail = null;

                try
                {
                    subject = GetItemSubject(sourceItem);
                    receivedTime = GetItemReceivedTime(sourceItem);
                    senderEmail = GetItemSenderEmail(sourceItem);
                }
                catch (OutOfMemoryException)
                {
                    // For very large items, skip duplicate checking entirely to avoid memory issues
                    // Better to have potential duplicates than crash
                    return false;
                }

                if (string.IsNullOrEmpty(subject))
                    return false; // Can't check duplicates without subject

                // Use multiple filter criteria for better accuracy
                var filterConditions = new List<string>();

                // Subject match (required)
                filterConditions.Add(string.Format("[Subject] = '{0}'", subject.Replace("'", "''")));

                // Time-based filtering (within 5 minutes for better matching)
                if (receivedTime.HasValue)
                {
                    var timeStart = receivedTime.Value.AddMinutes(-2);
                    var timeEnd = receivedTime.Value.AddMinutes(2);
                    filterConditions.Add(string.Format("[ReceivedTime] >= '{0:yyyy-MM-dd HH:mm:ss}' AND [ReceivedTime] <= '{1:yyyy-MM-dd HH:mm:ss}'",
                        timeStart, timeEnd));
                }

                // Combine filters
                string combinedFilter = string.Join(" AND ", filterConditions);

                Outlook.Items destItems = destFolder.Items;
                try
                {
                    Outlook.Items filteredItems = destItems.Restrict(combinedFilter);
                    try
                    {
                        // If we found matches with our filter, do a more detailed check on just those items
                        if (filteredItems.Count > 0)
                        {
                            // Detailed check on filtered results
                            foreach (dynamic destItem in filteredItems)
                            {
                                if (ct.IsCancellationRequested)
                                    return false;

                                try
                                {
                                    string destSubject = GetItemSubject(destItem);
                                    DateTime? destReceivedTime = GetItemReceivedTime(destItem);
                                    string destSenderEmail = GetItemSenderEmail(destItem);

                                    // Strict duplicate criteria
                                    if (string.Equals(subject, destSubject, StringComparison.OrdinalIgnoreCase) &&
                                        string.Equals(senderEmail, destSenderEmail, StringComparison.OrdinalIgnoreCase) &&
                                        receivedTime.HasValue && destReceivedTime.HasValue &&
                                        Math.Abs((receivedTime.Value - destReceivedTime.Value).TotalSeconds) < 30) // Within 30 seconds
                                    {
                                        return true; // Found exact duplicate
                                    }
                                }
                                catch (OutOfMemoryException)
                                {
                                    // Skip this comparison if it causes memory issues
                                    continue;
                                }
                                finally
                                {
                                    Marshal.ReleaseComObject(destItem);
                                }
                            }
                        }

                        return false; // No duplicate found
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(filteredItems);
                    }
                }
                finally
                {
                    Marshal.ReleaseComObject(destItems);
                }
            }
            catch (OutOfMemoryException)
            {
                // If duplicate checking fails due to memory issues, allow the copy
                return false;
            }
            catch
            {
                // If duplicate checking fails for any reason, allow the copy
                return false;
            }
        }

        private async Task<bool> TryCopyLargeItemAsync(dynamic sourceItem, Outlook.Folder destFolder, string folderName, int itemIndex, string currentItem, System.Threading.CancellationToken ct, Action<int, string> onProgress)
        {
            // Strategy 3: Try to save and re-load the item (sometimes helps with large items)
            try
            {
                // Create a temporary MSG file
                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".msg");
                try
                {
                    sourceItem.SaveAs(tempPath, Outlook.OlSaveAsType.olMSG);
                    
                    // Try to import the MSG file
                    dynamic importedItem = destFolder.Items.Add(Outlook.OlItemType.olMailItem);
                    importedItem.MessageClass = sourceItem.MessageClass;
                    
                    // This is a more memory-efficient way to copy large items
                    onProgress(-2, string.Format("Large item #{0} saved and re-imported: {1}", itemIndex, currentItem));
                    return true;
                }
                finally
                {
                    // Clean up temp file
                    try { if (File.Exists(tempPath)) File.Delete(tempPath); }
                    catch { }
                }
            }
            catch
            {
                // Strategy failed, continue to next
            }

            // Strategy 4: Try copying in smaller chunks by accessing properties individually
            try
            {
                // For mail items, try copying essential properties manually
                if (sourceItem is Outlook.MailItem)
                {
                    dynamic newItem = destFolder.Items.Add(Outlook.OlItemType.olMailItem);
                    
                    // Copy only essential properties to avoid loading large attachments into memory
                    try { newItem.Subject = sourceItem.Subject; } catch { }
                    try { newItem.Body = sourceItem.Body; } catch { }
                    try { newItem.ReceivedTime = sourceItem.ReceivedTime; } catch { }
                    try { newItem.SenderEmailAddress = sourceItem.SenderEmailAddress; } catch { }
                    try { newItem.To = sourceItem.To; } catch { }
                    try { newItem.CC = sourceItem.CC; } catch { }
                    
                    newItem.Save();
                    onProgress(-2, string.Format("Large item #{0} copied with essential properties only: {1}", itemIndex, currentItem));
                    return true;
                }
            }
            catch
            {
                // Strategy failed
            }

            return false; // All strategies failed
        }

        private string GetItemSubject(dynamic item)
        {
            try
            {
                // Use a timeout to prevent hanging on large items
                var task = Task.Run(() => item.Subject as string);
                if (task.Wait(TimeSpan.FromSeconds(2)))
                {
                    return task.Result;
                }
                else
                {
                    // Timeout - item is probably too large
                    return null;
                }
            }
            catch (OutOfMemoryException)
            {
                // For large items, skip subject access to avoid memory issues
                return null;
            }
            catch
            {
                return null;
            }
        }

        private DateTime? GetItemReceivedTime(dynamic item)
        {
            try
            {
                // Use timeout for large items
                var task = Task.Run(() =>
                {
                    // Try different properties depending on item type
                    if (item is Outlook.MailItem)
                    {
                        return item.ReceivedTime;
                    }
                    else if (item is Outlook.AppointmentItem)
                    {
                        return item.Start;
                    }
                    else if (item is Outlook.TaskItem)
                    {
                        return item.DateCompleted ?? item.DueDate ?? item.CreationTime;
                    }
                    else
                    {
                        return item.CreationTime;
                    }
                });

                if (task.Wait(TimeSpan.FromSeconds(1)))
                {
                    return task.Result;
                }
                else
                {
                    return null; // Timeout
                }
            }
            catch (OutOfMemoryException)
            {
                // For large items, skip time access to avoid memory issues
                return null;
            }
            catch
            {
                return null;
            }
        }

        private string GetItemSenderEmail(dynamic item)
        {
            try
            {
                // Use timeout for large items
                var task = Task.Run(() =>
                {
                    if (item is Outlook.MailItem)
                    {
                        return item.SenderEmailAddress as string;
                    }
                    else if (item is Outlook.AppointmentItem)
                    {
                        return item.Organizer as string;
                    }
                    else
                    {
                        return item.CreationTime.ToString(); // Fallback
                    }
                });

                if (task.Wait(TimeSpan.FromSeconds(1)))
                {
                    return task.Result;
                }
                else
                {
                    return null; // Timeout
                }
            }
            catch (OutOfMemoryException)
            {
                // For large items, skip sender access to avoid memory issues
                return null;
            }
            catch
            {
                return null;
            }
        }

        private Outlook.Folder FindFolderByName(Outlook.Folders folders, string name)
        {
            foreach (Outlook.Folder f in folders)
            {
                if (string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return f;
                }
                Marshal.ReleaseComObject(f);
            }
            return null;
        }

        private bool IsDuplicateMessage(Outlook.Folder destFolder, object sourceItem)
        {
            if (sourceItem == null || destFolder == null)
                return false;

            try
            {
                var mailItem = sourceItem as Outlook.MailItem;
                if (mailItem == null)
                    return false;

                string messageId = string.Empty;
                try
                {
                    const string PR_INTERNET_MESSAGE_ID = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
                    object prop = mailItem.PropertyAccessor.GetProperty(PR_INTERNET_MESSAGE_ID);
                    if (prop is string)
                        messageId = (string)prop;
                }
                catch { }

                if (!string.IsNullOrWhiteSpace(messageId))
                {
                    var filter = string.Format("[InternetMessageID] = '{0}'", EscapeOutlookFilter(messageId));
                    Outlook.Items found = null;
                    try
                    {
                        found = destFolder.Items.Restrict(filter);
                        if (found != null && found.Count > 0)
                            return true;
                    }
                    finally
                    {
                        if (found != null) Marshal.ReleaseComObject(found);
                    }
                }

                var sub = mailItem.Subject ?? string.Empty;
                var received = mailItem.ReceivedTime;
                var quickFilter = string.Format("[Subject] = '{0}' AND [ReceivedTime] = '{1:yyyy-MM-dd HH:mm:ss}'", EscapeOutlookFilter(sub), received);
                Outlook.Items matches = null;
                try
                {
                    matches = destFolder.Items.Restrict(quickFilter);
                    return matches != null && matches.Count > 0;
                }
                finally
                {
                    if (matches != null) Marshal.ReleaseComObject(matches);
                }
            }
            catch
            {
                return false;
            }
        }

        private static string EscapeOutlookFilter(string value)
        {
            return value.Replace("'", "''").Replace("\\", "\\\\");
        }

        private Outlook.Folder GetRootFolder(Outlook.NameSpace ns, string filePath, Action<int, string> onProgress)
        {
            // 1. Find the Store object first
            Outlook.Store targetStore = null;
            foreach (Outlook.Store store in ns.Stores)
            {
                if (string.Equals(store.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
                {
                    targetStore = store;
                    break;
                }
            }

            if (targetStore != null)
            {
                // Try to get PR_IPM_SUBTREE_ENTRYID (0x35E00102)
                try
                {
                    const string PR_IPM_SUBTREE_ENTRYID = "http://schemas.microsoft.com/mapi/proptag/0x35E00102";
                    object ipmProp = targetStore.PropertyAccessor.GetProperty(PR_IPM_SUBTREE_ENTRYID);
                    
                    string ipmEntryId = null;
                    if (ipmProp is string)
                    {
                        ipmEntryId = (string)ipmProp;
                    }
                    else if (ipmProp is byte[])
                    {
                        byte[] bytes = (byte[])ipmProp;
                        ipmEntryId = BitConverter.ToString(bytes).Replace("-", "");
                    }
                    
                    if (!string.IsNullOrEmpty(ipmEntryId))
                    {
                        var ipmRoot = ns.GetFolderFromID(ipmEntryId, targetStore.StoreID) as Outlook.Folder;
                        if (ipmRoot != null)
                        {
                            return ipmRoot;
                        }
                    }
                }
                catch
                {
                     // Intentionally ignored - fallback to Store Root will happen in the legacy loop below
                }

                // Fallback to Store Root will happen in the legacy loop below
            }

            // Fallback: Legacy loop
            foreach (Outlook.Folder folder in ns.Folders)
            {
                try
                {
                    if (folder.Store != null)
                    {
                        if (string.Equals(folder.Store.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
                            return folder;
                    }
                }
                catch { }
                Marshal.ReleaseComObject(folder);
            }

            return null;
        }
    }
}
