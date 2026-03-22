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
        private const int MaxConcurrentItemsPerFolder = 5; // Copy up to 5 items concurrently per folder

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

        public async Task MergeFilesAsync(string[] sourceFiles, string destinationPst, System.Threading.CancellationToken ct, Action<int, string> onProgress)
        {
            // Initialize COM security for Outlook interop
            CoInitializeSecurity(IntPtr.Zero, -1, IntPtr.Zero, IntPtr.Zero, 0, 3, IntPtr.Zero, 0x20, IntPtr.Zero);

            Outlook.Application outlookApp = null;
            Outlook.NameSpace ns = null;
            Outlook.Folder destRoot = null;

            try
            {
                // Create Outlook application (assuming we're already on STA thread)
                outlookApp = new Outlook.Application();
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

                    tasks.Add(ProcessSourcePstAsync(ns, sourceFile, destRoot, semaphore, i + 1, ct, onProgress));
                }

                // Wait for all PST processing tasks to complete
                await Task.WhenAll(tasks);

                ns.RemoveStore(destRoot);
                onProgress(100, "Merge process completed");
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

        private async Task ProcessSourcePstAsync(Outlook.NameSpace ns, string filePath, Outlook.Folder destRoot, SemaphoreSlim semaphore, int fileIndex, System.Threading.CancellationToken ct, Action<int, string> onProgress)
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

                    await CopyFoldersAsync(sourceRoot, destRoot, ct, onProgress);

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

        private async Task CopyFoldersAsync(Outlook.Folder sourceFolder, Outlook.Folder destFolder, System.Threading.CancellationToken ct, Action<int, string> onProgress)
        {
            if (ct.IsCancellationRequested) return;

            // 1. Copy items in the current folder with parallel processing
            Outlook.Items sourceItems = sourceFolder.Items;
            int itemCount = sourceItems.Count;

            if (itemCount > 0)
            {
                // OPTIMIZATION: Process items in parallel batches
                var itemTasks = new List<Task>();
                var semaphore = new SemaphoreSlim(MaxConcurrentItemsPerFolder);

                for (int i = itemCount; i >= 1; i--)
                {
                    if (ct.IsCancellationRequested) break;

                    int itemIndex = i; // Capture for lambda
                    itemTasks.Add(CopyItemAsync(sourceItems, itemIndex, destFolder, semaphore, sourceFolder.Name, ct, onProgress));
                }

                // Wait for all items in this folder to be copied
                await Task.WhenAll(itemTasks);
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
                        await CopyFoldersAsync(sourceSubFolder, destSubFolder, ct, onProgress);
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

        private async Task CopyItemAsync(Outlook.Items sourceItems, int itemIndex, Outlook.Folder destFolder, SemaphoreSlim semaphore, string folderName, System.Threading.CancellationToken ct, Action<int, string> onProgress)
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

                    string currentItem = GetItemSubject(dynItem) ?? "<No Subject>";
                    onProgress(-2, string.Format("Copying item #{0} in {1}: {2}", itemIndex, folderName, currentItem));

                    // Check for duplicates before copying
                    if (await IsDuplicateItemAsync(dynItem, destFolder, ct))
                    {
                        onProgress(-2, string.Format("Skipping duplicate item #{0} in {1}: {2}", itemIndex, folderName, currentItem));
                        return; // Skip this item
                    }

                    // Attempt to copy item (retry on transient errors only)
                    int maxRetries = 2;
                    Exception lastException = null;

                    for (int attempt = 1; attempt <= maxRetries; attempt++)
                    {
                        Exception delayEx = null;
                        try
                        {
                            // We copy and then move to preserve the source PST in case of failure
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
                            await Task.Delay(25, ct); // Use async delay outside catch
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
                // Get key properties from source item
                string subject = GetItemSubject(sourceItem);
                DateTime? receivedTime = GetItemReceivedTime(sourceItem);
                string senderEmail = GetItemSenderEmail(sourceItem);

                if (string.IsNullOrEmpty(subject))
                    return false; // Can't check duplicates without subject

                // Check existing items in destination folder
                Outlook.Items destItems = destFolder.Items;
                try
                {
                    // Use Restrict method to filter items efficiently
                    string filter = string.Format("[Subject] = '{0}'", subject.Replace("'", "''"));
                    Outlook.Items filteredItems = destItems.Restrict(filter);

                    try
                    {
                        foreach (dynamic destItem in filteredItems)
                        {
                            if (ct.IsCancellationRequested)
                                return false;

                            try
                            {
                                // Check if this is likely the same item
                                string destSubject = GetItemSubject(destItem);
                                DateTime? destReceivedTime = GetItemReceivedTime(destItem);
                                string destSenderEmail = GetItemSenderEmail(destItem);

                                // Consider it a duplicate if subject, sender, and received time match (within 1 minute)
                                if (string.Equals(subject, destSubject, StringComparison.OrdinalIgnoreCase) &&
                                    string.Equals(senderEmail, destSenderEmail, StringComparison.OrdinalIgnoreCase) &&
                                    receivedTime.HasValue && destReceivedTime.HasValue &&
                                    Math.Abs((receivedTime.Value - destReceivedTime.Value).TotalMinutes) < 1)
                                {
                                    return true; // Found duplicate
                                }
                            }
                            finally
                            {
                                Marshal.ReleaseComObject(destItem);
                            }
                        }
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

                return false; // No duplicate found
            }
            catch
            {
                // If duplicate checking fails, err on the side of caution and allow the copy
                return false;
            }
        }

        private string GetItemSubject(dynamic item)
        {
            try
            {
                return item.Subject as string;
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
