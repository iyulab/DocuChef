namespace DocuChef.Helpers;

public static class FileUtility
{
    /// <summary>
    /// Copies a file with retry logic for handling file access conflicts
    /// </summary>
    public static void CopyFileWithRetry(string sourcePath, string destinationPath, int maxRetries)
    {
        RetryOperation<bool>(() =>
        {
            File.Copy(sourcePath, destinationPath, true);
            return true;
        }, maxRetries);
    }

    /// <summary>
    /// Moves a file with retry logic for handling file access conflicts
    /// </summary>
    public static void MoveFileWithRetry(string sourcePath, string destinationPath, int maxRetries)
    {
        bool success = RetryOperation<bool>(() =>
        {
            // Delete destination if it exists, then move the file
            if (File.Exists(destinationPath))
            {
                File.Delete(destinationPath);
            }
            File.Move(sourcePath, destinationPath);
            return true;
        }, maxRetries, cleanupOnFailure: () =>
        {
            // If all retries failed, try copy as fallback
            try
            {
                File.Copy(sourcePath, destinationPath, true);
                File.Delete(sourcePath); // Clean up source
                return true;
            }
            catch (Exception copyEx)
            {
                Logger.Error($"Final copy attempt also failed: {copyEx.Message}");
                throw new DocuChefException($"Could not write to output file: {destinationPath}.", copyEx);
            }
        });

        // Clean up source file if it still exists and we succeeded
        if (success && File.Exists(sourcePath))
        {
            try
            {
                File.Delete(sourcePath);
            }
            catch (Exception ex)
            {
                // Log but don't fail if we can't delete the temp file
                Logger.Warning($"Could not delete temporary file {sourcePath}: {ex.Message}");
            }
        }
    }    /// <summary>
    /// Generic retry operation with configurable handler
    /// </summary>
    public static T RetryOperation<T>(Func<T> operation, int maxRetries,
        Action<int, Exception>? onRetry = null,
        string? failureMessage = null,
        Func<T>? cleanupOnFailure = null)
    {
        int retryCount = 0;

        while (true)
        {
            try
            {
                return operation();
            }
            catch (IOException ex)
            {
                retryCount++;

                if (retryCount >= maxRetries)
                {
                    if (cleanupOnFailure != null)
                    {
                        return cleanupOnFailure();
                    }

                    Logger.Error($"{failureMessage ?? "Operation failed"} after {maxRetries} attempts: {ex.Message}");
                    throw;
                }

                onRetry?.Invoke(retryCount, ex);

                // Wait a bit before retrying with exponential backoff
                System.Threading.Thread.Sleep(500 * retryCount);
            }
        }
    }

    /// <summary>
    /// Opens a file with the default application registered for its file type
    /// </summary>
    /// <param name="filePath">Path to the file to open</param>
    public static void OpenWithDefaultApplication(string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));

        if (!File.Exists(filePath))
            throw new FileNotFoundException("File not found", filePath);

        try
        {
            // Use ProcessStartInfo to open the file with default application
            using var process = new System.Diagnostics.Process();
            process.StartInfo.FileName = filePath;
            process.StartInfo.UseShellExecute = true;
            process.Start();
        }
        catch (Exception ex)
        {
            // Log error but don't throw
            Logger.Error($"Failed to open file {filePath}: {ex.Message}");
        }
    }
}