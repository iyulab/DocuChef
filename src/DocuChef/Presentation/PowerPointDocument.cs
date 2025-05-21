namespace DocuChef.Presentation;

/// <summary>
/// Represents a generated PowerPoint document
/// </summary>
public class PowerPointDocument : IDish, IDisposable
{
    private readonly string _filePath;
    private bool _isDisposed;

    /// <summary>
    /// Gets the file path of the document
    /// </summary>
    public string FilePath => _filePath;

    internal PowerPointDocument(string filePath)
    {
        _filePath = filePath ?? throw new ArgumentNullException(nameof(filePath));
    }

    /// <summary>
    /// Saves the document to the specified path
    /// </summary>
    public void SaveAs(string filePath)
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));

        try
        {
            // Ensure directory exists
            FileExtensions.EnsureDirectoryExists(filePath);

            // Create a copy of the document
            File.Copy(_filePath, filePath, true);
            Logger.Info($"PowerPoint document saved to {filePath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to save PowerPoint document to {filePath}", ex);
            throw new DocuChefException($"Failed to save PowerPoint document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Saves the document to a stream
    /// </summary>
    public void SaveAs(Stream stream)
    {
        ThrowIfDisposed();

        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        try
        {
            using (var fileStream = new FileStream(_filePath, FileMode.Open, FileAccess.Read))
            {
                fileStream.CopyTo(stream);
            }
            Logger.Info("PowerPoint document saved to stream");
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to save PowerPoint document to stream", ex);
            throw new DocuChefException($"Failed to save PowerPoint document to stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Opens the document with the default application
    /// </summary>
    public void Open()
    {
        ThrowIfDisposed();

        try
        {
            // Use ProcessStartInfo to launch default PowerPoint viewer
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = _filePath,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
            Logger.Info("PowerPoint document opened with default application");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to open PowerPoint document: {ex.Message}", ex);
            throw new DocuChefException($"Failed to open PowerPoint document: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    public void Dispose()
    {
        Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// Disposes resources
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        if (_isDisposed) return;

        if (disposing)
        {
            // Delete the temporary file if it exists
            if (File.Exists(_filePath) && _filePath.Contains(Path.GetTempPath()))
            {
                try
                {
                    File.Delete(_filePath);
                    Logger.Debug("Temporary PowerPoint document deleted");
                }
                catch (Exception ex)
                {
                    Logger.Debug($"Failed to delete temporary PowerPoint document: {ex.Message}");
                }
            }

            Logger.Debug("PowerPoint document disposed");
        }

        _isDisposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(PowerPointDocument));
    }
}