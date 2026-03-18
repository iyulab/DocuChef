namespace DocuChef.Word;

/// <summary>
/// Represents a generated Word document
/// </summary>
public class WordDocument : IDish
{
    private MemoryStream? _stream;
    private string? _filePath;
    private bool _isDisposed;

    /// <summary>
    /// Gets the file path of the document
    /// </summary>
    public string? FilePath => _filePath;

    /// <summary>
    /// Creates a new Word document from a memory stream
    /// </summary>
    internal WordDocument(MemoryStream stream)
    {
        _stream = stream ?? throw new ArgumentNullException(nameof(stream));
    }

    /// <summary>
    /// Creates a new Word document from a memory stream with a file path
    /// </summary>
    internal WordDocument(MemoryStream stream, string filePath) : this(stream)
    {
        _filePath = filePath;
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
            FileExtensions.EnsureDirectoryExists(filePath);

            using var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write);
            _stream!.Position = 0;
            _stream.CopyTo(fileStream);

            _filePath = filePath;
            Logger.Info($"Word document saved to {filePath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to save Word document to {filePath}", ex);
            throw new DocuChefException($"Failed to save Word document: {ex.Message}", ex);
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
            _stream!.Position = 0;
            _stream.CopyTo(stream);
            Logger.Info("Word document saved to stream");
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to save Word document to stream", ex);
            throw new DocuChefException($"Failed to save Word document to stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Opens the document with the default application
    /// </summary>
    public void Open()
    {
        ThrowIfDisposed();

        if (string.IsNullOrEmpty(_filePath))
        {
            // Save to a temporary file first
            string tempPath = Path.Combine(Path.GetTempPath(), $"DocuChef_{Guid.NewGuid():N}.docx");
            SaveAs(tempPath);
        }

        FileUtility.OpenWithDefaultApplication(_filePath!);
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
    /// Protected implementation of Dispose pattern
    /// </summary>
    protected virtual void Dispose(bool disposing)
    {
        if (_isDisposed) return;

        if (disposing)
        {
            _stream?.Dispose();
            _stream = null;
            Logger.Debug("Word document disposed");
        }

        _isDisposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(WordDocument));
    }
}
