using ClosedXML.Excel;

namespace DocuChef.Excel;

/// <summary>
/// Represents a generated Excel document
/// </summary>
public class ExcelDocument : IDish, IDisposable
{
    private readonly IXLWorkbook _workbook;
    private bool _isDisposed;
    private string? _filePath;

    /// <summary>
    /// The underlying XLWorkbook instance
    /// </summary>
    public IXLWorkbook Workbook => _workbook;

    /// <summary>
    /// Gets the file path of the document
    /// </summary>
    public string? FilePath => _filePath;

    internal ExcelDocument(IXLWorkbook workbook)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _filePath = null;
    }

    internal ExcelDocument(IXLWorkbook workbook, string filePath) : this(workbook)
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
            // Ensure directory exists
            FileExtensions.EnsureDirectoryExists(filePath);

            _workbook.SaveAs(filePath);
            _filePath = filePath; // Update the file path
            Logger.Info($"Excel document saved to {filePath}");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to save Excel document to {filePath}", ex);
            throw new DocuChefException($"Failed to save Excel document: {ex.Message}", ex);
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
            _workbook.SaveAs(stream);
            Logger.Info("Excel document saved to stream");
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to save Excel document to stream", ex);
            throw new DocuChefException($"Failed to save Excel document to stream: {ex.Message}", ex);
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
            // If no file path is set, save to a temporary file first
            string tempPath = Path.Combine(Path.GetTempPath(), $"DocuChef_{Guid.NewGuid():N}.xlsx");
            SaveAs(tempPath);
        }

        try
        {
            // Use ProcessStartInfo to launch default Excel viewer
            var psi = new System.Diagnostics.ProcessStartInfo
            {
                FileName = _filePath,
                UseShellExecute = true
            };
            System.Diagnostics.Process.Start(psi);
            Logger.Info("Excel document opened with default application");
        }
        catch (Exception ex)
        {
            Logger.Error($"Failed to open Excel document: {ex.Message}", ex);
            throw new DocuChefException($"Failed to open Excel document: {ex.Message}", ex);
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
            _workbook?.Dispose();
            Logger.Debug("Excel document disposed");
        }

        _isDisposed = true;
    }

    private void ThrowIfDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(ExcelDocument));
    }
}