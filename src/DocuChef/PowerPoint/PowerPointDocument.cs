namespace DocuChef.PowerPoint;

/// <summary>
/// Represents a generated PowerPoint document
/// </summary>
public class PowerPointDocument : IDisposable
{
    private readonly PresentationDocument _presentationDocument;
    private readonly string _documentPath;
    private bool _isDisposed;
    private bool _isSaved;

    /// <summary>
    /// The underlying OpenXml PresentationDocument instance
    /// </summary>
    public PresentationDocument PresentationDocument => _presentationDocument;

    /// <summary>
    /// Creates a new PowerPoint document
    /// </summary>
    internal PowerPointDocument(PresentationDocument presentationDocument, string documentPath)
    {
        _presentationDocument = presentationDocument ?? throw new ArgumentNullException(nameof(presentationDocument));
        _documentPath = documentPath ?? throw new ArgumentNullException(nameof(documentPath));
        _isSaved = false;
    }

    /// <summary>
    /// Saves the document to the specified path
    /// </summary>
    public void SaveAs(string filePath)
    {
        EnsureNotDisposed();

        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));

        try
        {
            // Save all document parts first
            SaveAllDocumentParts();

            // Close the presentation document
            _presentationDocument.Dispose();
            _isSaved = true;

            // Ensure target directory exists
            filePath.EnsureDirectoryExists();

            // Copy to final location
            File.Copy(_documentPath, filePath, true);

            // Verify the output file
            var fileInfo = new FileInfo(filePath);
            if (!fileInfo.Exists || fileInfo.Length == 0)
            {
                throw new DocuChefException("Failed to create valid output file.");
            }

            Logger.Info($"PowerPoint document saved successfully to {filePath}");
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
        EnsureNotDisposed();

        if (stream == null)
            throw new ArgumentNullException(nameof(stream));

        try
        {
            // Save all document parts first
            SaveAllDocumentParts();

            // Close the presentation document
            _presentationDocument.Dispose();
            _isSaved = true;

            // Copy to stream
            using (var fileStream = new FileStream(_documentPath, FileMode.Open, FileAccess.Read))
            {
                fileStream.CopyTo(stream);
            }

            Logger.Info("PowerPoint document saved successfully to stream");
        }
        catch (Exception ex)
        {
            Logger.Error("Failed to save PowerPoint document to stream", ex);
            throw new DocuChefException($"Failed to save PowerPoint document to stream: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Saves all document parts to ensure proper document structure
    /// </summary>
    private void SaveAllDocumentParts()
    {
        try
        {
            if (_presentationDocument?.PresentationPart == null)
                return;

            // Save the presentation
            _presentationDocument.PresentationPart.Presentation.Save();

            // Save all slide parts
            foreach (var slidePart in _presentationDocument.PresentationPart.SlideParts)
            {
                slidePart?.Slide?.Save();
            }

            // Save slide master parts
            foreach (var masterPart in _presentationDocument.PresentationPart.SlideMasterParts)
            {
                masterPart?.SlideMaster?.Save();
            }

            // Save theme part
            _presentationDocument.PresentationPart.ThemePart?.Theme?.Save();

            // Save view properties
            _presentationDocument.PresentationPart.ViewPropertiesPart?.ViewProperties?.Save();

            // Final document save
            _presentationDocument.Save();
        }
        catch (Exception ex)
        {
            Logger.Error("Error saving document parts", ex);
            throw new DocuChefException("Failed to save document parts: " + ex.Message, ex);
        }
    }

    /// <summary>
    /// Ensure the document is not disposed
    /// </summary>
    private void EnsureNotDisposed()
    {
        if (_isDisposed)
            throw new ObjectDisposedException(nameof(PowerPointDocument));
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
            try
            {
                if (!_isSaved && _presentationDocument != null)
                {
                    _presentationDocument.Dispose();
                }

                if (!string.IsNullOrEmpty(_documentPath) && File.Exists(_documentPath))
                {
                    File.Delete(_documentPath);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error disposing PowerPoint document resources", ex);
            }
        }

        _isDisposed = true;
    }
}