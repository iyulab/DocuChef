using DocuChef.Helpers;

namespace DocuChef.Presentation;

/// <summary>
/// Represents a PowerPoint document generated by the template engine
/// </summary>
public class PowerPointDocument : IDish
{
    private readonly string filePath;
    private bool isDisposed;
    
    /// <summary>
    /// Creates a new PowerPoint document
    /// </summary>
    /// <param name="filePath">Path to the document file</param>
    public PowerPointDocument(string filePath)
    {
        this.filePath = filePath;
    }
    
    /// <summary>
    /// Gets the file path of the document
    /// </summary>
    public string FilePath => filePath;
    
    /// <summary>
    /// Saves the document to a new file path
    /// </summary>
    /// <param name="filePath">The target file path</param>
    public void SaveAs(string filePath)
    {
        if (isDisposed)
            throw new ObjectDisposedException(nameof(PowerPointDocument));
        
        if (string.IsNullOrEmpty(filePath))
            throw new ArgumentNullException(nameof(filePath));
        
        // Copy the document to the new path
        File.Copy(this.filePath, filePath, true);
    }
    
    /// <summary>
    /// Saves the document to a stream
    /// </summary>
    /// <param name="stream">The target stream</param>
    public void SaveAs(Stream stream)
    {
        if (isDisposed)
            throw new ObjectDisposedException(nameof(PowerPointDocument));
        
        if (stream == null)
            throw new ArgumentNullException(nameof(stream));
        
        // Copy the document to the stream
        using (var fileStream = File.OpenRead(filePath))
        {
            fileStream.CopyTo(stream);
        }
    }
    
    /// <summary>
    /// Opens the document with the default application
    /// </summary>
    public void Open()
    {
        if (isDisposed)
            throw new ObjectDisposedException(nameof(PowerPointDocument));
        
        // Open the document with the default application
        FileUtility.OpenWithDefaultApplication(filePath);
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
        if (!isDisposed)
        {
            if (disposing)
            {
                // Delete the temporary file if it's in the temp directory
                if (filePath.Contains(Path.GetTempPath()) && File.Exists(filePath))
                {
                    try
                    {
                        File.Delete(filePath);
                    }
                    catch
                    {
                        // Ignore errors during cleanup
                    }
                }
            }
            
            isDisposed = true;
        }
    }
    
    /// <summary>
    /// Finalizer
    /// </summary>
    ~PowerPointDocument()
    {
        Dispose(false);
    }
}
