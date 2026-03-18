namespace DocuChef;

/// <summary>
/// Interface for all document outputs (dishes) in DocuChef
/// </summary>
public interface IDish : IDisposable
{    /// <summary>
     /// Gets the file path of the document
     /// </summary>
    string? FilePath { get; }

    /// <summary>
    /// Saves the document to the specified path
    /// </summary>
    void SaveAs(string filePath);

    /// <summary>
    /// Saves the document to a stream
    /// </summary>
    void SaveAs(Stream stream);

    /// <summary>
    /// Opens the document with the default application
    /// </summary>
    void Open();
}


/// <summary>
/// Provides cooking-themed extension methods for documents
/// </summary>
public static class DishExtensions
{
    /// <summary>
    /// Serves (saves) a document to a file
    /// </summary>
    public static void Serve(this IDish dish, string filePath) => dish.SaveAs(filePath);

    /// <summary>
    /// Serves (saves) a document to a stream
    /// </summary>
    public static void Serve(this IDish dish, Stream stream) => dish.SaveAs(stream);

    /// <summary>
    /// Presents (opens) a document with the default application
    /// </summary>
    public static void Present(this IDish dish) => dish.Open();
}