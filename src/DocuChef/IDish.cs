using DocuChef.Excel;
using DocuChef.Presentation;

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
    public static void Serve<T>(this T document, string filePath) where T : class
    {
        if (document is ExcelDocument excelDoc)
            excelDoc.SaveAs(filePath);
        else if (document is PowerPointDocument powerPointDoc)
            powerPointDoc.SaveAs(filePath);
        else
            throw new InvalidOperationException($"Document type {typeof(T).Name} is not supported");
    }

    /// <summary>
    /// Serves (saves) a document to a stream
    /// </summary>
    public static void Serve<T>(this T document, Stream stream) where T : class
    {
        if (document is ExcelDocument excelDoc)
            excelDoc.SaveAs(stream);
        else if (document is PowerPointDocument powerPointDoc)
            powerPointDoc.SaveAs(stream);
        else
            throw new InvalidOperationException($"Document type {typeof(T).Name} is not supported");
    }
}