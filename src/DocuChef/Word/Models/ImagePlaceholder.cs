namespace DocuChef.Word.Models;

/// <summary>
/// Represents an image placeholder in a Word template
/// </summary>
public class ImagePlaceholder
{
    /// <summary>
    /// Path to the image file
    /// </summary>
    public required string Path { get; set; }

    /// <summary>
    /// Optional width in EMUs (English Metric Units)
    /// </summary>
    public long? Width { get; set; }

    /// <summary>
    /// Optional height in EMUs (English Metric Units)
    /// </summary>
    public long? Height { get; set; }
}
