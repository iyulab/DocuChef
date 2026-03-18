namespace DocuChef.Word;

/// <summary>
/// Configuration options for Word template processing
/// </summary>
public class WordOptions
{
    /// <summary>
    /// Enable verbose logging for troubleshooting
    /// </summary>
    public bool EnableVerboseLogging { get; set; } = false;

    /// <summary>
    /// Throw exception when a variable is not found
    /// </summary>
    public bool ThrowOnMissingVariable { get; set; } = false;
}
