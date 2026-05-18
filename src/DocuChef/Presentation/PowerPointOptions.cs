namespace DocuChef.Presentation;

/// <summary>
/// Configuration options for PowerPoint template processing
/// </summary>
public class PowerPointOptions
{
    /// <summary>
    /// Enable verbose logging for troubleshooting
    /// </summary>
    public bool EnableVerboseLogging { get; set; }
    
    /// <summary>
    /// Throw exception when a variable is not found
    /// </summary>
    public bool ThrowOnMissingVariable { get; set; }
    
    /// <summary>
    /// Whether to populate global variables (Today, Now, UserName, etc.)
    /// </summary>
    public bool RegisterGlobalVariables { get; set; } = true;

}