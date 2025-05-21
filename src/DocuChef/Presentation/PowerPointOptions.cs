namespace DocuChef.Presentation;

/// <summary>
/// Options for customizing PowerPoint template processing
/// </summary>
public class PowerPointOptions
{
    /// <summary>
    /// Whether to analyze the template immediately upon initialization
    /// </summary>
    public bool AnalyzeOnInit { get; set; } = true;

    /// <summary>
    /// Whether to register standard global variables
    /// </summary>
    public bool RegisterGlobalVariables { get; set; } = true;

    /// <summary>
    /// Whether to enable verbose logging for debugging
    /// </summary>
    public bool EnableVerboseLogging { get; set; } = false;

    /// <summary>
    /// Whether to throw exceptions for missing variables instead of showing placeholders
    /// </summary>
    public bool ThrowOnMissingVariable { get; set; } = false;

    /// <summary>
    /// Whether to update template with detected implicit directives
    /// </summary>
    public bool UpdateImplicitDirectives { get; set; } = true;

    /// <summary>
    /// Whether to create or update slide notes
    /// </summary>
    public bool EnableSlideNotes { get; set; } = false;

    /// <summary>
    /// Whether to create slide notes only for directives (if slide notes are enabled)
    /// </summary>
    public bool OnlyCreateDirectiveNotes { get; set; } = true;

    /// <summary>
    /// Maximum number of items to process in iterations (like foreach)
    /// </summary>
    public int MaxIterationItems { get; set; } = 1000;

    /// <summary>
    /// The delimiter character(s) used to define collection hierarchy relationships
    /// Default is underscore (>) for backward compatibility
    /// </summary>
    public string HierarchyDelimiter { get; set; } = ">";

    // Static current instance for global access
    internal static PowerPointOptions Current { get; private set; } = new PowerPointOptions();

    /// <summary>
    /// Sets the current options instance for global access
    /// </summary>
    internal static void SetCurrentOptions(PowerPointOptions options)
    {
        Current = options ?? new PowerPointOptions();
    }
}