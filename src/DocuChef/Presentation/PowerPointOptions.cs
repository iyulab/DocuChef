namespace DocuChef.Presentation;

/// <summary>
/// Options for customizing PowerPoint template processing
/// </summary>
public class PowerPointOptions
{
    // Private static instance with thread safety
    private static readonly object _lock = new object();
    private static PowerPointOptions _current = new PowerPointOptions();

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
    /// Default is greater than (>) for nested collections
    /// </summary>
    public string HierarchyDelimiter { get; set; } = ">";

    /// <summary>
    /// Gets the current global options instance
    /// </summary>
    internal static PowerPointOptions Current => _current;

    /// <summary>
    /// Sets the current options instance for global access in a thread-safe manner
    /// </summary>
    internal static void SetCurrentOptions(PowerPointOptions options)
    {
        if (options == null)
            throw new ArgumentNullException(nameof(options));

        lock (_lock)
        {
            _current = options;
        }
    }

    /// <summary>
    /// Creates a new instance with default values
    /// </summary>
    public PowerPointOptions() { }

    /// <summary>
    /// Creates a deep copy of the specified options
    /// </summary>
    public PowerPointOptions(PowerPointOptions source)
    {
        if (source == null) return;

        AnalyzeOnInit = source.AnalyzeOnInit;
        RegisterGlobalVariables = source.RegisterGlobalVariables;
        EnableVerboseLogging = source.EnableVerboseLogging;
        ThrowOnMissingVariable = source.ThrowOnMissingVariable;
        UpdateImplicitDirectives = source.UpdateImplicitDirectives;
        EnableSlideNotes = source.EnableSlideNotes;
        OnlyCreateDirectiveNotes = source.OnlyCreateDirectiveNotes;
        MaxIterationItems = source.MaxIterationItems;
        HierarchyDelimiter = source.HierarchyDelimiter;
    }

    /// <summary>
    /// Creates a copy of the current options
    /// </summary>
    public PowerPointOptions Clone() => new PowerPointOptions(this);
}