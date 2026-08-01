namespace DocuChef.Presentation.Models;

/// <summary>
/// Represents a directive for controlling template processing
/// </summary>
public class Directive
{
    public DirectiveType Type { get; set; }
    public string CollectionPath { get; set; } = string.Empty;
    public int MaxItems { get; set; }
    public int Offset { get; set; }
    public string RangeType { get; set; } = string.Empty;
    public string SourceName { get; set; } = string.Empty;

    /// <summary>
    /// Alias of <see cref="CollectionPath"/>, kept for compatibility.
    /// The two were independent fields: directive parsing filled <see cref="CollectionPath"/>
    /// while slide planning read this one, so every directive written by hand resolved to an
    /// empty collection. Sharing one backing property makes that divergence impossible.
    /// </summary>
    [Obsolete("Use CollectionPath instead. This alias is scheduled for removal.")]
    public string SourcePath
    {
        get => CollectionPath;
        set => CollectionPath = value;
    }

    public string AliasName { get; set; } = string.Empty;
    public RangeBoundary RangeBoundary { get; set; } = RangeBoundary.Single;
}

/// <summary>
/// Types of directives supported by the template engine
/// </summary>
public enum DirectiveType
{
    Foreach,
    Range,
    Alias
}

/// <summary>
/// Boundaries for range directives
/// </summary>
public enum RangeBoundary
{
    Single,
    Begin,
    End
}
