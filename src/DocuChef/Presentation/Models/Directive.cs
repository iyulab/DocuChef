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
    public string SourcePath { get; set; } = string.Empty;
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
