namespace DocuChef.PowerPoint;

/// <summary>
/// Represents an array reference in document template
/// </summary>
public class ArrayReference
{
    /// <summary>
    /// The name of the array
    /// </summary>
    public string ArrayName { get; set; }

    /// <summary>
    /// The index referenced in the array
    /// </summary>
    public int Index { get; set; }

    /// <summary>
    /// The property path after the array index (if any)
    /// </summary>
    public string PropertyPath { get; set; }

    /// <summary>
    /// The full pattern matched in the text
    /// </summary>
    public string Pattern { get; set; }
}