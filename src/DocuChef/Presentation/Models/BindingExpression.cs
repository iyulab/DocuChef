using System.Collections.Generic;

namespace DocuChef.Presentation.Models;

/// <summary>
/// Represents a binding expression found in a template
/// </summary>
public class BindingExpression
{
    public BindingExpression()
    {
        // Initialize with empty dictionary to avoid null reference exceptions
        ArrayIndices = new Dictionary<string, int>();
    }

    /// <summary>
    /// The original binding expression text from the template
    /// </summary>
    public string OriginalExpression { get; set; } = string.Empty;
    
    /// <summary>
    /// The data path parsed from the expression
    /// </summary>
    public string DataPath { get; set; } = string.Empty;
    
    /// <summary>
    /// Format specifier for the value
    /// </summary>
    public string FormatSpecifier { get; set; } = string.Empty;
    
    /// <summary>
    /// Whether the expression uses the context operator
    /// </summary>
    public bool UsesContextOperator { get; set; }
    
    /// <summary>
    /// Whether the expression is a conditional expression
    /// </summary>
    public bool IsConditional { get; set; }
    
    /// <summary>
    /// Whether the expression includes a method call
    /// </summary>
    public bool IsMethodCall { get; set; }
    
    /// <summary>
    /// Dictionary of array names to their indices used in the expression
    /// </summary>
    public Dictionary<string, int> ArrayIndices { get; set; }
}
