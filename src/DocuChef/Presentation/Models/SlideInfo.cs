namespace DocuChef.Presentation.Models;

/// <summary>
/// Information about a slide extracted from template analysis
/// Contains slide identification, type, directives, and binding expressions
/// </summary>
public class SlideInfo
{
    /// <summary>
    /// Slide ID in the presentation
    /// </summary>
    public int SlideId { get; set; }
    
    /// <summary>
    /// Type of slide (Static, Source, Cloned)
    /// </summary>
    public SlideType Type { get; set; }
    
    /// <summary>
    /// Original position in template
    /// </summary>
    public int Position { get; set; }
    
    /// <summary>
    /// Collection name if this slide processes a collection
    /// </summary>
    public string? CollectionName { get; set; }
    
    /// <summary>
    /// List of directives found in slide notes
    /// </summary>
    public List<Directive> Directives { get; set; } = new List<Directive>();
    
    /// <summary>
    /// List of binding expressions found in slide content
    /// </summary>
    public List<BindingExpression> BindingExpressions { get; set; } = new List<BindingExpression>();
    
    /// <summary>
    /// Maximum array index found in expressions (for determining items per slide)
    /// </summary>
    public int MaxArrayIndex { get; set; } = -1;
    
    /// <summary>
    /// Whether this slide contains array references
    /// </summary>
    public bool HasArrayReferences => MaxArrayIndex >= 0;
    
    /// <summary>
    /// Number of items this slide can display (MaxArrayIndex + 1)
    /// </summary>
    public int ItemsPerSlide => MaxArrayIndex + 1;
}

/// <summary>
/// Types of slides in the template processing
/// </summary>
public enum SlideType
{
    /// <summary>
    /// Static slide with no data binding
    /// </summary>
    Static,
    
    /// <summary>
    /// Source slide that will be cloned for collections
    /// </summary>
    Source,
    
    /// <summary>
    /// Cloned slide generated from source
    /// </summary>
    Cloned
}