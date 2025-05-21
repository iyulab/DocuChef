namespace DocuChef.Presentation.Models;

/// <summary>
/// Defines slide types based on their role in the presentation
/// </summary>
public enum SlideType
{
    /// <summary>
    /// Original slide in the template
    /// </summary>
    Original,

    /// <summary>
    /// Slide that is skipped during processing
    /// </summary>
    Skipped,

    /// <summary>
    /// Source slide used as a template for cloning
    /// </summary>
    Source,

    /// <summary>
    /// Cloned slide derived from a source
    /// </summary>
    Cloned
}