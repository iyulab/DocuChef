namespace DocuChef.Presentation.Models;

/// <summary>
/// Defines slide types
/// </summary>
public enum SlideType
{
    /// <summary>
    /// Regular slide with no directives
    /// </summary>
    Regular,

    /// <summary>
    /// Slide with foreach directive
    /// </summary>
    Foreach,

    /// <summary>
    /// Slide with if directive
    /// </summary>
    If
}
