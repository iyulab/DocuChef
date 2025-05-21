using DocuChef.Presentation.Models;

namespace DocuChef.Presentation.Processing;

/// <summary>
/// Helper class to group slides by type
/// </summary>
internal class SlideTypeMap
{
    /// <summary>
    /// Regular slides without directives
    /// </summary>
    public List<SlideInfo> RegularSlides { get; set; } = new List<SlideInfo>();

    /// <summary>
    /// Slides with foreach directives
    /// </summary>
    public List<SlideInfo> ForeachSlides { get; set; } = new List<SlideInfo>();

    /// <summary>
    /// Slides with if directives
    /// </summary>
    public List<SlideInfo> IfSlides { get; set; } = new List<SlideInfo>();
}