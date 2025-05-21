using DocuChef.Presentation.Directives;

namespace DocuChef.Presentation.Models;

/// <summary>
/// Represents the overall presentation context
/// </summary>
internal class PPTContext
{
    /// <summary>
    /// Collection of slide contexts in the presentation
    /// </summary>
    public List<SlideContext> Slides { get; set; } = new List<SlideContext>();

    /// <summary>
    /// Collection of template slide information
    /// </summary>
    public List<SlideTemplateContext> TemplateSlides { get; set; } = new List<SlideTemplateContext>();

    /// <summary>
    /// Title of the presentation
    /// </summary>
    public string Title { get; set; }

    /// <summary>
    /// Gets template statistics
    /// </summary>
    public TemplateStatistics GetTemplateStatistics()
    {
        return new TemplateStatistics
        {
            TotalSlides = TemplateSlides.Count,
            ForeachSlides = TemplateSlides.Count(s => s.TemplateType == SlideTemplateType.Foreach),
            IfSlides = TemplateSlides.Count(s => s.TemplateType == SlideTemplateType.If),
            RegularSlides = TemplateSlides.Count(s => s.TemplateType == SlideTemplateType.Regular)
        };
    }
}

/// <summary>
/// Contains statistics about the template
/// </summary>
internal class TemplateStatistics
{
    /// <summary>
    /// Total number of slides in the template
    /// </summary>
    public int TotalSlides { get; set; }

    /// <summary>
    /// Number of slides with foreach directives
    /// </summary>
    public int ForeachSlides { get; set; }

    /// <summary>
    /// Number of slides with if directives
    /// </summary>
    public int IfSlides { get; set; }

    /// <summary>
    /// Number of regular slides without directives
    /// </summary>
    public int RegularSlides { get; set; }
}

/// <summary>
/// Represents slide information from the template
/// </summary>
internal class SlideTemplateContext
{
    /// <summary>
    /// Gets the slide ID
    /// </summary>
    public uint SlideId { get; set; }

    /// <summary>
    /// Gets the relationship ID
    /// </summary>
    public string RelationshipId { get; set; }

    /// <summary>
    /// Gets the slide note text
    /// </summary>
    public string NoteText { get; set; }

    /// <summary>
    /// Gets the foreach directive if present
    /// </summary>
    public ForeachDirective ForeachDirective { get; set; }

    /// <summary>
    /// Gets the if directive if present
    /// </summary>
    public IfDirective IfDirective { get; set; }

    /// <summary>
    /// Gets a value indicating whether this slide has directives
    /// </summary>
    public bool HasDirectives => ForeachDirective != null || IfDirective != null;

    /// <summary>
    /// Gets a value indicating the type of slide
    /// </summary>
    public SlideTemplateType TemplateType
    {
        get
        {
            if (ForeachDirective != null)
                return SlideTemplateType.Foreach;
            else if (IfDirective != null)
                return SlideTemplateType.If;
            else
                return SlideTemplateType.Regular;
        }
    }
}

/// <summary>
/// Defines the slide template type
/// </summary>
public enum SlideTemplateType
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