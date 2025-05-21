using DocuChef.Presentation.Directives;
using DocuChef.Presentation.Models;

namespace DocuChef.Presentation.Processing;

/// <summary>
/// Helper class to group slides by type
/// </summary>
internal class SlideTypeMap
{
    /// <summary>
    /// Original slides without directives
    /// </summary>
    public List<SlideInfo> OriginalSlides { get; set; } = new List<SlideInfo>();

    /// <summary>
    /// Source slides for collections (with foreach directive)
    /// </summary>
    public List<SlideInfo> SourceSlides { get; set; } = new List<SlideInfo>();

    /// <summary>
    /// Cloned slides generated from source slides
    /// </summary>
    public List<SlideInfo> ClonedSlides { get; set; } = new List<SlideInfo>();

    /// <summary>
    /// Skipped slides that won't be included in the final presentation
    /// </summary>
    public List<SlideInfo> SkippedSlides { get; set; } = new List<SlideInfo>();

    /// <summary>
    /// Groups slides by their type
    /// </summary>
    public static SlideTypeMap GroupSlidesByType(List<SlideInfo> slides)
    {
        var result = new SlideTypeMap();

        foreach (var slide in slides)
        {
            switch (slide.Type)
            {
                case SlideType.Original:
                    result.OriginalSlides.Add(slide);
                    break;
                case SlideType.Source:
                    result.SourceSlides.Add(slide);
                    break;
                case SlideType.Cloned:
                    result.ClonedSlides.Add(slide);
                    break;
                case SlideType.Skipped:
                    result.SkippedSlides.Add(slide);
                    break;
            }
        }

        return result;
    }

    /// <summary>
    /// Groups slides by their directive type
    /// </summary>
    public static SlideTypeMap GroupSlidesByDirectiveType(List<SlideInfo> slides)
    {
        var result = new SlideTypeMap();

        foreach (var slide in slides)
        {
            if (!slide.HasDirective)
            {
                result.OriginalSlides.Add(slide);
            }
            else if (slide.DirectiveType == DirectiveType.Foreach)
            {
                result.SourceSlides.Add(slide);
            }
            else if (slide.DirectiveType == DirectiveType.If)
            {
                // If directive slides could be either included (Source) or excluded (Skipped)
                // depending on condition evaluation
                result.SourceSlides.Add(slide);
            }
        }

        return result;
    }
}