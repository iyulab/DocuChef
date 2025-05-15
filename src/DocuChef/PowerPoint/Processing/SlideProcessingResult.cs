namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Result of processing slide directives
/// </summary>
internal class SlideProcessingResult
{
    /// <summary>
    /// The original slide part
    /// </summary>
    public SlidePart SlidePart { get; set; }

    /// <summary>
    /// Whether the slide was processed
    /// </summary>
    public bool WasProcessed { get; set; }

    /// <summary>
    /// Slides generated during processing
    /// </summary>
    public List<SlidePart> GeneratedSlides { get; set; } = new List<SlidePart>();
}