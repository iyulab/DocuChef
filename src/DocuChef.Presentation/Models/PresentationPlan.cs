namespace DocuChef.Presentation.Models;

/// <summary>
/// Represents the plan for generating the presentation
/// </summary>
public class PresentationPlan
{
    /// <summary>
    /// List of planned slides
    /// </summary>
    public List<PlannedSlide> Slides { get; set; } = new List<PlannedSlide>();

    /// <summary>
    /// Adds a planned slide to the plan
    /// </summary>
    public void AddSlide(PlannedSlide slide)
    {
        if (slide == null)
            throw new ArgumentNullException(nameof(slide));

        Slides.Add(slide);
    }

    /// <summary>
    /// Gets all slides that will be included in the final presentation
    /// </summary>
    public IEnumerable<PlannedSlide> IncludedSlides =>
        Slides.Where(s => s.Operation != SlideOperation.Skip);

    /// <summary>
    /// Gets all slides that will be skipped in the final presentation
    /// </summary>
    public IEnumerable<PlannedSlide> SkippedSlides =>
        Slides.Where(s => s.Operation == SlideOperation.Skip);

    /// <summary>
    /// Gets planned slide operations summary
    /// </summary>
    public PlanSummary GetSummary()
    {
        return new PlanSummary
        {
            TotalSlides = Slides.Count,
            KeptSlides = Slides.Count(s => s.Operation == SlideOperation.Keep),
            ClonedSlides = Slides.Count(s => s.Operation == SlideOperation.Clone),
            SkippedSlides = Slides.Count(s => s.Operation == SlideOperation.Skip)
        };
    }
}
