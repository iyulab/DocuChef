namespace DocuChef.Presentation.Models;

/// <summary>
/// Contains a summary of the presentation plan
/// </summary>
public class PlanSummary
{
    /// <summary>
    /// Total slides in the plan
    /// </summary>
    public int TotalSlides { get; set; }

    /// <summary>
    /// Number of slides to keep as is
    /// </summary>
    public int KeptSlides { get; set; }

    /// <summary>
    /// Number of slides to clone
    /// </summary>
    public int ClonedSlides { get; set; }

    /// <summary>
    /// Number of slides to skip
    /// </summary>
    public int SkippedSlides { get; set; }

    /// <summary>
    /// Total number of slides that will be included in the final presentation
    /// </summary>
    public int TotalIncludedSlides => KeptSlides + ClonedSlides;

    /// <summary>
    /// Formats the summary as a string
    /// </summary>
    public override string ToString()
    {
        return $"Total: {TotalSlides}, " +
               $"Included: {TotalIncludedSlides} ({KeptSlides} kept, {ClonedSlides} cloned), " +
               $"Skipped: {SkippedSlides}";
    }
}
