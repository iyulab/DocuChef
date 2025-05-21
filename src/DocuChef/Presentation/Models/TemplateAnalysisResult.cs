namespace DocuChef.Presentation.Models;

/// <summary>
/// Contains template analysis results
/// </summary>
public class TemplateAnalysisResult
{
    /// <summary>
    /// Total number of slides in the template
    /// </summary>
    public int TotalSlides { get; set; }

    /// <summary>
    /// Number of source slides for foreach directives
    /// </summary>
    public int ForeachSourceSlides { get; set; }

    /// <summary>
    /// Number of source slides for if directives
    /// </summary>
    public int IfSourceSlides { get; set; }

    /// <summary>
    /// Number of original slides without directives
    /// </summary>
    public int OriginalSlides { get; set; }

    /// <summary>
    /// Number of slides with implicitly detected directives
    /// </summary>
    public int ImplicitDirectives { get; set; }

    /// <summary>
    /// Returns a string summary of the analysis
    /// </summary>
    public override string ToString()
    {
        string implicitInfo = ImplicitDirectives > 0 ? $", Implicit: {ImplicitDirectives}" : "";

        return $"Total: {TotalSlides} slides " +
               $"(Original: {OriginalSlides}, " +
               $"Foreach source: {ForeachSourceSlides}, " +
               $"If source: {IfSourceSlides}{implicitInfo})";
    }
}