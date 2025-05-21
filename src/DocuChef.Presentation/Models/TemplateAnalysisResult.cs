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
               $"(Regular: {RegularSlides}, " +
               $"Foreach: {ForeachSlides}, " +
               $"If: {IfSlides}{implicitInfo})";
    }
}