namespace DocuChef.PowerPoint;

/// <summary>
/// Extension methods for working with OpenXml objects
/// </summary>
public static class PresentationElementExtensions
{
    /// <summary>
    /// Gets the name of a PowerPoint shape
    /// </summary>
    public static string GetShapeName(this Shape shape)
    {
        if (shape?.NonVisualShapeProperties?.NonVisualDrawingProperties == null)
            return null;

        var nvdp = shape.NonVisualShapeProperties.NonVisualDrawingProperties;

        // Try Name property
        if (!string.IsNullOrWhiteSpace(nvdp.Name?.Value))
            return nvdp.Name.Value;

        // Try Title property  
        if (!string.IsNullOrWhiteSpace(nvdp.Title?.Value))
            return nvdp.Title.Value;

        return null;
    }

    /// <summary>
    /// Gets the text content of an element
    /// </summary>
    public static string GetText(this OpenXmlElement element)
    {
        return OpenXmlShapeHelper.GetText(element);
    }

    /// <summary>
    /// Gets notes from a PowerPoint slide
    /// </summary>
    public static string GetNotes(this SlidePart slidePart)
    {
        if (slidePart?.NotesSlidePart?.NotesSlide == null)
            return string.Empty;

        // Get all text from notes slide
        var allTexts = slidePart.NotesSlidePart.NotesSlide
            .Descendants<A.Text>()
            .Select(t => t.Text)
            .Where(t => !string.IsNullOrEmpty(t))
            .ToList();

        // Look for directive text (starts with #)
        var directiveText = allTexts.FirstOrDefault(t => t.StartsWith("#"));
        if (!string.IsNullOrEmpty(directiveText))
            return directiveText;

        // Return combined text excluding numbers
        return string.Join(" ", allTexts.Where(t => !Regex.IsMatch(t, @"^\d+$")));
    }

    /// <summary>
    /// Check if element contains expressions
    /// </summary>
    public static bool ContainsExpressions(this OpenXmlElement element)
    {
        return OpenXmlShapeHelper.HasExpressions(element);
    }
}