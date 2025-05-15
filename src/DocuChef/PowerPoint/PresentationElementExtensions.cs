using DocuChef.PowerPoint.Helpers;

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
    /// Sets text content in an element
    /// </summary>
    public static void SetText(this OpenXmlElement element, string text)
    {
        OpenXmlShapeHelper.SetText(element, text);
    }

    /// <summary>
    /// Clears text content from a PowerPoint shape
    /// </summary>
    public static void ClearText(this Shape shape)
    {
        if (shape?.TextBody == null)
            return;

        // Remove all existing paragraphs
        var existingParagraphs = shape.TextBody.Elements<A.Paragraph>().ToList();
        foreach (var paragraph in existingParagraphs)
        {
            shape.TextBody.RemoveChild(paragraph);
        }

        // Add a single empty paragraph to maintain document structure
        var emptyParagraph = new A.Paragraph();
        var run = new A.Run();
        var text = new A.Text("");
        run.AppendChild(text);
        emptyParagraph.AppendChild(run);
        shape.TextBody.AppendChild(emptyParagraph);
    }

    /// <summary>
    /// Sets visibility of a PowerPoint shape
    /// </summary>
    public static void SetVisibility(this Shape shape, bool visible)
    {
        if (visible)
            PowerPointShapeHelper.ShowShape(shape);
        else
            PowerPointShapeHelper.HideShape(shape);
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
    /// Creates a copy of an OpenXml part
    /// </summary>
    public static void CopyTo(this OpenXmlPart sourcePart, OpenXmlPart targetPart)
    {
        using (var stream = sourcePart.GetStream())
        {
            stream.Position = 0;
            targetPart.FeedData(stream);
        }
    }

    /// <summary>
    /// Check if element contains expressions
    /// </summary>
    public static bool ContainsExpressions(this OpenXmlElement element)
    {
        return OpenXmlShapeHelper.HasExpressions(element);
    }

    /// <summary>
    /// Process expressions in element
    /// </summary>
    public static string ProcessExpressions(this OpenXmlElement element, IExpressionEvaluator evaluator, Dictionary<string, object> variables)
    {
        string text = element.GetText();
        if (string.IsNullOrEmpty(text))
            return text;

        return ExpressionHelper.ProcessExpressions(text, evaluator, variables);
    }
}