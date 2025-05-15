namespace DocuChef.Helpers;

/// <summary>
/// Helper class for OpenXML shape operations (PowerPoint and Word)
/// </summary>
public static class OpenXmlShapeHelper
{
    /// <summary>
    /// Get text content from an OpenXML element (generic for PPT and Word)
    /// </summary>
    public static string GetText<T>(T element) where T : OpenXmlElement
    {
        if (element == null)
            return string.Empty;

        var sb = new StringBuilder();
        var texts = element.Descendants<A.Text>();

        foreach (var text in texts)
        {
            if (!string.IsNullOrEmpty(text.Text))
            {
                sb.Append(text.Text);
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Check if element has expressions
    /// </summary>
    public static bool HasExpressions(OpenXmlElement element)
    {
        return element.Descendants<A.Text>()
            .Any(t => ExpressionHelper.ContainsExpressions(t.Text));
    }
}