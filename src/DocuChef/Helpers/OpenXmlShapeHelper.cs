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
    /// Set text content in an OpenXML element
    /// </summary>
    public static void SetText<T>(T element, string text) where T : OpenXmlElement
    {
        if (element == null)
            return;

        // Clear existing text
        var existingTexts = element.Descendants<A.Text>().ToList();
        foreach (var existingText in existingTexts)
        {
            existingText.Text = "";
        }

        // Set new text in first text element or create new one
        var firstText = existingTexts.FirstOrDefault();
        if (firstText != null)
        {
            firstText.Text = text;
        }
    }

    /// <summary>
    /// Hide an OpenXML element
    /// </summary>
    public static void HideElement(OpenXmlElement element)
    {
        if (element == null)
            return;

        // Set Hidden attribute if available
        var hiddenAttr = element.GetAttributes().FirstOrDefault(a => a.LocalName == "hidden");
        if (hiddenAttr.Value != null)
        {
            element.SetAttribute(new OpenXmlAttribute("", "hidden", "", "true"));
        }

        // Set visibility for Drawing elements
        if (element is DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties nvdp)
        {
            nvdp.Hidden = true;
        }
    }

    /// <summary>
    /// Show an OpenXML element
    /// </summary>
    public static void ShowElement(OpenXmlElement element)
    {
        if (element == null)
            return;

        // Remove Hidden attribute if available
        var hiddenAttr = element.GetAttributes().FirstOrDefault(a => a.LocalName == "hidden");
        if (hiddenAttr.Value != null)
        {
            element.SetAttribute(new OpenXmlAttribute("", "hidden", "", "false"));
        }

        // Set visibility for Drawing elements
        if (element is DocumentFormat.OpenXml.Drawing.NonVisualDrawingProperties nvdp)
        {
            nvdp.Hidden = false;
        }
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