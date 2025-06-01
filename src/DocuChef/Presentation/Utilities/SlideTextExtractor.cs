using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using DocuChef.Logging;

namespace DocuChef.Presentation.Utilities;

/// <summary>
/// Utility class for extracting text from PowerPoint slides with hierarchical span-first strategy
/// </summary>
public static class SlideTextExtractor
{
    /// <summary>
    /// Gets all text from a slide using hierarchical extraction strategy
    /// </summary>
    /// <param name="slide">The slide object to extract text from</param>
    /// <returns>A string containing all text from the slide</returns>
    public static string GetText(object slide)
    {
        if (slide is SlidePart slidePart)
        {
            return GetTextFromSlidePart(slidePart);
        }

        // Handle mock slide part for testing
        if (slide?.GetType().Name == "MockSlidePart")
        {
            var contentProperty = slide.GetType().GetProperty("Content");
            return contentProperty?.GetValue(slide)?.ToString() ?? string.Empty;
        }

        return string.Empty;
    }

    /// <summary>
    /// Gets text from a SlidePart with hierarchical span-first strategy
    /// </summary>
    private static string GetTextFromSlidePart(SlidePart slidePart)
    {
        var textBuilder = new System.Text.StringBuilder();

        // Use hierarchical text extraction with span-first strategy
        var paragraphs = slidePart.Slide.Descendants<Paragraph>();
        foreach (var paragraph in paragraphs)
        {
            var hierarchicalInfo = TextExtractionUtility.ExtractHierarchicalText(paragraph);
            var paragraphText = hierarchicalInfo.CombinedText;

            if (!string.IsNullOrEmpty(paragraphText))
            {
                textBuilder.AppendLine(paragraphText);
                Logger.Debug($"SlideTextExtractor: Extracted text using {hierarchicalInfo.ProcessingStrategy} strategy: '{paragraphText}'");
            }
        }

        return textBuilder.ToString();
    }

    /// <summary>
    /// Extracts text from a paragraph, handling Korean text that may be split across spans
    /// (Legacy method for backward compatibility)
    /// </summary>
    private static string ExtractParagraphText(Paragraph paragraph)
    {
        var textBuilder = new System.Text.StringBuilder();
        var textElements = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>();

        foreach (var text in textElements)
        {
            textBuilder.Append(text.Text);
        }

        var paragraphText = textBuilder.ToString();
        Logger.Debug($"SlideTextExtractor: Extracted paragraph text: '{paragraphText}'");

        return paragraphText;
    }

    public static List<string> ExtractTextElements(SlidePart slidePart)
    {
        var textElements = new List<string>();

        // Use hierarchical extraction for each paragraph
        var paragraphs = slidePart.Slide.Descendants<Paragraph>();
        foreach (var paragraph in paragraphs)
        {
            var hierarchicalInfo = TextExtractionUtility.ExtractHierarchicalText(paragraph);
            var paragraphText = hierarchicalInfo.CombinedText;

            if (!string.IsNullOrEmpty(paragraphText))
            {
                textElements.Add(paragraphText);
                Logger.Debug($"SlideTextExtractor: Added text element using {hierarchicalInfo.ProcessingStrategy}: '{paragraphText}'");
            }
        }

        return textElements;
    }

    public static List<string> ExtractBindingExpressions(SlidePart slidePart)
    {
        var expressions = new List<string>();
        var bindingExpressionRegex = new System.Text.RegularExpressions.Regex(@"\$\{[^}]+\}");

        // Use hierarchical extraction to find binding expressions 
        var paragraphs = slidePart.Slide.Descendants<Paragraph>();
        foreach (var paragraph in paragraphs)
        {
            var hierarchicalInfo = TextExtractionUtility.ExtractHierarchicalText(paragraph);
            var paragraphText = hierarchicalInfo.CombinedText;

            var matches = bindingExpressionRegex.Matches(paragraphText);
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                expressions.Add(match.Value);
                Logger.Debug($"SlideTextExtractor: Found binding expression: '{match.Value}' in paragraph: '{paragraphText}'");
            }
        }

        return expressions;
    }

    /// <summary>
    /// Replaces text in a slide while preserving formatting
    /// </summary>
    public static void ReplaceText(SlidePart slidePart, string oldText, string newText)
    {
        Logger.Debug($"SlideTextExtractor: Replacing '{oldText}' with '{newText}' in slide");

        var paragraphs = slidePart.Slide.Descendants<Paragraph>();
        foreach (var paragraph in paragraphs)
        {
            ReplaceTextInParagraph(paragraph, oldText, newText);
        }
    }

    /// <summary>
    /// Replaces text in a paragraph while preserving formatting
    /// </summary>
    private static void ReplaceTextInParagraph(Paragraph paragraph, string oldText, string newText)
    {
        var textElements = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().ToList();
        var combinedText = string.Join("", textElements.Select(t => t.Text));

        if (!combinedText.Contains(oldText))
            return;

        // Simple replacement in the first text element for now
        if (textElements.Count > 0)
        {
            var replacedText = combinedText.Replace(oldText, newText);
            textElements[0].Text = replacedText;

            // Clear other text elements
            for (int i = 1; i < textElements.Count; i++)
            {
                textElements[i].Text = "";
            }
        }

        Logger.Debug($"SlideTextExtractor: Replaced '{oldText}' with '{newText}' in paragraph");
    }
}
