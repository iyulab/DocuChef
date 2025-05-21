using DocuChef.Presentation.Handlers;
using DocuChef.Presentation.Models;
using DocuChef.Presentation.Utils;
using DollarSignEngine;

namespace DocuChef.Presentation.Core;

/// <summary>
/// Handles data binding to slide elements using DollarSignEngine
/// </summary>
internal class DataBinder
{
    /// <summary>
    /// Binds data to all text elements in a slide using slide context
    /// </summary>
    public static void BindDataWithContext(SlidePart slidePart, Models.SlideContext context, object data)
    {
        if (slidePart == null || context == null)
        {
            Logger.Debug("Cannot bind data: slidePart or context is null");
            return;
        }

        try
        {
            Logger.Debug($"Starting data binding for slide with context: {context.GetContextDescription()}");

            // Create DollarSignOptions with context-aware variable resolver
            var options = CreateDollarSignOptions(data);

            // Get snapshot of all text elements before we modify anything
            var allTextElements = CaptureTextElements(slidePart);
            LogTextElements(allTextElements);

            // First check for expressions that need to be processed at the Run level
            ProcessRunLevelExpressions(slidePart, context, options, allTextElements);

            // Then process expressions at the Paragraph level for split expressions
            ProcessTextElementsByParagraph(slidePart, context, options, allTextElements);

            // Finally, process any remaining individual text elements
            ProcessIndividualTextElements(slidePart, context, options, allTextElements);

            Logger.Debug("Data binding completed successfully");
        }
        catch (Exception ex)
        {
            Logger.Error($"Error binding data to slide: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Creates the DollarSignOptions for expression evaluation
    /// </summary>
    private static DollarSignOptions CreateDollarSignOptions(object data)
    {
        return DollarSignOptions.Default
            .WithDollarSignSyntax()
            .WithGlobalData(data)
            .WithErrorHandler((expr, ex) =>
            {
                Logger.Debug($"Expression error: '{expr}' - {ex.Message}");
                return string.Empty;  // Return empty on error
            });
    }

    /// <summary>
    /// Logs captured text elements for debugging
    /// </summary>
    private static void LogTextElements(List<TextElementInfo> textElements)
    {
        Logger.Debug($"Captured {textElements.Count} text elements for processing");

        for (int i = 0; i < textElements.Count; i++)
        {
            var textInfo = textElements[i];
            Logger.Debug($"Text element {i}: '{textInfo.Text}' (Parent type: {textInfo.TextElement.Parent?.GetType().Name})");
        }
    }

    /// <summary>
    /// Captures all text elements with their text values before any processing
    /// </summary>
    private static List<TextElementInfo> CaptureTextElements(SlidePart slidePart)
    {
        var result = new List<TextElementInfo>();
        var textElements = slidePart.Slide.Descendants<D.Text>().ToList();

        foreach (var element in textElements)
        {
            result.Add(new TextElementInfo
            {
                TextElement = element,
                Text = element.Text ?? string.Empty,
                IsProcessed = false,
                Shape = element.FindShape(),
                Run = element.Parent as D.Run
            });
        }

        return result;
    }

    /// <summary>
    /// Processes text elements at the Run level for intact expressions within a single Run
    /// </summary>
    private static void ProcessRunLevelExpressions(SlidePart slidePart, SlideContext context,
        DollarSignOptions options, List<TextElementInfo> allTextElements)
    {
        Logger.Debug("Processing Run-level expressions to preserve formatting");

        foreach (var textInfo in allTextElements)
        {
            string text = textInfo.Text;

            // Skip if already processed or empty
            if (textInfo.IsProcessed || string.IsNullOrEmpty(text))
                continue;

            // Check if this Run contains a complete expression
            if (ContainsCompleteExpression(text))
            {
                Logger.Debug($"Found complete expression in Run: '{text}'");

                // Check for index out of bounds first
                if (IsIndexOutOfBounds(text, context))
                {
                    HideElementOrShape(textInfo);
                    textInfo.IsProcessed = true;
                    continue;
                }

                ProcessExpression(textInfo, text, context, options);
            }
        }
    }

    /// <summary>
    /// Processes text elements grouped by their paragraphs to handle split expressions
    /// </summary>
    private static void ProcessTextElementsByParagraph(SlidePart slidePart, SlideContext context,
        DollarSignOptions options, List<TextElementInfo> allTextElements)
    {
        try
        {
            // Group text elements by their paragraph parent
            var paragraphs = slidePart.Slide.Descendants<D.Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                // Find all text elements that belong to this paragraph and aren't processed yet
                var paragraphTextElements = allTextElements
                    .Where(t => !t.IsProcessed && IsParentOrAncestor(t.TextElement, paragraph))
                    .ToList();

                if (paragraphTextElements.Count <= 0)
                    continue;

                // Check if any element contains expression markers
                if (!ContainsPotentialExpressions(paragraphTextElements))
                    continue;

                // Build combined text
                string combinedText = CombineTextElements(paragraphTextElements);
                Logger.Debug($"Combined paragraph text: '{combinedText}'");

                // Check if combined text contains potential expressions
                if (ContainsAnyExpressionMarkers(combinedText))
                {
                    // Check for index out of bounds in combined expression
                    if (IsIndexOutOfBounds(combinedText, context))
                    {
                        // Find the shape and hide it
                        var shape = paragraph.FindShape();
                        if (shape != null)
                        {
                            shape.Hide();
                            Logger.Debug($"Hiding shape due to index out of bounds in expression: '{combinedText}'");

                            // Mark all text elements in this paragraph as processed
                            MarkElementsAsProcessed(paragraphTextElements);
                            continue;
                        }
                    }

                    ProcessParagraphExpression(paragraph, paragraphTextElements, combinedText, context, options);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error in ProcessTextElementsByParagraph: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Processes individual text elements that weren't handled as part of paragraphs
    /// </summary>
    private static void ProcessIndividualTextElements(SlidePart slidePart, SlideContext context,
        DollarSignOptions options, List<TextElementInfo> allTextElements)
    {
        // Process remaining (unprocessed) text elements individually
        foreach (var textInfo in allTextElements.Where(t => !t.IsProcessed))
        {
            string originalText = textInfo.Text;
            if (string.IsNullOrEmpty(originalText))
                continue;

            // Skip elements already processed as part of a paragraph
            if (textInfo.IsProcessed)
                continue;

            // Check if text contains potential bindings
            if (ContainsAnyExpressionMarkers(originalText))
            {
                try
                {
                    // Check for index out of bounds first
                    if (IsIndexOutOfBounds(originalText, context))
                    {
                        HideElementOrShape(textInfo);
                        continue;
                    }

                    ProcessExpression(textInfo, originalText, context, options);
                }
                catch (Exception ex)
                {
                    Logger.Error($"Error evaluating expression '{originalText}': {ex.Message}", ex);
                }
            }
        }
    }

    /// <summary>
    /// Processes an expression and updates the text element or shape
    /// </summary>
    private static void ProcessExpression(TextElementInfo textInfo, string text, SlideContext context, DollarSignOptions options)
    {
        try
        {
            // Handle Image expressions
            if (text.Contains($"ppt.{nameof(PPTMethods.Image)}"))
            {
                var shape = textInfo.Shape;
                if (shape != null)
                {
                    string evaluatedText = EvaluateExpression(text, context, options);
                    ImageHandler.Process(shape, evaluatedText);
                    textInfo.IsProcessed = true;
                    return;
                }
            }

            // Process regular text expressions
            string newText = EvaluateExpression(text, context, options);

            // Only update if text changed
            if (newText != text)
            {
                Logger.Debug($"Binding text: '{text}' -> '{newText}'");
                textInfo.TextElement.Text = newText;
                textInfo.IsProcessed = true;
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating expression '{text}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Processes an expression for a paragraph and updates the text elements
    /// </summary>
    private static void ProcessParagraphExpression(D.Paragraph paragraph, List<TextElementInfo> paragraphTextElements,
        string combinedText, SlideContext context, DollarSignOptions options)
    {
        try
        {
            // If this is an image expression, process it differently
            if (combinedText.Contains($"ppt.{nameof(PPTMethods.Image)}"))
            {
                var shape = paragraph.FindShape();
                if (shape != null)
                {
                    string evaluatedText = EvaluateExpression(combinedText, context, options);
                    ImageHandler.Process(shape, evaluatedText);
                    MarkElementsAsProcessed(paragraphTextElements);
                    return;
                }
            }

            // Normal text binding
            string newText = EvaluateExpression(combinedText, context, options);

            // Only update if text changed
            if (newText != combinedText)
            {
                Logger.Debug($"Binding combined text: '{combinedText}' -> '{newText}'");
                UpdateParagraphText(paragraphTextElements, newText);
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error evaluating combined expression '{combinedText}': {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Evaluates an expression using DollarSignEngine
    /// </summary>
    private static string EvaluateExpression(string expression, SlideContext context, DollarSignOptions options)
    {
        var data = context.GetData();
        var pattern = @"\$\{(.*?)(" + @"\" + PowerPointOptions.Current.HierarchyDelimiter + @")(.*?)\}";
        var expressionReform = System.Text.RegularExpressions.Regex.Replace(expression, pattern, m =>
        {
            return "${" + m.Groups[1].Value + "__" + m.Groups[3].Value + "}";
        });
        var r = DollarSign.Eval(expressionReform, data, options);
        return r;
    }

    /// <summary>
    /// Updates the text elements in a paragraph after evaluating an expression
    /// </summary>
    private static void UpdateParagraphText(List<TextElementInfo> paragraphTextElements, string newText)
    {
        // Update the first text element with the entire result
        var firstTextInfo = paragraphTextElements.First();
        if (firstTextInfo.TextElement != null)
        {
            firstTextInfo.TextElement.Text = newText;
            firstTextInfo.IsProcessed = true;

            // Clear all other text elements in this paragraph
            for (int i = 1; i < paragraphTextElements.Count; i++)
            {
                var textInfo = paragraphTextElements[i];
                if (textInfo.TextElement != null)
                {
                    textInfo.TextElement.Text = string.Empty;
                    textInfo.IsProcessed = true;
                }
            }
        }
    }

    /// <summary>
    /// Hides a text element or its containing shape
    /// </summary>
    private static void HideElementOrShape(TextElementInfo textInfo)
    {
        var shape = textInfo.Shape;
        if (shape != null)
        {
            shape.Hide();
            Logger.Debug($"Hiding shape for text with out of bounds index: '{textInfo.Text}'");
        }
        else
        {
            textInfo.TextElement.Hide();
            Logger.Debug($"Hiding text element with out of bounds index: '{textInfo.Text}'");
        }
    }

    /// <summary>
    /// Marks all text elements in a collection as processed
    /// </summary>
    private static void MarkElementsAsProcessed(List<TextElementInfo> textElements)
    {
        foreach (var textInfo in textElements)
        {
            textInfo.IsProcessed = true;
        }
    }

    /// <summary>
    /// Combines the text from multiple text elements
    /// </summary>
    private static string CombineTextElements(List<TextElementInfo> textElements)
    {
        StringBuilder fullText = new StringBuilder();
        foreach (var textInfo in textElements)
        {
            fullText.Append(textInfo.Text);
        }
        return fullText.ToString();
    }

    /// <summary>
    /// Checks if any text element in a collection contains potential expressions
    /// </summary>
    private static bool ContainsPotentialExpressions(List<TextElementInfo> textElements)
    {
        foreach (var textInfo in textElements)
        {
            string content = textInfo.Text;
            if (!string.IsNullOrEmpty(content) &&
                (content.Contains("${") || content.Contains("$") ||
                 content.Contains("{") || content.Contains('[')))
            {
                return true;
            }
        }
        return false;
    }

    /// <summary>
    /// Checks if a text contains any expression markers
    /// </summary>
    private static bool ContainsAnyExpressionMarkers(string text)
    {
        return text.Contains("{") || text.Contains("$") || text.Contains('[');
    }

    /// <summary>
    /// Checks if a text element contains a complete expression (${...})
    /// </summary>
    private static bool ContainsCompleteExpression(string text)
    {
        if (string.IsNullOrEmpty(text))
            return false;

        // Check for ${...} pattern
        int openIndex = text.IndexOf("${");
        if (openIndex >= 0)
        {
            int closeIndex = text.IndexOf("}", openIndex);
            if (closeIndex > openIndex)
            {
                // Make sure there's no other ${ before the closing }
                string between = text.Substring(openIndex + 2, closeIndex - openIndex - 2);
                return !between.Contains("${");
            }
        }

        // Check for $..$ pattern
        openIndex = text.IndexOf("$");
        if (openIndex >= 0 && openIndex < text.Length - 1)
        {
            int closeIndex = text.IndexOf("$", openIndex + 1);
            if (closeIndex > openIndex)
            {
                // Make sure there's no other $ in between
                string between = text.Substring(openIndex + 1, closeIndex - openIndex - 1);
                return !between.Contains("$");
            }
        }

        return false;
    }

    /// <summary>
    /// Checks if an element has the specified parent somewhere in its hierarchy
    /// </summary>
    private static bool IsParentOrAncestor(OpenXmlElement element, OpenXmlElement potentialParent)
    {
        if (element == null || potentialParent == null)
            return false;

        OpenXmlElement current = element.Parent;
        while (current != null)
        {
            if (current == potentialParent)
                return true;

            current = current.Parent;
        }

        return false;
    }

    /// <summary>
    /// Checks if an expression contains an index that is out of bounds
    /// </summary>
    private static bool IsIndexOutOfBounds(string text, SlideContext context)
    {
        try
        {
            if (text.Contains(context.CollectionName))
            {
                var ndx = text.GetBetween(context.CollectionName + "[", "]");
                if (int.TryParse(ndx, out var index) && index >= context.TotalItems)
                {
                    Logger.Debug($"Index out of bounds: {context.CollectionName}[{index}], TotalItems: {context.TotalItems}");
                    return true;
                }
            }

            return false;
        }
        catch (Exception ex)
        {
            Logger.Debug($"Error checking index bounds: {ex.Message}");
            return false;
        }
    }
}

/// <summary>
/// Helper class to track text elements and their processing state
/// </summary>
internal class TextElementInfo
{
    /// <summary>
    /// The text element itself
    /// </summary>
    public D.Text TextElement { get; set; }

    /// <summary>
    /// The text content captured at the beginning
    /// </summary>
    public string Text { get; set; }

    /// <summary>
    /// Whether this element has been processed
    /// </summary>
    public bool IsProcessed { get; set; }

    /// <summary>
    /// The shape containing this text element (if found)
    /// </summary>
    public P.Shape Shape { get; set; }

    /// <summary>
    /// The Run element that contains this text element (if available)
    /// </summary>
    public D.Run Run { get; set; }
}