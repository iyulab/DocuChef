using DocuChef.Presentation.Core;

namespace DocuChef.Presentation.Utils;

/// <summary>
/// Utilities for handling text in PowerPoint slides
/// </summary>
internal static class TextProcessingUtils
{
    /// <summary>
    /// Distributes text containing expressions across text elements
    /// </summary>
    public static void DistributeTextWithExpressions(
        List<SlideManager.RunInfo> runMap,
        string text,
        List<SlideManager.ExpressionRange> expressionRanges)
    {
        if (runMap.Count == 0)
            return;

        // Clear all text elements
        foreach (var run in runMap)
        {
            run.TextElement.Text = string.Empty;
        }

        // Find non-expression text segments
        var textSegments = new List<SlideManager.TextSegment>();
        int lastEnd = 0;

        foreach (var range in expressionRanges)
        {
            // Add segment before this expression
            if (range.Start > lastEnd)
            {
                textSegments.Add(new SlideManager.TextSegment
                {
                    Start = lastEnd,
                    End = range.Start - 1,
                    Text = text.Substring(lastEnd, range.Start - lastEnd)
                });
            }

            // Add the expression segment
            textSegments.Add(new SlideManager.TextSegment
            {
                Start = range.Start,
                End = range.End,
                Text = range.Expression,
                IsExpression = true
            });

            lastEnd = range.End + 1;
        }

        // Add any text after the last expression
        if (lastEnd < text.Length)
        {
            textSegments.Add(new SlideManager.TextSegment
            {
                Start = lastEnd,
                End = text.Length - 1,
                Text = text.Substring(lastEnd)
            });
        }

        // Now assign segments to runs
        int currentRun = 0;

        foreach (var segment in textSegments)
        {
            if (currentRun >= runMap.Count)
                currentRun = runMap.Count - 1; // Use last run if we run out

            // For expressions, try to keep them in a single run
            if (segment.IsExpression)
            {
                // If this run already has text and we have more runs available, go to next run
                if (runMap[currentRun].TextElement.Text.Length > 0 && currentRun < runMap.Count - 1)
                    currentRun++;

                runMap[currentRun].TextElement.Text += segment.Text;

                // Move to next run after an expression
                if (currentRun < runMap.Count - 1)
                    currentRun++;
            }
            else
            {
                // For non-expression text, add to current run
                runMap[currentRun].TextElement.Text += segment.Text;
            }
        }
    }

    /// <summary>
    /// Updates paragraph text while preserving the original run structure and formatting
    /// </summary>
    public static void UpdateParagraphTextPreservingRuns(
        D.Paragraph paragraph,
        List<D.Text> textElements,
        string originalText,
        string newText)
    {
        // Special case: if there's only one text element, update it directly
        if (textElements.Count == 1)
        {
            textElements[0].Text = newText;
            return;
        }

        // Create maps of the original run structure
        var runMap = new List<SlideManager.RunInfo>();
        int currentPosition = 0;

        foreach (var textElement in textElements)
        {
            if (textElement.Text == null)
                continue;

            int length = textElement.Text.Length;
            if (length > 0)
            {
                runMap.Add(new SlideManager.RunInfo
                {
                    TextElement = textElement,
                    StartPosition = currentPosition,
                    EndPosition = currentPosition + length - 1,
                    Run = textElement.Parent as D.Run
                });
                currentPosition += length;
            }
        }

        // If the text length hasn't changed, we can distribute it proportionally
        if (originalText.Length == newText.Length)
        {
            for (int i = 0; i < runMap.Count; i++)
            {
                var run = runMap[i];
                run.TextElement.Text = newText.Substring(run.StartPosition, run.EndPosition - run.StartPosition + 1);
            }
        }
        else
        {
            // For changed lengths, try to distribute based on expression boundaries
            // Find all expressions in the adjusted text
            var expressionRanges = SlideManager.FindExpressionRanges(newText);

            if (expressionRanges.Count > 0)
            {
                DistributeTextWithExpressions(runMap, newText, expressionRanges);
            }
            else
            {
                // Fallback: put all in first text element, clear others
                textElements[0].Text = newText;
                for (int i = 1; i < textElements.Count; i++)
                {
                    textElements[i].Text = string.Empty;
                }
            }
        }
    }

    /// <summary>
    /// Combines the text from multiple text elements
    /// </summary>
    public static string CombineTextElements(List<D.Text> textElements)
    {
        StringBuilder fullText = new StringBuilder();
        foreach (var textElement in textElements)
        {
            if (textElement.Text != null)
                fullText.Append(textElement.Text);
        }
        return fullText.ToString();
    }

    /// <summary>
    /// Updates the content of a TextBody element, preserving required structure
    /// </summary>
    public static void UpdateTextBodyContent(P.TextBody textBody, string content)
    {
        // Make sure body properties exist, based on standard PowerPoint defaults
        if (textBody.BodyProperties == null)
        {
            textBody.AppendChild(new D.BodyProperties());
        }

        // Make sure list style exists
        if (textBody.ListStyle == null)
        {
            textBody.AppendChild(new D.ListStyle());
        }

        // Clear existing paragraphs
        var paragraphs = textBody.Elements<D.Paragraph>().ToList();
        foreach (var para in paragraphs)
        {
            para.Remove();
        }

        // Create and add new paragraph
        var newParagraph = new D.Paragraph();
        var newRun = new D.Run();
        var newText = new D.Text() { Text = content };

        newRun.AppendChild(newText);
        newParagraph.AppendChild(newRun);
        textBody.AppendChild(newParagraph);
    }
}