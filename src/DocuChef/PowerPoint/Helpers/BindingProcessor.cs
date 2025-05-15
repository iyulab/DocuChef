using DocuChef.PowerPoint.Processing;

namespace DocuChef.PowerPoint.Helpers;

/// <summary>
/// Helper class for processing binding with expression evaluation
/// </summary>
internal class BindingProcessor
{
    private readonly PowerPointProcessor _processor;
    private readonly Dictionary<string, object> _variables;

    /// <summary>
    /// Initialize binding processor
    /// </summary>
    public BindingProcessor(PowerPointProcessor processor, Dictionary<string, object> variables)
    {
        _processor = processor;
        _variables = variables ?? new Dictionary<string, object>();
    }
    
    /// <summary>
    /// Process individual runs for expressions
    /// </summary>
    private bool ProcessIndividualRuns(P.Shape shape)
    {
        bool hasChanges = false;

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || !ExpressionHelper.ContainsExpressions(textElement.Text))
                    continue;

                string processedText = ExpressionHelper.ProcessExpressions(textElement.Text, _processor, _variables);
                if (processedText != textElement.Text)
                {
                    textElement.Text = processedText;
                    hasChanges = true;
                }
            }
        }

        return hasChanges;
    }

    /// <summary>
    /// Process expressions that span across multiple runs
    /// </summary>
    private bool ProcessCrossRunExpressions(P.Shape shape)
    {
        bool hasChanges = false;

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
        {
            // Reconstruct paragraph text
            var (paragraphText, runMappings) = ReconstructParagraphText(paragraph);

            if (!ExpressionHelper.ContainsExpressions(paragraphText))
                continue;

            // Process expressions
            string processedText = ExpressionHelper.ProcessExpressions(paragraphText, _processor, _variables);
            if (processedText == paragraphText)
                continue;

            // Map processed text back to runs
            if (MapProcessedTextToRuns(paragraph, runMappings, processedText))
                hasChanges = true;
        }

        return hasChanges;
    }

    /// <summary>
    /// Reconstruct complete paragraph text
    /// </summary>
    private (string Text, List<(A.Run Run, int StartPos, int Length)> RunMappings) ReconstructParagraphText(A.Paragraph paragraph)
    {
        var sb = new StringBuilder();
        var runMappings = new List<(A.Run Run, int StartPos, int Length)>();

        foreach (var run in paragraph.Elements<A.Run>())
        {
            var textElement = run.GetFirstChild<A.Text>();
            if (textElement != null && !string.IsNullOrEmpty(textElement.Text))
            {
                int startPos = sb.Length;
                string text = textElement.Text;
                sb.Append(text);
                runMappings.Add((run, startPos, text.Length));
            }
        }

        return (sb.ToString(), runMappings);
    }

    /// <summary>
    /// Map processed text back to runs
    /// </summary>
    private bool MapProcessedTextToRuns(A.Paragraph paragraph, List<(A.Run Run, int StartPos, int Length)> runMappings, string processedText)
    {
        if (runMappings.Count == 0)
            return false;

        // Simple case: single run
        if (runMappings.Count == 1)
        {
            var textElement = runMappings[0].Run.GetFirstChild<A.Text>();
            if (textElement != null)
            {
                textElement.Text = processedText;
                return true;
            }
            return false;
        }

        // Complex case: distribute text across runs
        DistributeTextAcrossRuns(runMappings, processedText);
        return true;
    }

    /// <summary>
    /// Distribute text across multiple runs
    /// </summary>
    private void DistributeTextAcrossRuns(List<(A.Run Run, int StartPos, int Length)> runMappings, string processedText)
    {
        // Calculate distribution ratio
        double ratio = (double)processedText.Length / runMappings.Sum(r => r.Length);
        int currentPos = 0;

        for (int i = 0; i < runMappings.Count; i++)
        {
            var runInfo = runMappings[i];
            runInfo.Run.RemoveAllChildren<A.Text>();

            // Calculate text portion for this run
            string runText;
            if (i == runMappings.Count - 1)
            {
                // Last run gets remaining text
                runText = currentPos < processedText.Length
                    ? processedText.Substring(currentPos)
                    : string.Empty;
            }
            else
            {
                // Distribute proportionally
                int newLength = (int)Math.Ceiling(runInfo.Length * ratio);
                newLength = Math.Min(newLength, processedText.Length - currentPos);

                runText = newLength > 0 && currentPos < processedText.Length
                    ? processedText.Substring(currentPos, newLength)
                    : string.Empty;

                currentPos += runText.Length;
            }

            runInfo.Run.AppendChild(new A.Text(runText));
        }
    }

    /// <summary>
    /// Check if expressions span across runs
    /// </summary>
    private bool ContainsCrossRunExpressions(P.Shape shape)
    {
        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            bool hasOpenBrace = false;

            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement == null || string.IsNullOrEmpty(textElement.Text))
                    continue;

                string text = textElement.Text;

                // Check for complete expressions
                if (text.Contains("${") && text.Contains("}"))
                    continue;

                // Check for partial expressions
                if (text.Contains("${"))
                    hasOpenBrace = true;

                if (text.Contains("}") && hasOpenBrace)
                    return true;
            }

            if (hasOpenBrace)
                return true;
        }

        return false;
    }
}