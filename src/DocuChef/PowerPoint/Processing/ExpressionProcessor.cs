using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Unified processor for handling expressions in templates
/// </summary>
internal class ExpressionProcessor
{
    private readonly IExpressionEvaluator _evaluator;
    private readonly PowerPointContext _context;
    private static readonly Regex ExpressionPattern = new(@"\$\{([^{}]+)\}", RegexOptions.Compiled);
    private static readonly Regex ArrayIndexPattern = new(@"(\w+)\[(\d+)\]", RegexOptions.Compiled);

    /// <summary>
    /// Initialize expression processor
    /// </summary>
    public ExpressionProcessor(IExpressionEvaluator evaluator, PowerPointContext context)
    {
        _evaluator = evaluator ?? throw new ArgumentNullException(nameof(evaluator));
        _context = context ?? throw new ArgumentNullException(nameof(context));
    }

    /// <summary>
    /// Process expressions in text with provided variables
    /// </summary>
    public string ProcessExpressions(string text, Dictionary<string, object> variables)
    {
        if (string.IsNullOrEmpty(text) || !ExpressionHelper.ContainsExpressions(text))
            return text;

        // Debug for context
        Logger.Debug($"Processing expressions in text: {text.Substring(0, Math.Min(text.Length, 50))}...");
        if (Logger.MinimumLevel <= Logger.LogLevel.Debug && _context.HierarchicalIndices.Count > 0)
        {
            Logger.Debug("Current hierarchical indices:");
            foreach (var kvp in _context.HierarchicalIndices.OrderBy(x => x.Key))
            {
                Logger.Debug($"  {kvp.Key} = {kvp.Value}");
            }
        }

        return ExpressionPattern.Replace(text, match =>
        {
            try
            {
                string expression = match.Value;
                string expressionContent = match.Groups[1].Value;

                // Check for array index references that need to be adjusted
                if (ContainsArrayReference(expressionContent))
                {
                    expression = AdjustArrayIndices(expression, expressionContent);
                }

                // Process the adjusted expression
                var result = _evaluator.EvaluateCompleteExpression(expression, variables);

                // Handle null results
                if (result == null)
                {
                    Logger.Warning($"Expression '{expression}' evaluated to null");
                    return "";
                }

                return result.ToString();
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error evaluating expression '{match.Value}': {ex.Message}");
                return ""; // Replace with empty string on error
            }
        });
    }

    /// <summary>
    /// Check if expression contains array references
    /// </summary>
    private bool ContainsArrayReference(string expression)
    {
        return expression.Contains("[") && expression.Contains("]");
    }

    /// <summary>
    /// Adjust array indices based on hierarchical context
    /// </summary>
    private string AdjustArrayIndices(string fullExpression, string expressionContent)
    {
        // Don't adjust if no hierarchical indices available
        if (_context.HierarchicalIndices == null || _context.HierarchicalIndices.Count == 0)
            return fullExpression;

        try
        {
            // Match array references like: Collection[n]
            var arrayMatches = ArrayIndexPattern.Matches(expressionContent);

            if (arrayMatches.Count == 0)
                return fullExpression;

            string adjustedExpression = expressionContent;
            bool madeAdjustments = false;

            foreach (Match arrayMatch in arrayMatches)
            {
                if (arrayMatch.Groups.Count < 3)
                    continue;

                string arrayName = arrayMatch.Groups[1].Value;
                string indexString = arrayMatch.Groups[2].Value;

                // Skip adjustment if we can't parse the index
                if (!int.TryParse(indexString, out int arrayIndex))
                    continue;

                // Check hierarchical context for nested paths
                string contextKey = FindBestMatchingContextKey(arrayName);
                if (!string.IsNullOrEmpty(contextKey) && _context.HierarchicalIndices.TryGetValue(contextKey, out int baseIndex))
                {
                    // Calculate adjusted index: template index + base index from context
                    int adjustedIndex = arrayIndex + baseIndex;

                    // Replace the index in the expression
                    string originalRef = $"{arrayName}[{indexString}]";
                    string adjustedRef = $"{arrayName}[{adjustedIndex}]";

                    adjustedExpression = adjustedExpression.Replace(originalRef, adjustedRef);
                    madeAdjustments = true;

                    Logger.Debug($"Adjusted array index: {originalRef} -> {adjustedRef} (base: {baseIndex})");
                }
                else
                {
                    Logger.Debug($"No context index found for array: {arrayName}");
                }
            }

            return madeAdjustments ? "${" + adjustedExpression + "}" : fullExpression;
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error adjusting array indices in '{expressionContent}': {ex.Message}");
            return fullExpression;
        }
    }

    /// <summary>
    /// Find the best matching context key for an array reference
    /// </summary>
    private string FindBestMatchingContextKey(string arrayName)
    {
        // Direct match for the array name
        if (_context.HierarchicalIndices.ContainsKey(arrayName))
        {
            Logger.Debug($"Found direct match for array name: {arrayName}");
            return arrayName;
        }

        // Direct array reference match (e.g. Items[0])
        var arrayPattern = new System.Text.RegularExpressions.Regex($@"{arrayName}\[(\d+)\]");
        foreach (var key in _context.HierarchicalIndices.Keys)
        {
            if (arrayPattern.IsMatch(key))
            {
                Logger.Debug($"Found direct array reference match: {key}");
                return key;
            }
        }

        // Look for underscore paths ending with this array name
        foreach (var key in _context.HierarchicalIndices.Keys)
        {
            // Check for underscore format: Parent_ArrayName
            if (key.EndsWith($"_{arrayName}", StringComparison.OrdinalIgnoreCase) ||
                key.Equals(arrayName, StringComparison.OrdinalIgnoreCase))
            {
                Logger.Debug($"Found matching context key: {key} for array: {arrayName}");
                return key;
            }
        }

        // Try to find more complex matches
        // This is a fallback for cases where the array might be part of a longer path
        var possibleMatches = _context.HierarchicalIndices.Keys
            .Where(k => k.Contains(arrayName))
            .OrderByDescending(k => k.Split('_').Length) // Prioritize more specific paths
            .ToList();

        if (possibleMatches.Any())
        {
            var bestMatch = possibleMatches.First();
            Logger.Debug($"Found possible matching context key: {bestMatch} for array: {arrayName}");
            return bestMatch;
        }

        // Check if we have a collection with item or items (common naming pattern)
        if (arrayName.Equals("Item", StringComparison.OrdinalIgnoreCase) ||
            arrayName.Equals("Items", StringComparison.OrdinalIgnoreCase))
        {
            // Look for any keys that might contain Items or Item
            var itemMatches = _context.HierarchicalIndices.Keys
                .Where(k => k.Contains("Item", StringComparison.OrdinalIgnoreCase))
                .OrderByDescending(k => k.Length) // Prioritize longer, more specific matches
                .ToList();

            if (itemMatches.Any())
            {
                var bestMatch = itemMatches.First();
                Logger.Debug($"Found item collection match: {bestMatch} for array: {arrayName}");
                return bestMatch;
            }
        }

        Logger.Debug($"No matching context key found for array: {arrayName}");
        return null;
    }

    /// <summary>
    /// Process expressions in a shape
    /// </summary>
    public bool ProcessShapeExpressions(P.Shape shape, Dictionary<string, object> variables)
    {
        if (shape.TextBody == null)
            return false;

        bool hasChanges = false;
        bool hasArrayReferences = false;
        bool hasInvalidArrayReferences = false;

        // Check for array references before processing
        string originalText = GetShapeTextContent(shape);
        hasArrayReferences = ContainsArrayReference(originalText);

        // Check for array references that might be out of bounds
        if (hasArrayReferences)
        {
            hasInvalidArrayReferences = CheckForInvalidArrayReferences(originalText, variables);

            // If there are invalid array references, hide the shape immediately
            if (hasInvalidArrayReferences)
            {
                Logger.Debug($"Shape '{shape.GetShapeName()}' contains invalid array references, hiding");
                ShapeHelper.HideShape(shape);
                return true;
            }
        }

        // Check if expressions span across runs
        var crossRunExpressions = TextProcessingHelper.ContainsCrossRunExpressions(shape);

        if (crossRunExpressions)
        {
            // Process expressions that span across multiple runs
            foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>().ToList())
            {
                // Reconstruct paragraph text
                var (paragraphText, runMappings) = TextProcessingHelper.ReconstructParagraphText(paragraph);

                if (!ExpressionHelper.ContainsExpressions(paragraphText))
                    continue;

                // Process expressions
                string processedText = ProcessExpressions(paragraphText, variables);
                if (processedText == paragraphText)
                    continue;

                // Map processed text back to runs
                if (TextProcessingHelper.MapProcessedTextToRuns(paragraph, runMappings, processedText))
                    hasChanges = true;
            }
        }
        else
        {
            // Process individual runs
            foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
            {
                foreach (var run in paragraph.Elements<A.Run>())
                {
                    var textElement = run.GetFirstChild<A.Text>();
                    if (textElement == null || !ExpressionHelper.ContainsExpressions(textElement.Text))
                        continue;

                    string processedText = ProcessExpressions(textElement.Text, variables);
                    if (processedText != textElement.Text)
                    {
                        textElement.Text = processedText;
                        hasChanges = true;
                    }
                }
            }
        }

        // If text was processed, check if we should hide this shape
        if (hasChanges || hasArrayReferences)
        {
            // Get the processed text content
            string processedText = GetShapeTextContent(shape);

            // Check if the shape should be hidden after processing
            bool shouldHide = ShouldHideShapeWithEmptyContent(shape) ||
                             (hasArrayReferences && IsOutOfBoundsArrayShape(shape, processedText));

            if (shouldHide)
            {
                ShapeHelper.HideShape(shape);
                Logger.Debug($"Hiding shape '{shape.GetShapeName()}' due to empty/invalid content after processing");
                return true;
            }
        }

        return hasChanges;
    }

    /// <summary>
    /// Check if a shape contains invalid array references that exceed collection size
    /// </summary>
    private bool CheckForInvalidArrayReferences(string text, Dictionary<string, object> variables)
    {
        try
        {
            // Find all array references in the text
            var matches = ArrayIndexPattern.Matches(text);
            if (matches.Count == 0)
                return false;

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                string arrayName = match.Groups[1].Value;
                string indexStr = match.Groups[2].Value;

                if (!int.TryParse(indexStr, out int arrayIndex))
                    continue;

                // Check if the array exists
                if (!variables.TryGetValue(arrayName, out var arrayObj) || arrayObj == null)
                    continue;

                // Check collection size
                int collectionSize = DataNavigationHelper.GetCollectionCount(arrayObj);

                // Check if index is valid
                bool validIndex = false;

                // First check context indices (for collection foreach iteration)
                if (_context.HierarchicalIndices.TryGetValue($"{arrayName}[{arrayIndex}]", out int actualIndex))
                {
                    // If the mapped index is -1, it's specifically marked as invalid
                    if (actualIndex == -1)
                    {
                        Logger.Debug($"Array reference {arrayName}[{arrayIndex}] mapped to invalid index (-1)");
                        return true;
                    }

                    // Check if the mapped index is valid
                    validIndex = actualIndex >= 0 && actualIndex < collectionSize;
                    Logger.Debug($"Array reference {arrayName}[{arrayIndex}] mapped to {actualIndex}, valid: {validIndex}");
                }
                else
                {
                    // Direct index check
                    validIndex = arrayIndex >= 0 && arrayIndex < collectionSize;
                    Logger.Debug($"Direct array reference {arrayName}[{arrayIndex}], valid: {validIndex}");
                }

                if (!validIndex)
                {
                    Logger.Debug($"Invalid array reference detected: {arrayName}[{arrayIndex}], collection size: {collectionSize}");
                    return true;
                }
            }

            return false;
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error checking array references: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Get concatenated text content from a shape
    /// </summary>
    private string GetShapeTextContent(P.Shape shape)
    {
        if (shape.TextBody == null)
            return string.Empty;

        var textBuilder = new StringBuilder();

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement != null && !string.IsNullOrEmpty(textElement.Text))
                {
                    textBuilder.Append(textElement.Text);
                }
            }
            textBuilder.Append("\n");
        }

        return textBuilder.ToString();
    }

    /// <summary>
    /// Check if a shape contains array references that are out of bounds after processing
    /// </summary>
    private bool IsOutOfBoundsArrayShape(P.Shape shape, string processedText)
    {
        // If the processed text is empty or just whitespace, the shape should be hidden
        if (string.IsNullOrWhiteSpace(processedText))
            return true;

        // Check for any remaining ${...} expressions that weren't processed
        if (ExpressionHelper.ContainsExpressions(processedText))
        {
            // Extract the array references from remaining expressions
            var matches = ArrayIndexPattern.Matches(processedText);
            if (matches.Count > 0)
            {
                // If there are array references in unprocessed expressions, they might be out of bounds
                Logger.Debug($"Shape contains unprocessed array references: {processedText}");
                return true;
            }
        }

        // Check for "empty" indicators that might have been added by the expression evaluator
        if (processedText.Contains("[Error:") ||
            processedText.Contains("[Index out of bounds]") ||
            processedText.Contains("[]") ||
            processedText.Contains("undefined") ||
            processedText.Contains("null"))
        {
            Logger.Debug($"Shape contains error indicators after processing: {processedText}");
            return true;
        }

        return false;
    }

    /// <summary>
    /// Determines if a shape should be hidden after processing
    /// </summary>
    private bool ShouldHideShapeWithEmptyContent(P.Shape shape)
    {
        // This is a heuristic to detect shapes that should be hidden
        // because they contain references to out-of-bounds array items
        bool isEmpty = true;
        bool hasArrayPatterns = false;
        bool hasErrorIndicators = false;

        foreach (var paragraph in shape.TextBody.Elements<A.Paragraph>())
        {
            foreach (var run in paragraph.Elements<A.Run>())
            {
                var textElement = run.GetFirstChild<A.Text>();
                if (textElement != null)
                {
                    string text = textElement.Text;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        isEmpty = false;

                        // Check for error indicators
                        if (text.Contains("[Error:") ||
                            text.Contains("[Index out of bounds]") ||
                            text.Contains("undefined") ||
                            text.Contains("null"))
                        {
                            hasErrorIndicators = true;
                        }
                    }

                    // Check for remnants of array patterns that weren't processed
                    // This might indicate a reference to an item that's out of bounds
                    if (text.Contains("[") && text.Contains("]"))
                    {
                        hasArrayPatterns = true;
                    }
                }
            }
        }

        // Hide if one of these conditions is true:
        // 1. The shape is empty after processing
        // 2. The shape has array patterns but is empty (likely out-of-bounds references)
        // 3. The shape contains error indicators
        return isEmpty || (hasArrayPatterns && isEmpty) || hasErrorIndicators;
    }
}

/// <summary>
/// Helper class for text processing in OpenXML documents
/// </summary>
internal static class TextProcessingHelper
{
    /// <summary>
    /// Check if expressions span across runs
    /// </summary>
    public static bool ContainsCrossRunExpressions(P.Shape shape)
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

    /// <summary>
    /// Reconstruct complete paragraph text
    /// </summary>
    public static (string Text, List<(A.Run Run, int StartPos, int Length)> RunMappings) ReconstructParagraphText(A.Paragraph paragraph)
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
    public static bool MapProcessedTextToRuns(A.Paragraph paragraph, List<(A.Run Run, int StartPos, int Length)> runMappings, string processedText)
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
    private static void DistributeTextAcrossRuns(List<(A.Run Run, int StartPos, int Length)> runMappings, string processedText)
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
}