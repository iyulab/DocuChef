namespace DocuChef.Presentation.Core;

/// <summary>
/// Adjusts expression indices in text based on slide context
/// </summary>
internal class ExpressionAdjuster
{
    /// <summary>
    /// Adjusts array indices in expressions based on the current slide context's offset
    /// </summary>
    /// <param name="originalText">Original text containing expressions</param>
    /// <param name="context">Slide context with offset information</param>
    /// <returns>Adjusted text with updated indices</returns>
    public static string AdjustExpressionIndices(string originalText, Models.SlideContext context)
    {
        if (string.IsNullOrEmpty(originalText) || context == null || context.Offset == 0)
            return originalText;

        Logger.Debug($"Adjusting expression indices in text: '{originalText}' with context {context.CollectionName}, offset {context.Offset}, isGroup: {context.IsGroupContext}, itemsInGroup: {context.ItemsInGroup}");

        // Create builder for result
        StringBuilder result = new StringBuilder();
        int currentPos = 0;
        int nextExprStart;

        // Get the hierarchy delimiter
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;
        bool isNestedContext = context.CollectionName.Contains(delimiter);

        // Handle hierarchical paths if delimiter is present
        string[] segments = isNestedContext
            ? context.CollectionName.Split(new[] { delimiter }, StringSplitOptions.None)
            : new[] { context.CollectionName };

        // Find all expressions starting with "${" and ending with "}"
        while ((nextExprStart = originalText.IndexOf("${", currentPos)) != -1)
        {
            // Append text before the expression
            result.Append(originalText.Substring(currentPos, nextExprStart - currentPos));

            // Find the end of the expression
            int exprEnd = originalText.IndexOf("}", nextExprStart);
            if (exprEnd == -1)
            {
                // No closing bracket found, simply copy the rest of the text
                result.Append(originalText.Substring(nextExprStart));
                break;
            }

            // Extract the full expression
            string fullExpr = originalText.Substring(nextExprStart, exprEnd - nextExprStart + 1);
            bool expressionAdjusted = false;

            // Process expression based on its type and context
            (bool isAdjusted, string adjustedExpr) = ProcessExpression(fullExpr, context, segments, isNestedContext);

            if (isAdjusted)
            {
                result.Append(adjustedExpr);
                expressionAdjusted = true;
            }
            else
            {
                result.Append(fullExpr);
            }

            // Move to the next position
            currentPos = exprEnd + 1;
        }

        // Append any remaining text
        if (currentPos < originalText.Length)
        {
            result.Append(originalText.Substring(currentPos));
        }

        return result.ToString();
    }

    /// <summary>
    /// Processes an expression and adjusts it based on the context
    /// </summary>
    private static (bool isAdjusted, string adjustedExpr) ProcessExpression(
        string expression, Models.SlideContext context, string[] segments, bool isNestedContext)
    {
        // Check for direct collection name match
        string searchKey = context.CollectionName + "[";
        if (expression.Contains(searchKey))
        {
            string adjustedExpr = AdjustDirectCollectionIndices(expression, searchKey, context);
            return (true, adjustedExpr);
        }

        // If this is a hierarchical context, check for segment matches
        if (isNestedContext)
        {
            // Check each segment separately, starting from the most specific (last)
            for (int i = segments.Length - 1; i >= 0; i--)
            {
                string segmentKey = segments[i] + "[";
                if (expression.Contains(segmentKey))
                {
                    string adjustedExpr = AdjustHierarchicalSegmentIndices(expression, segmentKey, context, i);
                    return (true, adjustedExpr);
                }
            }
        }

        // Check for Item[index] in group contexts
        if (context.IsGroupContext && expression.Contains("Item["))
        {
            string adjustedExpr = AdjustGroupItemIndices(expression, context);
            return (true, adjustedExpr);
        }

        // No adjustment needed
        return (false, expression);
    }

    /// <summary>
    /// Adjusts indices for direct collection name matches
    /// </summary>
    private static string AdjustDirectCollectionIndices(string expression, string searchKey, Models.SlideContext context)
    {
        int keyPos = expression.IndexOf(searchKey);
        int keyLength = searchKey.Length;

        // Extract the index
        int indexStart = keyPos + keyLength;
        int indexEnd = expression.IndexOf("]", indexStart);

        if (indexEnd != -1)
        {
            string indexStr = expression.Substring(indexStart, indexEnd - indexStart);

            if (int.TryParse(indexStr, out int originalIndex))
            {
                // Calculate adjusted index
                int adjustedIndex = ApplyDirectIndexAdjustment(originalIndex, context);

                Logger.Debug($"Adjusted direct index in expression: {originalIndex} -> {adjustedIndex}");

                // Create the adjusted expression
                string beforeIndex = expression.Substring(0, indexStart);
                string afterIndex = expression.Substring(indexEnd);
                return beforeIndex + adjustedIndex + afterIndex;
            }
        }

        // Return original if no adjustment was made
        return expression;
    }

    /// <summary>
    /// Adjusts indices for hierarchical segment matches
    /// </summary>
    private static string AdjustHierarchicalSegmentIndices(string expression, string segmentKey, Models.SlideContext context, int segmentIndex)
    {
        int keyPos = expression.IndexOf(segmentKey);
        int keyLength = segmentKey.Length;

        // Extract the index
        int indexStart = keyPos + keyLength;
        int indexEnd = expression.IndexOf("]", indexStart);

        if (indexEnd != -1)
        {
            string indexStr = expression.Substring(indexStart, indexEnd - indexStart);

            if (int.TryParse(indexStr, out int originalIndex))
            {
                // Get the offset for this segment level
                int segmentOffset = GetSegmentOffset(context, segmentIndex);

                // Apply the offset
                int adjustedIndex = originalIndex + segmentOffset;

                Logger.Debug($"Adjusted hierarchical segment index in expression: {originalIndex} -> {adjustedIndex} (segment '{segmentKey}', offset {segmentOffset})");

                // Create the adjusted expression
                string beforeIndex = expression.Substring(0, indexStart);
                string afterIndex = expression.Substring(indexEnd);
                return beforeIndex + adjustedIndex + afterIndex;
            }
        }

        // Return original if no adjustment was made
        return expression;
    }

    /// <summary>
    /// Gets the appropriate offset for a segment level
    /// </summary>
    private static int GetSegmentOffset(Models.SlideContext context, int segmentIndex)
    {
        // For parent segments, use parent context offsets if available
        if (segmentIndex < context.HierarchyLevel && context.ParentContext != null)
        {
            return context.ParentContext.Offset;
        }
        else if (context.LevelOffsets.TryGetValue(segmentIndex, out int levelOffset))
        {
            return levelOffset;
        }

        // Default to current context offset
        return context.Offset;
    }

    /// <summary>
    /// Adjusts indices for Item[index] in group contexts
    /// </summary>
    private static string AdjustGroupItemIndices(string expression, Models.SlideContext context)
    {
        // For group contexts, Item[0], Item[1], etc. should not be adjusted because they refer
        // to items within the current group, not to global collection indices
        return expression;
    }

    /// <summary>
    /// Applies direct index adjustment for exact collection name matches
    /// </summary>
    private static int ApplyDirectIndexAdjustment(int originalIndex, Models.SlideContext context)
    {
        // Create detailed debug log to aid in troubleshooting
        Logger.Debug($"Adjusting index {originalIndex} for context: {context.CollectionName}, offset: {context.Offset}, isGroup: {context.IsGroupContext}, itemsInGroup: {context.ItemsInGroup}");

        if (context.IsGroupContext)
        {
            // If original index is within the range of items that could be shown on a single slide
            if (originalIndex < context.ItemsInGroup)
            {
                // Adjust it by the offset of the current group
                int adjustedIndex = originalIndex + context.Offset;
                Logger.Debug($"Group context adjustment: {originalIndex} -> {adjustedIndex} (adding offset {context.Offset})");
                return adjustedIndex;
            }
            else
            {
                // Handle index outside the group range
                int adjustedIndex = originalIndex + context.Offset;
                Logger.Debug($"Group context adjustment (out of group range): {originalIndex} -> {adjustedIndex} (adding offset {context.Offset})");
                return adjustedIndex;
            }
        }

        // Simple offset adjustment for non-group contexts
        int simpleAdjustedIndex = originalIndex + context.Offset;
        Logger.Debug($"Direct context adjustment: {originalIndex} + {context.Offset} = {simpleAdjustedIndex}");
        return simpleAdjustedIndex;
    }
}