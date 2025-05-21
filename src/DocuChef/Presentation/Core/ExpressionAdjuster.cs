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

        Logger.Debug($"Adjusting expression indices in text: '{originalText}' with context {context.CollectionName}, offset {context.Offset}");

        // Simple string-based approach without regex
        StringBuilder result = new StringBuilder();
        int currentPos = 0;
        int nextExprStart;

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

            // Check for collection name in the expression
            string searchKey = context.CollectionName + "[";
            int collectionPos = fullExpr.IndexOf(searchKey);

            if (collectionPos != -1)
            {
                // Found collection name with bracket, now extract the index
                int indexStart = collectionPos + searchKey.Length;
                int indexEnd = fullExpr.IndexOf("]", indexStart);

                if (indexEnd != -1)
                {
                    string indexStr = fullExpr.Substring(indexStart, indexEnd - indexStart);

                    if (int.TryParse(indexStr, out int originalIndex))
                    {
                        // Calculate adjusted index
                        int adjustedIndex = ApplyDirectIndexAdjustment(originalIndex, context);

                        Logger.Debug($"Adjusted index in expression: {originalIndex} -> {adjustedIndex}");

                        // Create the adjusted expression
                        string beforeIndex = fullExpr.Substring(0, indexStart);
                        string afterIndex = fullExpr.Substring(indexEnd);
                        string adjustedExpr = beforeIndex + adjustedIndex + afterIndex;

                        // Append the adjusted expression
                        result.Append(adjustedExpr);
                    }
                    else
                    {
                        // Not a valid index, keep the original expression
                        result.Append(fullExpr);
                    }
                }
                else
                {
                    // No closing bracket for index, keep the original expression
                    result.Append(fullExpr);
                }
            }
            else
            {
                // Check if this is a parent or child collection match
                string adjustedExpr = CheckForParentChildMatch(fullExpr, context);
                result.Append(adjustedExpr);
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
    /// Checks for parent or child collection matches
    /// </summary>
    private static string CheckForParentChildMatch(string expression, Models.SlideContext context)
    {
        // Get the hierarchy delimiter from options
        string delimiter = PowerPointOptions.Current.HierarchyDelimiter;

        // If context collection contains the hierarchy delimiter, check for parent/child matches
        if (context.CollectionName.Contains(delimiter))
        {
            // Split using the configured delimiter
            string[] segments = context.CollectionName.Split(
                new[] { delimiter },
                StringSplitOptions.None);

            // Check for parent segment match
            if (segments.Length > 0)
            {
                string parentKey = segments[0] + "[";
                int parentPos = expression.IndexOf(parentKey, StringComparison.OrdinalIgnoreCase);

                if (parentPos != -1)
                {
                    return AdjustIndexInExpression(expression, parentPos, parentKey.Length, context, true);
                }
            }

            // Check for child segment match
            if (segments.Length > 1)
            {
                string childKey = segments[1] + "[";
                int childPos = expression.IndexOf(childKey, StringComparison.OrdinalIgnoreCase);

                if (childPos != -1)
                {
                    return AdjustIndexInExpression(expression, childPos, childKey.Length, context, false);
                }
            }
        }

        // No match found, return original expression
        return expression;
    }

    /// <summary>
    /// Adjusts index in expression
    /// </summary>
    private static string AdjustIndexInExpression(string expression, int keyPos, int keyLength, Models.SlideContext context, bool isParent)
    {
        // Extract the index
        int indexStart = keyPos + keyLength;
        int indexEnd = expression.IndexOf("]", indexStart);

        if (indexEnd != -1)
        {
            string indexStr = expression.Substring(indexStart, indexEnd - indexStart);

            if (int.TryParse(indexStr, out int originalIndex))
            {
                // Calculate adjusted index based on whether it's parent or child
                int adjustedIndex;

                if (isParent && context.ParentContext != null)
                {
                    adjustedIndex = originalIndex + context.ParentContext.Offset;
                    Logger.Debug($"Parent match adjustment: {originalIndex} + parent offset {context.ParentContext.Offset} = {adjustedIndex}");
                }
                else
                {
                    adjustedIndex = originalIndex + context.Offset;
                    Logger.Debug($"Child/Default match adjustment: {originalIndex} + {context.Offset} = {adjustedIndex}");
                }

                // Create the adjusted expression
                string beforeIndex = expression.Substring(0, indexStart);
                string afterIndex = expression.Substring(indexEnd);
                return beforeIndex + adjustedIndex + afterIndex;
            }
        }

        // No valid index found, return original
        return expression;
    }

    /// <summary>
    /// Applies direct index adjustment for exact collection name matches
    /// </summary>
    private static int ApplyDirectIndexAdjustment(int originalIndex, Models.SlideContext context)
    {
        // Create detailed debug log to aid in troubleshooting
        Logger.Debug($"Adjusting index {originalIndex} for context: {context.CollectionName}, offset: {context.Offset}, isGroup: {context.IsGroupContext}, itemsInGroup: {context.ItemsInGroup}");

        // For grouped contexts (multiple items per slide), special handling is needed
        if (context.IsGroupContext)
        {
            // If original index is within the range of items that could be shown on a single slide
            // For example, with max: 5, indices 0-4 could appear on a single slide
            if (originalIndex < context.ItemsInGroup)
            {
                // This is a fixed position within the group (e.g., Items[0], Items[1], etc.)
                // So we adjust it by the offset of the current group
                int adjustedIndex = originalIndex + context.Offset;
                Logger.Debug($"Group context adjustment: {originalIndex} -> {adjustedIndex} (adding offset {context.Offset})");
                return adjustedIndex;
            }
            else
            {
                // The index is outside the range of what this group can show
                // For safety, we still adjust it, but this should generally not occur
                // unless the template has expressions for items that won't be displayed
                int adjustedIndex = originalIndex + context.Offset;
                Logger.Debug($"Group context adjustment (out of group range): {originalIndex} -> {adjustedIndex} (adding offset {context.Offset})");
                return adjustedIndex;
            }
        }

        // Simple offset adjustment for exact collection name matches (non-group contexts)
        int simpleAdjustedIndex = originalIndex + context.Offset;
        Logger.Debug($"Direct context adjustment: {originalIndex} + {context.Offset} = {simpleAdjustedIndex}");
        return simpleAdjustedIndex;
    }
}