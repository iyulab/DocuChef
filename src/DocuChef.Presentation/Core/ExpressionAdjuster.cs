using System.Text;

namespace DocuChef.Presentation.Core
{
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
            // If context collection contains underscores, check for parent/child matches
            if (context.CollectionName.Contains('_'))
            {
                string[] segments = context.CollectionName.Split('_');

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
            // For grouped contexts (multiple items per slide), special handling is needed
            if (context.IsGroupContext)
            {
                // If original index is within group bounds and using an offset that is for the current slide,
                // we can use a simple relative index
                if (originalIndex < context.ItemsInGroup)
                {
                    Logger.Debug($"Group context adjustment: keeping {originalIndex} as is since it's within group bounds");
                    return originalIndex + context.Offset;
                }
            }

            // Simple offset adjustment for exact collection name matches
            int adjustedIndex = originalIndex + context.Offset;
            Logger.Debug($"Direct context adjustment: {originalIndex} + {context.Offset} = {adjustedIndex}");
            return adjustedIndex;
        }
    }
}