using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocuChef.Logging;
using DrawingText = DocumentFormat.OpenXml.Drawing.Text;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Handles expression updating operations in slides
/// Extracted from SlideGenerator to improve code organization
/// </summary>
internal class ExpressionUpdater
{
    private static readonly Regex ArrayIndexPattern = new(@"(\w+)\[(\d+)\]", RegexOptions.Compiled);
    private static readonly Regex ContextOperatorPattern = new(@"(\w+)>(\w+)", RegexOptions.Compiled);
    private readonly ElementHider _elementHider;

    public ExpressionUpdater()
    {
        _elementHider = new ElementHider();
    }    /// <summary>
         /// Updates expressions in the slide with the given index offset and context, 
         /// and hides elements that exceed data bounds
         /// Note: This adjusts array indices and resolves context operators, 
         /// actual data binding happens in DataBinder.cs
         /// </summary>
    public void UpdateExpressionsWithIndexOffset(SlidePart slidePart, int indexOffset, object? data, string? contextPath = null)
    {
        if (slidePart?.Slide == null)
            return;

        Logger.Debug($"ExpressionUpdater: Updating expressions with index offset {indexOffset}, contextPath: '{contextPath ?? "null"}'");

        try
        {
            var textElements = slidePart.Slide.Descendants<DrawingText>().ToList();
            var elementsToHide = CollectElementsToHide(textElements, indexOffset, data);

            UpdateTextElements(textElements, indexOffset, elementsToHide, contextPath);
            HideOverflowElements(elementsToHide);
        }
        catch (Exception ex)
        {
            Logger.Warning($"ExpressionUpdater: Error updating expressions with index offset: {ex.Message}");
        }
    }/// <summary>
     /// Collects elements that should be hidden due to data overflow
     /// </summary>
    private List<OpenXmlElement> CollectElementsToHide(List<DrawingText> textElements, int indexOffset, object? data)
    {
        var elementsToHide = new List<OpenXmlElement>();

        foreach (var textElement in textElements)
        {
            if (string.IsNullOrEmpty(textElement.Text))
                continue;

            if (_elementHider.ShouldHideElement(textElement.Text, indexOffset, data))
            {
                var parentShape = _elementHider.FindParentShape(textElement);
                if (parentShape != null && !elementsToHide.Contains(parentShape))
                {
                    elementsToHide.Add(parentShape);
                    Logger.Debug($"ExpressionUpdater: Marking shape for hiding due to data overflow in expression: '{textElement.Text}'");
                }
            }
        }

        return elementsToHide;
    }    /// <summary>
         /// Updates text elements with adjusted array indices
         /// </summary>
    private void UpdateTextElements(List<DrawingText> textElements, int indexOffset, List<OpenXmlElement> elementsToHide, string? contextPath)
    {
        foreach (var textElement in textElements)
        {
            if (string.IsNullOrEmpty(textElement.Text))
                continue;

            // Skip elements that will be hidden
            var parentShape = _elementHider.FindParentShape(textElement);
            if (parentShape != null && elementsToHide.Contains(parentShape))
                continue;

            var updatedText = textElement.Text;

            // First, resolve context operators if contextPath is provided
            if (!string.IsNullOrEmpty(contextPath))
            {
                updatedText = ResolveContextOperators(updatedText, contextPath);
            }

            // Then, adjust array indices with offset
            if (indexOffset > 0)
            {
                updatedText = AdjustArrayIndicesInText(updatedText, indexOffset);
            }
            if (updatedText != textElement.Text)
            {
                Logger.Debug($"ExpressionUpdater: Updated expression from '{textElement.Text}' to '{updatedText}'");
                textElement.Text = updatedText;
            }

            if (!string.IsNullOrEmpty(contextPath))
            {
                var contextResolvedText = ResolveContextOperators(textElement.Text, contextPath);
                if (contextResolvedText != textElement.Text)
                {
                    Logger.Debug($"ExpressionUpdater: Resolved context operators in expression from '{textElement.Text}' to '{contextResolvedText}'");
                    textElement.Text = contextResolvedText;
                }
            }
        }
    }

    /// <summary>
    /// Hides elements that contain data overflow
    /// </summary>
    private void HideOverflowElements(List<OpenXmlElement> elementsToHide)
    {
        foreach (var element in elementsToHide)
        {
            _elementHider.HideElement(element);
        }
    }

    /// <summary>
    /// Adjusts array indices in text expressions
    /// Example: "${Items[0].Name}" becomes "${Items[2].Name}" with offset 2
    /// </summary>
    private string AdjustArrayIndicesInText(string text, int indexOffset)
    {
        if (string.IsNullOrEmpty(text) || indexOffset <= 0)
            return text;

        return ArrayIndexPattern.Replace(text, match =>
        {
            var arrayName = match.Groups[1].Value;
            var currentIndex = int.Parse(match.Groups[2].Value);

            // Don't adjust Items array indices - they should always start from 0 within each Type
            if (arrayName.EndsWith("Items") || arrayName.Contains(">Items"))
            {
                Logger.Debug($"ExpressionUpdater: Skipping index adjustment for Items array: '{match.Value}' (Items indices should remain 0-based)");
                return match.Value;
            }

            var newIndex = currentIndex + indexOffset;
            var result = $"{arrayName}[{newIndex}]";
            Logger.Debug($"ExpressionUpdater: Adjusted array index from '{match.Value}' to '{result}' (offset: {indexOffset})");
            return result;
        });
    }    /// <summary>
         /// Applies alias transformations to expressions in text
         /// Example: "${Items[0]}" becomes "${Products>Items[0]}" when alias "Products>Items as Items" is defined
         /// </summary>
    public string ApplyAliases(string text, Dictionary<string, string> aliasMap)
    {
        if (string.IsNullOrEmpty(text) || aliasMap == null || aliasMap.Count == 0)
        {
            Logger.Debug($"ExpressionUpdater.ApplyAliases: Skipping - text='{text}', aliasMap count={aliasMap?.Count ?? 0}");
            return text;
        }

        Logger.Debug($"ExpressionUpdater.ApplyAliases: Processing text '{text}' with {aliasMap.Count} alias mappings");

        var result = text;
        var expressionPattern = new Regex(@"\$\{([^}]+)\}", RegexOptions.Compiled);
        var matches = expressionPattern.Matches(text);

        Logger.Debug($"ExpressionUpdater.ApplyAliases: Found {matches.Count} expression matches");

        result = expressionPattern.Replace(result, match =>
        {
            var expression = match.Groups[1].Value;
            Logger.Debug($"ExpressionUpdater.ApplyAliases: Processing expression '{expression}'");

            var transformedExpression = TransformAliasExpression(expression, aliasMap);

            if (transformedExpression != expression)
            {
                Logger.Debug($"ExpressionUpdater: Applied alias transformation from '${{{expression}}}' to '${{{transformedExpression}}}'");
            }
            else
            {
                Logger.Debug($"ExpressionUpdater: No alias transformation for '${{{expression}}}'");
            }

            return "${" + transformedExpression + "}";
        });

        Logger.Debug($"ExpressionUpdater.ApplyAliases: Final result '{result}'");
        return result;
    }    /// <summary>
         /// Transforms a single expression using alias mapping
         /// </summary>
    private string TransformAliasExpression(string expression, Dictionary<string, string> aliasMap)
    {
        foreach (var alias in aliasMap)
        {
            var aliasName = alias.Key;
            var aliasPath = alias.Value;

            // Check if expression starts with the alias name
            if (expression.StartsWith(aliasName))
            {
                // Replace the alias with the full path
                // Example: "Items[0]" becomes "Products>Items[0]" when alias is "Products>Items as Items"
                var remainingPart = expression.Substring(aliasName.Length);
                return aliasPath + remainingPart;
            }

            // Also check for alias usage within function calls or other contexts
            // Use word boundary to ensure we match complete variable names
            var pattern = $@"\b{Regex.Escape(aliasName)}\b";
            if (Regex.IsMatch(expression, pattern))
            {
                var transformed = Regex.Replace(expression, pattern, aliasPath);
                if (transformed != expression)
                {
                    Logger.Debug($"ExpressionUpdater: Applied alias transformation from '{expression}' to '{transformed}'");
                    return transformed;
                }
            }
        }

        return expression;
    }

    /// <summary>
    /// Resolves context operators (>) in expressions using the current context path
    /// Example: "Categories>Items[0].Name" with contextPath "Categories[1]" becomes "Categories__1__Items[0].Name"
    /// </summary>
    private string ResolveContextOperators(string text, string contextPath)
    {
        if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(contextPath) || !text.Contains(">"))
            return text;

        Logger.Debug($"ExpressionUpdater: Resolving context operators in '{text}' with contextPath '{contextPath}'");

        // Extract current index from contextPath (e.g., "Categories[1]" -> 1)
        var contextIndexMatch = System.Text.RegularExpressions.Regex.Match(contextPath, @"(\w+)\[(\d+)\]");
        if (!contextIndexMatch.Success)
        {
            Logger.Debug($"ExpressionUpdater: No index found in contextPath '{contextPath}'");
            return text;
        }

        var contextCollectionName = contextIndexMatch.Groups[1].Value;
        var currentIndex = contextIndexMatch.Groups[2].Value;

        // Replace context operators with indexed notation
        // Categories>Items becomes Categories__1__Items (when currentIndex = 1)
        var result = ContextOperatorPattern.Replace(text, match =>
        {
            var leftSide = match.Groups[1].Value;
            var rightSide = match.Groups[2].Value;

            // Only replace if the left side matches our context collection
            if (leftSide.Equals(contextCollectionName, StringComparison.OrdinalIgnoreCase))
            {
                var replacement = $"{leftSide}__{currentIndex}__{rightSide}";
                Logger.Debug($"ExpressionUpdater: Replaced '{match.Value}' with '{replacement}'");
                return replacement;
            }

            return match.Value;
        });

        return result;
    }
}
