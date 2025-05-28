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
    private readonly ElementHider _elementHider;

    public ExpressionUpdater()
    {
        _elementHider = new ElementHider();
    }

    /// <summary>
    /// Updates expressions in the slide with the given index offset and hides elements that exceed data bounds
    /// Note: This only adjusts array indices, actual data binding happens in DataBinder.cs
    /// </summary>
    public void UpdateExpressionsWithIndexOffset(SlidePart slidePart, int indexOffset, object? data)
    {
        if (slidePart?.Slide == null || indexOffset <= 0)
            return;

        Logger.Debug($"ExpressionUpdater: Updating expressions with index offset {indexOffset}"); try
        {
            var textElements = slidePart.Slide.Descendants<DrawingText>().ToList();
            var elementsToHide = CollectElementsToHide(textElements, indexOffset, data);

            UpdateTextElements(textElements, indexOffset, elementsToHide);
            HideOverflowElements(elementsToHide);
        }
        catch (Exception ex)
        {
            Logger.Warning($"ExpressionUpdater: Error updating expressions with index offset: {ex.Message}");
        }
    }    /// <summary>
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
    private void UpdateTextElements(List<DrawingText> textElements, int indexOffset, List<OpenXmlElement> elementsToHide)
    {
        foreach (var textElement in textElements)
        {
            if (string.IsNullOrEmpty(textElement.Text))
                continue;

            // Skip elements that will be hidden
            var parentShape = _elementHider.FindParentShape(textElement);
            if (parentShape != null && elementsToHide.Contains(parentShape))
                continue;

            var updatedText = AdjustArrayIndicesInText(textElement.Text, indexOffset);
            if (updatedText != textElement.Text)
            {
                Logger.Debug($"ExpressionUpdater: Updated expression from '{textElement.Text}' to '{updatedText}'");
                textElement.Text = updatedText;
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
            var newIndex = currentIndex + indexOffset;
            return $"{arrayName}[{newIndex}]";
        });
    }
}
