using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocuChef.Logging;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Handles element hiding operations based on data bounds checking
/// Extracted from SlideGenerator to improve code organization
/// </summary>
internal class ElementHider
{
    private static readonly Regex ArrayIndexPattern = new(@"(\w+)\[(\d+)\]", RegexOptions.Compiled);

    /// <summary>
    /// Checks if an element should be hidden due to data index overflow
    /// </summary>
    public bool ShouldHideElement(string text, int indexOffset, object? data)
    {
        if (string.IsNullOrEmpty(text) || data == null)
            return false;

        var matches = ArrayIndexPattern.Matches(text);

        foreach (Match match in matches)
        {
            var arrayName = match.Groups[1].Value;
            var currentIndex = int.Parse(match.Groups[2].Value);
            var finalIndex = currentIndex + indexOffset;

            Logger.Debug($"ElementHider: Checking array bounds for {arrayName}[{finalIndex}] (original: {arrayName}[{currentIndex}] + offset {indexOffset})");

            if (!IsIndexValid(arrayName, finalIndex, data))
            {
                Logger.Debug($"ElementHider: Array index {arrayName}[{finalIndex}] is out of bounds, should hide element");
                return true;
            }
        }

        return false;
    }    /// <summary>
         /// Hides an element by completely removing it or making it invisible
         /// </summary>
    public void HideElement(OpenXmlElement element)
    {
        try
        {
            Logger.Debug($"ElementHider: Attempting to hide element of type: {element.GetType().Name}");
            Logger.Debug($"ElementHider: Element has parent: {element.Parent != null}");

            if (element.Parent != null)
            {
                Logger.Debug($"ElementHider: Removing element from parent of type: {element.Parent.GetType().Name}");
                element.Remove();
                Logger.Debug("ElementHider: Successfully removed element from slide");
            }
            else
            {
                Logger.Debug("ElementHider: Element has no parent, making invisible");
                MakeElementInvisible(element);
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"ElementHider: Error hiding element: {ex.Message}");
        }
    }

    /// <summary>
    /// Finds the parent shape element for a text element
    /// </summary>
    public OpenXmlElement? FindParentShape(OpenXmlElement element)
    {
        var current = element.Parent;

        while (current != null)
        {
            if (IsShapeElement(current))
            {
                return current;
            }
            current = current.Parent;
        }

        return null;
    }

    /// <summary>
    /// Checks if the specified array index is valid for the given data
    /// </summary>
    private bool IsIndexValid(string arrayName, int index, object data)
    {
        try
        {
            if (TryGetArrayFromDictionary(data, arrayName, out var arrayValue) ||
                TryGetArrayFromProperty(data, arrayName, out arrayValue))
            {
                return IsValidIndexForCollection(arrayValue, index);
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"ElementHider: Error checking array bounds for {arrayName}[{index}]: {ex.Message}");
        }

        return true; // Default to not hiding if we can't determine bounds
    }

    /// <summary>
    /// Tries to get array value from dictionary
    /// </summary>
    private static bool TryGetArrayFromDictionary(object data, string arrayName, out object? arrayValue)
    {
        arrayValue = null;

        if (data is Dictionary<string, object> dict && dict.TryGetValue(arrayName, out arrayValue))
        {
            return true;
        }

        return false;
    }

    /// <summary>
    /// Tries to get array value from object property using reflection
    /// </summary>
    private static bool TryGetArrayFromProperty(object data, string arrayName, out object? arrayValue)
    {
        arrayValue = null;

        var property = data.GetType().GetProperty(arrayName);
        if (property != null)
        {
            arrayValue = property.GetValue(data);
            return true;
        }

        return false;
    }

    /// <summary>
    /// Checks if the index is valid for the given collection
    /// </summary>
    private static bool IsValidIndexForCollection(object? arrayValue, int index)
    {
        if (arrayValue is System.Collections.IList list)
        {
            return index >= 0 && index < list.Count;
        }

        if (arrayValue is System.Collections.IEnumerable enumerable)
        {
            var count = enumerable.Cast<object>().Count();
            return index >= 0 && index < count;
        }

        return false;
    }

    /// <summary>
    /// Determines if an element is a shape-like element
    /// </summary>
    private static bool IsShapeElement(OpenXmlElement element)
    {
        return element is DocumentFormat.OpenXml.Presentation.Shape ||
               element is DocumentFormat.OpenXml.Presentation.Picture ||
               element is DocumentFormat.OpenXml.Presentation.GraphicFrame;
    }

    /// <summary>
    /// Makes an element invisible by setting various properties
    /// </summary>
    private void MakeElementInvisible(OpenXmlElement element)
    {
        try
        {
            switch (element)
            {
                case DocumentFormat.OpenXml.Presentation.Shape shape:
                    HideShape(shape);
                    break;
                case DocumentFormat.OpenXml.Presentation.Picture picture:
                    HidePicture(picture);
                    break;
                case DocumentFormat.OpenXml.Presentation.ConnectionShape connectionShape:
                    HideConnectionShape(connectionShape);
                    break;
                default:
                    RemoveUnknownElement(element);
                    break;
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"ElementHider: Error making element invisible: {ex.Message}");
        }
    }

    /// <summary>
    /// Hides a shape element
    /// </summary>
    private void HideShape(DocumentFormat.OpenXml.Presentation.Shape shape)
    {
        var nvSpPr = shape.NonVisualShapeProperties;
        if (nvSpPr?.NonVisualDrawingProperties != null)
        {
            nvSpPr.NonVisualDrawingProperties.Hidden = true;
            Logger.Debug("ElementHider: Set shape Hidden property to true");
        }
        else
        {
            SetShapeDimensionsToZero(shape);
        }
    }

    /// <summary>
    /// Hides a picture element
    /// </summary>
    private void HidePicture(DocumentFormat.OpenXml.Presentation.Picture picture)
    {
        var nvPicPr = picture.NonVisualPictureProperties;
        if (nvPicPr?.NonVisualDrawingProperties != null)
        {
            nvPicPr.NonVisualDrawingProperties.Hidden = true;
            Logger.Debug("ElementHider: Set picture Hidden property to true");
        }
        else
        {
            SetPictureDimensionsToZero(picture);
        }
    }

    /// <summary>
    /// Hides a connection shape element
    /// </summary>
    private void HideConnectionShape(DocumentFormat.OpenXml.Presentation.ConnectionShape connectionShape)
    {
        var nvCxnSpPr = connectionShape.NonVisualConnectionShapeProperties;
        if (nvCxnSpPr?.NonVisualDrawingProperties != null)
        {
            nvCxnSpPr.NonVisualDrawingProperties.Hidden = true;
            Logger.Debug("ElementHider: Set connection shape Hidden property to true");
        }
    }

    /// <summary>
    /// Removes unknown element types
    /// </summary>
    private void RemoveUnknownElement(OpenXmlElement element)
    {
        Logger.Debug($"ElementHider: Unknown element type {element.GetType().Name}, trying to remove");

        if (element.Parent != null)
        {
            element.Remove();
            Logger.Debug($"ElementHider: Removed unknown element type {element.GetType().Name}");
        }
    }

    /// <summary>
    /// Sets shape dimensions to zero as a fallback hiding method
    /// </summary>
    private void SetShapeDimensionsToZero(DocumentFormat.OpenXml.Presentation.Shape shape)
    {
        try
        {
            var spPr = shape.ShapeProperties;
            if (spPr?.Transform2D?.Extents != null)
            {
                spPr.Transform2D.Extents.Cx = 0;
                spPr.Transform2D.Extents.Cy = 0;
                Logger.Debug("ElementHider: Set shape dimensions to 0 as fallback");
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"ElementHider: Error setting shape dimensions to zero: {ex.Message}");
        }
    }

    /// <summary>
    /// Sets picture dimensions to zero as a fallback hiding method
    /// </summary>
    private void SetPictureDimensionsToZero(DocumentFormat.OpenXml.Presentation.Picture picture)
    {
        try
        {
            var spPr = picture.ShapeProperties;
            if (spPr?.Transform2D?.Extents != null)
            {
                spPr.Transform2D.Extents.Cx = 0;
                spPr.Transform2D.Extents.Cy = 0;
                Logger.Debug("ElementHider: Set picture dimensions to 0 as fallback");
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"ElementHider: Error setting picture dimensions to zero: {ex.Message}");
        }
    }
}
