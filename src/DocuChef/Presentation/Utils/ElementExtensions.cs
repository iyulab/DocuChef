using DocumentFormat.OpenXml.Presentation;

namespace DocuChef.Presentation.Utils;

public static class ElementExtensions
{
    /// <summary>
    /// Finds the parent Shape element by traversing up the OpenXml element hierarchy
    /// </summary>
    /// <param name="element">The OpenXml element to start from</param>
    /// <returns>The parent Shape element if found, null otherwise</returns>
    public static P.Shape? FindShape(this OpenXmlElement element)
    {
        if (element == null)
            throw new ArgumentNullException(nameof(element));

        // Check if the element is already a Shape
        if (element is P.Shape shape)
        {
            return shape;
        }

        // For other element types, traverse up the parent hierarchy
        var current = element;

        // Traverse up the hierarchy until we find a Shape or reach the root
        while (current != null)
        {
            // Check if the current parent is a Shape
            if (current is P.Shape parentShape)
            {
                return parentShape;
            }

            // Move to the next parent element
            current = current.Parent;
        }

        // No Shape was found in the hierarchy
        return null;
    }

    /// <summary>
    /// Hides an OpenXml element by setting the Hidden property
    /// </summary>
    /// <param name="element">The element to hide</param>
    public static void Hide(this OpenXmlElement element)
    {
        if (element == null)
            throw new ArgumentNullException(nameof(element));

        // Process appropriate hiding based on PowerPoint element type
        if (element is P.Shape shape)
        {
            // For Shape elements, set Hidden property in NonVisualShapeProperties
            var nvSpPr = shape.NonVisualShapeProperties;
            if (nvSpPr != null)
            {
                var cNvPr = nvSpPr.NonVisualDrawingProperties;
                if (cNvPr != null)
                {
                    cNvPr.Hidden = true;
                }
                else
                {
                    nvSpPr.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 0, Name = "", Hidden = true };
                }
            }
        }
        else if (element is P.GroupShape groupShape)
        {
            // Process for GroupShape elements
            var nvGrpSpPr = groupShape.NonVisualGroupShapeProperties;
            if (nvGrpSpPr != null)
            {
                var cNvPr = nvGrpSpPr.NonVisualDrawingProperties;
                if (cNvPr != null)
                {
                    cNvPr.Hidden = true;
                }
                else
                {
                    nvGrpSpPr.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 0, Name = "", Hidden = true };
                }
            }
        }
        else if (element is P.Picture picture)
        {
            // Process for Picture elements
            var nvPicPr = picture.NonVisualPictureProperties;
            if (nvPicPr != null)
            {
                var cNvPr = nvPicPr.NonVisualDrawingProperties;
                if (cNvPr != null)
                {
                    cNvPr.Hidden = true;
                }
                else
                {
                    nvPicPr.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 0, Name = "", Hidden = true };
                }
            }
        }
        else if (element is P.GraphicFrame graphicFrame)
        {
            // Process for GraphicFrame elements (tables, charts, etc.)
            var nvGraphicFramePr = graphicFrame.NonVisualGraphicFrameProperties;
            if (nvGraphicFramePr != null)
            {
                var cNvPr = nvGraphicFramePr.NonVisualDrawingProperties;
                if (cNvPr != null)
                {
                    cNvPr.Hidden = true;
                }
                else
                {
                    nvGraphicFramePr.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 0, Name = "", Hidden = true };
                }
            }
        }
        else if (element is P.ConnectionShape connShape)
        {
            // Process for connection lines
            var nvCxnSpPr = connShape.NonVisualConnectionShapeProperties;
            if (nvCxnSpPr != null)
            {
                var cNvPr = nvCxnSpPr.NonVisualDrawingProperties;
                if (cNvPr != null)
                {
                    cNvPr.Hidden = true;
                }
                else
                {
                    nvCxnSpPr.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 0, Name = "", Hidden = true };
                }
            }
        }
        else
        {
            // For other elements, try to find and hide the parent Shape
            var parentShape = element.FindShape();
            if (parentShape != null)
            {
                Hide(parentShape);
                return;
            }

            // If no parent Shape was found, try to set a hidden attribute directly
            element.SetAttribute(new OpenXmlAttribute("", "hidden", "", "true"));

            // Also try to set it on the parent if available
            element.Parent?.SetAttribute(new OpenXmlAttribute("", "hidden", "", "true"));
        }
    }
}