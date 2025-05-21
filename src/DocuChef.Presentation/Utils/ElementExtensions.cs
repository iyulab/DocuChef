using DocumentFormat.OpenXml.Presentation;

namespace DocuChef.Presentation.Utils;

public static class ElementExtensions
{
    public static P.Shape? FindShape(this OpenXmlElement element)
    {
        if (element == null)
            throw new ArgumentNullException(nameof(element));
        // Process according to the PowerPoint element type
        if (element is DocumentFormat.OpenXml.Drawing.Text text)
        {
            // Process for Text elements - find and return Shape element by traversing up
            var current = text.Parent;
            while (current != null)
            {
                if (current is P.Shape parentShape)
                {
                    return parentShape; // Return if Shape is found
                }
                // Move to the next parent element
                current = current.Parent;
            }
        }
        else if (element is P.Shape shape)
        {
            return shape; // Return directly if it's already a Shape element
        }

        return null;
    }

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
        else if (element is DocumentFormat.OpenXml.Drawing.Text text)
        {
            var parentShape = text.FindShape();
            if (parentShape != null)
            {
                Hide(parentShape);
                return;
            }

            // Process for Text elements - find and hide the Shape element by traversing up
            // First add hidden attribute to the text itself (if needed)
            text.SetAttribute(new OpenXmlAttribute("", "hidden", "", "true"));

            // Try general processing if Shape not found
            text.Parent?.SetAttribute(new OpenXmlAttribute("", "hidden", "", "true"));
        }
        else
        {
            // General processing for other PowerPoint elements
            // For general OpenXML elements, add a hidden attribute
            element.SetAttribute(new OpenXmlAttribute("", "hidden", "", "true"));
        }
    }
}