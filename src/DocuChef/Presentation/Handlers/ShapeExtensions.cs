using DocumentFormat.OpenXml.Presentation;

namespace DocuChef.Presentation.Handlers;

/// <summary>
/// Extension methods for PowerPoint shapes
/// </summary>
internal static class ShapeExtensions
{
    /// <summary>
    /// Hides a shape by making it invisible
    /// </summary>
    /// <param name="shape">The shape to hide</param>
    public static void Hide(this Shape shape)
    {
        if (shape?.ShapeProperties != null)
        {
            // Set the shape to be hidden by modifying its visibility
            var nvProps = shape.NonVisualShapeProperties?.NonVisualDrawingProperties;
            if (nvProps != null)
            {
                nvProps.Hidden = true;
            }
        }
    }
}
