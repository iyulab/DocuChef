namespace DocuChef.PowerPoint.Helpers;

/// <summary>
/// Helper class for PowerPoint-specific shape operations
/// </summary>
internal static class PowerPointShapeHelper
{
    private static readonly Dictionary<string, (long cx, long cy)> _originalDimensions = new();
    private static readonly Dictionary<string, (long x, long y)> _originalPositions = new();

    /// <summary>
    /// Find array references in shape with simplified implementation
    /// </summary>
    public static List<ArrayReference> FindArrayReferences(P.Shape shape)
    {
        var result = new List<ArrayReference>();

        if (shape?.TextBody == null)
            return result;

        var textRuns = shape.Descendants<A.Text>().ToList();
        foreach (var textRun in textRuns)
        {
            if (!string.IsNullOrEmpty(textRun.Text))
            {
                var references = ArrayReferenceHelper.ExtractArrayReferences(textRun.Text);
                result.AddRange(references);
            }
        }

        return result;
    }

    /// <summary>
    /// Hide a PowerPoint shape completely
    /// </summary>
    public static void HideShape(P.Shape shape)
    {
        if (shape == null)
            return;

        try
        {
            Logger.Debug($"Attempting to hide shape: ID={shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value}, Name={shape.GetShapeName()}");

            // 1. Set Hidden attribute
            SetHiddenAttribute(shape, true);
            Logger.Debug("Hidden attribute set");

            // 2. Clear text first
            if (shape.TextBody != null)
            {
                Logger.Debug("Clearing shape text");
                ClearShapeTextSafely(shape);
                Logger.Debug("Shape text cleared");
            }

            // 3. Set visibility to minimum
            SetMinimumVisibility(shape);
            Logger.Debug($"Shape hidden successfully");
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error hiding shape: {ex.Message}");
            Logger.Debug($"Stack trace: {ex.StackTrace}");
        }
    }

    /// <summary>
    /// Show a PowerPoint shape
    /// </summary>
    public static void ShowShape(P.Shape shape)
    {
        if (shape == null)
            return;

        try
        {
            Logger.Debug($"Attempting to show shape: ID={shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value}");

            SetHiddenAttribute(shape, false);
            RestoreShapeVisibility(shape);

            Logger.Debug($"Shape shown successfully");
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error showing shape: {ex.Message}");
        }
    }

    /// <summary>
    /// Check if shape is hidden
    /// </summary>
    public static bool IsShapeHidden(P.Shape shape)
    {
        return shape?.NonVisualShapeProperties?.NonVisualDrawingProperties?.Hidden?.Value ?? false;
    }

    private static void SetHiddenAttribute(P.Shape shape, bool hidden)
    {
        var nvProps = shape.NonVisualShapeProperties;
        if (nvProps?.NonVisualDrawingProperties != null)
        {
            Logger.Debug($"Setting hidden attribute to {hidden}");
            nvProps.NonVisualDrawingProperties.Hidden = new BooleanValue(hidden);
        }
    }

    private static void ClearShapeTextSafely(P.Shape shape)
    {
        if (shape?.TextBody == null)
            return;

        try
        {
            // Get all paragraphs as a list first
            var existingParagraphs = shape.TextBody.Elements<A.Paragraph>().ToList();
            Logger.Debug($"Found {existingParagraphs.Count} paragraphs to clear");

            // Remove all except the first one
            for (int i = existingParagraphs.Count - 1; i > 0; i--)
            {
                shape.TextBody.RemoveChild(existingParagraphs[i]);
            }

            // Clear text in the first paragraph
            if (existingParagraphs.Count > 0)
            {
                var firstParagraph = existingParagraphs[0];

                // Remove all runs
                var runs = firstParagraph.Elements<A.Run>().ToList();
                foreach (var run in runs)
                {
                    firstParagraph.RemoveChild(run);
                }

                // Add empty run with empty text
                var emptyRun = new A.Run();
                var emptyText = new A.Text("");
                emptyRun.AppendChild(emptyText);
                firstParagraph.AppendChild(emptyRun);
            }
            else
            {
                // No paragraphs exist, create one
                var emptyParagraph = new A.Paragraph();
                var emptyRun = new A.Run();
                var emptyText = new A.Text("");
                emptyRun.AppendChild(emptyText);
                emptyParagraph.AppendChild(emptyRun);
                shape.TextBody.AppendChild(emptyParagraph);
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error in ClearShapeTextSafely: {ex.Message}");
            Logger.Debug($"Stack trace: {ex.StackTrace}");
        }
    }

    private static void SetMinimumVisibility(P.Shape shape)
    {
        try
        {
            var transform = shape.ShapeProperties?.Transform2D;
            if (transform != null)
            {
                // Store original dimensions
                var shapeId = GetShapeId(shape);
                if (!string.IsNullOrEmpty(shapeId))
                {
                    // Store original dimensions
                    if (transform.Extents != null)
                    {
                        StoreOriginalDimensions(shapeId, transform.Extents);
                        transform.Extents.Cx = 0;
                        transform.Extents.Cy = 0;
                    }

                    // Store original position and move off-screen
                    if (transform.Offset != null)
                    {
                        StoreOriginalPosition(shapeId, transform.Offset);
                        transform.Offset.X = -50000000;
                        transform.Offset.Y = -50000000;
                    }
                }
            }

            // Make shape transparent
            var shapeProperties = shape.ShapeProperties;
            if (shapeProperties != null)
            {
                // Remove any existing fill
                var existingFill = shapeProperties.GetFirstChild<A.SolidFill>();
                if (existingFill != null)
                {
                    shapeProperties.RemoveChild(existingFill);
                }

                // Add transparent fill
                var noFill = new A.NoFill();
                shapeProperties.AppendChild(noFill);

                // Remove outline
                var existingOutline = shapeProperties.GetFirstChild<A.Outline>();
                if (existingOutline != null)
                {
                    shapeProperties.RemoveChild(existingOutline);
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Warning($"Error in SetMinimumVisibility: {ex.Message}");
        }
    }

    private static void RestoreShapeVisibility(P.Shape shape)
    {
        var transform = shape.ShapeProperties?.Transform2D;
        if (transform != null)
        {
            var shapeId = GetShapeId(shape);
            if (!string.IsNullOrEmpty(shapeId))
            {
                // Restore dimensions
                RestoreOriginalDimensions(shapeId, transform.Extents);

                // Restore position
                RestoreOriginalPosition(shapeId, transform.Offset);
            }
        }
    }

    private static string GetShapeId(P.Shape shape)
    {
        var id = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value;
        return id.HasValue ? $"shape_{id.Value}" : null;
    }

    private static void StoreOriginalDimensions(string shapeId, A.Extents extents)
    {
        if (extents == null) return;

        long cx = extents.Cx?.Value ?? 0;
        long cy = extents.Cy?.Value ?? 0;

        if (cx > 0 && cy > 0)
        {
            _originalDimensions[shapeId] = (cx, cy);
        }
    }

    private static void RestoreOriginalDimensions(string shapeId, A.Extents extents)
    {
        if (extents == null) return;

        if (_originalDimensions.TryGetValue(shapeId, out var dimensions))
        {
            extents.Cx = dimensions.cx;
            extents.Cy = dimensions.cy;
        }
    }

    private static void StoreOriginalPosition(string shapeId, A.Offset offset)
    {
        if (offset == null) return;

        long x = offset.X?.Value ?? 0;
        long y = offset.Y?.Value ?? 0;

        _originalPositions[shapeId] = (x, y);
    }

    private static void RestoreOriginalPosition(string shapeId, A.Offset offset)
    {
        if (offset == null) return;

        if (_originalPositions.TryGetValue(shapeId, out var position))
        {
            offset.X = position.x;
            offset.Y = position.y;
        }
    }
}