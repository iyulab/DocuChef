using DocumentFormat.OpenXml.Packaging;

namespace DocuChef.Presentation.Handlers;

/// <summary>
/// Handles image processing operations for PowerPoint presentations
/// </summary>
internal static class ImageHandler
{

    /// <summary>
    /// Processes an image in a shape using default settings
    /// </summary>
    public static void Process(P.Shape shape, string path)
    {
        if (shape == null)
        {
            Logger.Error("Shape is null");
            return;
        }

        if (string.IsNullOrEmpty(path))
        {
            Logger.Warning("Image file path is null or empty");
            shape.Hide();
            return;
        }

        try
        {
            // Use default values for width, height and aspect ratio
            Process(shape, path, 300, 200, true);
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing image data: {ex.Message}", ex);
            shape.Hide();
        }
    }

    /// <summary>
    /// Processes an image in a shape with custom dimensions
    /// </summary>
    private static void Process(P.Shape shape, string path, int width, int height, bool preserveAspectRatio)
    {
        if (shape == null)
        {
            Logger.Error("Shape is null");
            return;
        }

        if (string.IsNullOrEmpty(path))
        {
            Logger.Warning("Image path is null or empty");
            shape.Hide();
            return;
        }

        try
        {
            // Get the slide part that contains the shape
            var slidePart = GetSlidePart(shape);
            if (slidePart == null)
            {
                Logger.Error("Unable to find slide part for shape");
                shape.Hide();
                return;
            }

            // Resolve image path if it's a URL or file path
            string resolvedPath = ResolveImagePath(path);
            if (string.IsNullOrEmpty(resolvedPath))
            {
                Logger.Warning($"Unable to resolve image path: {path}");
                shape.Hide();
                return;
            }

            // Get shape properties
            var shapeProperties = ExtractShapeProperties(shape);

            // Process the image in the shape
            bool success = AddImageToShape(
                slidePart,
                shape,
                resolvedPath,
                shapeProperties.Id,
                shapeProperties.Name,
                shapeProperties.X,
                shapeProperties.Y,
                shapeProperties.Width,
                shapeProperties.Height,
                preserveAspectRatio,
                shapeProperties.Outline,
                shapeProperties.Description);

            if (!success)
            {
                Logger.Warning($"Failed to process image in shape: {path}");
                shape.Hide();
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing image: {ex.Message}", ex);
            shape.Hide();
        }
    }

    /// <summary>
    /// Gets the slide part that contains the given shape
    /// </summary>
    private static SlidePart? GetSlidePart(P.Shape shape)
    {
        var current = shape.Parent;
        while (current != null)
        {
            if (current is Slide slide)
            {
                return slide.SlidePart;
            }
            current = current.Parent;
        }
        return null;
    }    /// <summary>
         /// Resolves an image path to a local file path
         /// </summary>
    private static string ResolveImagePath(string path)
    {
        try
        {
            // Use ClosedXML's ImageHelper to handle both local paths and URLs
            return ClosedXML.Report.XLCustom.Functions.ImageHelper.GetImageFromPathOrUrl(path) ?? string.Empty;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error resolving image path '{path}': {ex.Message}");
            return string.Empty;
        }
    }    /// <summary>
         /// Extracts shape properties for image insertion
         /// </summary>
    private static ShapeProperties ExtractShapeProperties(P.Shape shape)
    {
        var properties = new ShapeProperties();

        var nvsp = shape.NonVisualShapeProperties;
        var nvdp = nvsp?.NonVisualDrawingProperties;
        var transform = shape.ShapeProperties?.Transform2D;

        // Extract shape ID and name
        properties.Id = nvdp?.Id?.Value ?? 1000u;
        properties.Name = nvdp?.Name?.Value ?? "Image_Shape";
        properties.Description = nvdp?.Description?.Value;

        // Extract position and dimensions
        properties.X = transform?.Offset?.X?.Value ?? 1524000;
        properties.Y = transform?.Offset?.Y?.Value ?? 1524000;
        properties.Width = transform?.Extents?.Cx?.Value ?? 3048000;
        properties.Height = transform?.Extents?.Cy?.Value ?? 2286000;

        // Extract outline if present
        properties.Outline = shape.ShapeProperties?.GetFirstChild<A.Outline>();

        return properties;
    }    /// <summary>
         /// Adds an image to a shape, replacing the shape with a picture element
         /// </summary>
    private static bool AddImageToShape(
        SlidePart slidePart,
        P.Shape shape,
        string imagePath,
        uint shapeId,
        string? shapeName,
        long x,
        long y,
        long width,
        long height,
        bool preserveAspectRatio,
        A.Outline? outline,
        string? description)
    {
        try
        {
            Logger.Debug($"Adding image to shape: ID={shapeId}, Name={shapeName}, Path={imagePath}");

            // Determine image content type
            string contentType = GetContentType(imagePath);
            if (string.IsNullOrEmpty(contentType))
            {
                Logger.Warning($"Unsupported image format: {Path.GetExtension(imagePath)}");
                return false;
            }

            // Generate a unique relationship ID
            string relationshipId = GenerateUniqueRelationshipId(slidePart);

            // Create image part based on content type
            ImagePart? imagePart = CreateImagePart(slidePart, contentType, relationshipId);
            if (imagePart == null)
                return false;

            // Add image data to part
            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }

            // Use appropriate outline or create default
            A.Outline? finalOutline = null;
            if (outline != null)
            {
                // Clone outline to avoid XML hierarchy issues
                finalOutline = outline.CloneNode(true) as A.Outline;
                if (finalOutline?.Parent != null)
                {
                    finalOutline.Remove();
                }
            }
            else
            {
                finalOutline = CreateDefaultOutline();
            }

            // Create picture element
            var picture = CreatePicture(
                relationshipId,
                shapeId,
                shapeName,
                x,
                y,
                width,
                height,
                preserveAspectRatio,
                finalOutline,
                description);

            // Replace original shape with picture
            var parent = shape.Parent;
            if (parent == null)
            {
                Logger.Error("Cannot find parent element for the shape");
                return false;
            }

            parent.InsertAfter(picture, shape);
            parent.RemoveChild(shape);

            Logger.Debug("Successfully replaced shape with image");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error adding image to shape: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Gets the content type for an image file
    /// </summary>
    private static string GetContentType(string imagePath)
    {
        string extension = Path.GetExtension(imagePath).ToLowerInvariant();
        switch (extension)
        {
            case ".jpg":
            case ".jpeg":
                return "image/jpeg";
            case ".png":
                return "image/png";
            case ".gif":
                return "image/gif";
            case ".bmp":
                return "image/bmp";
            case ".tiff":
            case ".tif":
                return "image/tiff";
            default:
                return string.Empty;
        }
    }

    /// <summary>
    /// Creates an image part based on content type
    /// </summary>
    private static ImagePart? CreateImagePart(SlidePart slidePart, string contentType, string relationshipId)
    {
        switch (contentType)
        {
            case "image/jpeg":
                return slidePart.AddImagePart(ImagePartType.Jpeg, relationshipId);
            case "image/png":
                return slidePart.AddImagePart(ImagePartType.Png, relationshipId);
            case "image/gif":
                return slidePart.AddImagePart(ImagePartType.Gif, relationshipId);
            case "image/bmp":
                return slidePart.AddImagePart(ImagePartType.Bmp, relationshipId);
            case "image/tiff":
                return slidePart.AddImagePart(ImagePartType.Tiff, relationshipId);
            default:
                Logger.Warning($"Unsupported content type: {contentType}");
                return null;
        }
    }

    /// <summary>
    /// Creates a default outline for a shape
    /// </summary>
    private static A.Outline CreateDefaultOutline(int width = 9525, string colorHex = "000000")
    {
        var outline = new A.Outline() { Width = width };
        var solidFill = new A.SolidFill();
        var rgbColor = new A.RgbColorModelHex() { Val = colorHex };
        solidFill.AppendChild(rgbColor);
        outline.AppendChild(solidFill);
        return outline;
    }    /// <summary>
         /// Creates a picture element with an image
         /// </summary>
    private static P.Picture CreatePicture(
        string relationshipId,
        uint shapeId,
        string? shapeName,
        long x,
        long y,
        long width,
        long height,
        bool preserveAspectRatio = true,
        A.Outline? outline = null,
        string? description = null)
    {
        var picture = new P.Picture();

        // NonVisualPictureProperties
        var nvPicProps = new P.NonVisualPictureProperties(
            new P.NonVisualDrawingProperties()
            {
                Id = shapeId,
                Name = shapeName,
                Description = description
            },
            new P.NonVisualPictureDrawingProperties(
                new A.PictureLocks() { NoChangeAspect = preserveAspectRatio }
            ),
            new P.ApplicationNonVisualDrawingProperties()
        );
        picture.AppendChild(nvPicProps);

        // BlipFill
        var blipFill = new P.BlipFill();
        var blip = new A.Blip() { Embed = relationshipId };
        blipFill.AppendChild(blip);
        blipFill.AppendChild(new A.SourceRectangle());
        var stretch = new A.Stretch();
        stretch.AppendChild(new A.FillRectangle());
        blipFill.AppendChild(stretch);
        picture.AppendChild(blipFill);

        // ShapeProperties
        var shapeProps = new P.ShapeProperties();
        var transform2D = new A.Transform2D();
        transform2D.Offset = new A.Offset() { X = x, Y = y };
        transform2D.Extents = new A.Extents() { Cx = width, Cy = height };
        shapeProps.AppendChild(transform2D);

        // Preset geometry
        shapeProps.AppendChild(new A.PresetGeometry(
            new A.AdjustValueList()
        )
        { Preset = A.ShapeTypeValues.Rectangle });

        // Outline
        if (outline != null)
        {
            shapeProps.AppendChild(outline);
        }

        picture.AppendChild(shapeProps);

        return picture;
    }

    /// <summary>
    /// Generates a unique relationship ID for a slide part
    /// </summary>
    private static string GenerateUniqueRelationshipId(SlidePart slidePart)
    {
        string baseId = "rImage";
        int counter = 1;

        var existingIds = slidePart.Parts.Select(p => p.RelationshipId).ToHashSet();

        string relationshipId;
        do
        {
            relationshipId = $"{baseId}{counter++}";
        } while (existingIds.Contains(relationshipId));

        return relationshipId;
    }    /// <summary>
         /// Clears temporary resources used by the image handler
         /// </summary>
    public static void Cleanup()
    {
        // No cleanup needed when using ClosedXML ImageHelper
        // The ImageHelper manages its own temporary files
    }

    /// <summary>    /// <summary>
    /// Helper class for storing shape properties
    /// </summary>
    private class ShapeProperties
    {
        public uint Id { get; set; }
        public string? Name { get; set; }
        public string? Description { get; set; }
        public long X { get; set; }
        public long Y { get; set; }
        public long Width { get; set; }
        public long Height { get; set; }
        public A.Outline? Outline { get; set; }
    }
}
