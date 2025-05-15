namespace DocuChef.PowerPoint.Helpers;

internal static class PowerPointHelper
{
    /// <summary>
    /// Creates appropriate ImagePart based on content type
    /// </summary>
    public static ImagePart? CreateImagePart(SlidePart slidePart, string contentType, string relationshipId)
    {
        switch (contentType)
        {
            case "image/jpeg":
                Logger.Debug("Adding JPEG image part");
                return slidePart.AddImagePart(ImagePartType.Jpeg, relationshipId);
            case "image/png":
                Logger.Debug("Adding PNG image part");
                return slidePart.AddImagePart(ImagePartType.Png, relationshipId);
            case "image/gif":
                Logger.Debug("Adding GIF image part");
                return slidePart.AddImagePart(ImagePartType.Gif, relationshipId);
            case "image/bmp":
                Logger.Debug("Adding BMP image part");
                return slidePart.AddImagePart(ImagePartType.Bmp, relationshipId);
            case "image/tiff":
                Logger.Debug("Adding TIFF image part");
                return slidePart.AddImagePart(ImagePartType.Tiff, relationshipId);
            default:
                Logger.Warning($"Unsupported content type: {contentType}");
                return null;
        }
    }

    /// <summary>
    /// Clones shape outline properly
    /// </summary>
    public static A.Outline? CloneOutline(A.Outline originalOutline)
    {
        if (originalOutline == null)
            return null;

        Logger.Debug("Cloning outline properties");

        var clonedOutline = originalOutline.CloneNode(true) as A.Outline;

        // Ensure cloned outline is not attached to any parent
        if (clonedOutline?.Parent != null)
        {
            clonedOutline.Remove();
        }

        return clonedOutline;
    }

    /// <summary>
    /// Creates default outline for shape
    /// </summary>
    public static A.Outline CreateDefaultOutline(int width = 9525, string colorHex = "000000")
    {
        Logger.Debug($"Creating default outline with width: {width}, color: #{colorHex}");

        var outline = new A.Outline() { Width = width };
        var solidFill = new A.SolidFill();
        var rgbColor = new A.RgbColorModelHex() { Val = colorHex };
        solidFill.AppendChild(rgbColor);
        outline.AppendChild(solidFill);

        return outline;
    }

    /// <summary>
    /// Creates Picture element with image
    /// </summary>
    public static Picture CreatePicture(
        string relationshipId,
        uint shapeId,
        string shapeName,
        long x,
        long y,
        long width,
        long height,
        bool preserveAspectRatio = true,
        A.Outline? outline = null)
    {
        Logger.Debug($"Creating picture: RelID={relationshipId}, ID={shapeId}, Name={shapeName}, " +
                     $"Position=({x}, {y}), Size=({width}, {height})");

        var picture = new Picture();

        // NonVisualPictureProperties
        var nvPicProps = new NonVisualPictureProperties(
            new NonVisualDrawingProperties()
            {
                Id = shapeId,
                Name = shapeName
            },
            new NonVisualPictureDrawingProperties(
                new A.PictureLocks() { NoChangeAspect = preserveAspectRatio }
            ),
            new ApplicationNonVisualDrawingProperties()
        );
        picture.AppendChild(nvPicProps);

        // BlipFill
        var blipFill = new BlipFill();
        var blip = new A.Blip() { Embed = relationshipId };
        blipFill.AppendChild(blip);
        blipFill.AppendChild(new A.SourceRectangle());
        var stretch = new A.Stretch();
        stretch.AppendChild(new A.FillRectangle());
        blipFill.AppendChild(stretch);
        picture.AppendChild(blipFill);

        // ShapeProperties
        var shapeProps = new ShapeProperties();
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
}