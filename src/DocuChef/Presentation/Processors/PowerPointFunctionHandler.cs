using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocuChef.Logging;

namespace DocuChef.Presentation.Processors;

/// <summary>
/// Handles PowerPoint-specific function processing
/// </summary>
public static class PowerPointFunctionHandler
{
    private static readonly Regex ImageFunctionRegex = new Regex(
        @"PPT_IMAGE_PLACEHOLDER:([^|]+)(?:\|(\d+))?(?:\|(\d+))?(?:\|(true|false))?",
        RegexOptions.Compiled | RegexOptions.IgnoreCase
    );

    /// <summary>
    /// Processes PowerPoint functions in a slide part
    /// </summary>
    public static void ProcessFunctions(SlidePart slidePart, object data, string? templatePath = null)
    {
        if (slidePart?.Slide == null)
            return;

        var textElements = slidePart.Slide.Descendants<A.Text>().ToList();
        
        foreach (var textElement in textElements)
        {
            if (string.IsNullOrEmpty(textElement.Text))
                continue;

            ProcessImageFunctions(slidePart, textElement, data, templatePath);
        }
    }

    /// <summary>
    /// Processes image function calls in a text element
    /// </summary>
    private static void ProcessImageFunctions(SlidePart slidePart, A.Text textElement, object data, string? templatePath)
    {
        if (string.IsNullOrEmpty(textElement.Text))
            return;

        var matches = ImageFunctionRegex.Matches(textElement.Text);
        if (matches.Count == 0)
            return;

        foreach (Match match in matches)
        {
            try
            {
                ProcessImageFunction(slidePart, textElement, match, data, templatePath);
            }
            catch (Exception ex)
            {
                Logger.Error($"PowerPointFunctionHandler: Error processing image function '{match.Value}' - {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Processes a single image function call
    /// </summary>
    private static void ProcessImageFunction(SlidePart slidePart, A.Text textElement, Match match, object data, string? templatePath)
    {
        var propertyPath = match.Groups[1].Value;
        var widthStr = match.Groups[2].Success ? match.Groups[2].Value : null;
        var heightStr = match.Groups[3].Success ? match.Groups[3].Value : null;
        var preserveAspectRatioStr = match.Groups[4].Success ? match.Groups[4].Value : "true";

        // Get image data from property path
        var imageData = GetImageData(data, propertyPath);
        if (imageData == null)
        {
            Logger.Debug($"PowerPointFunctionHandler: No image data found for property '{propertyPath}'");
            // Replace with empty string
            textElement.Text = textElement.Text.Replace(match.Value, string.Empty);
            return;
        }

        // Parse dimensions
        int? width = null, height = null;
        if (int.TryParse(widthStr, out var w)) width = w;
        if (int.TryParse(heightStr, out var h)) height = h;
        var preserveAspectRatio = bool.TryParse(preserveAspectRatioStr, out var preserve) ? preserve : true;

        try
        {
            // Insert image and replace function call with empty string
            InsertImage(slidePart, textElement, imageData, width, height, preserveAspectRatio);
            textElement.Text = textElement.Text.Replace(match.Value, string.Empty);
            
            Logger.Debug($"PowerPointFunctionHandler: Successfully inserted image for property '{propertyPath}'");
        }
        catch (Exception ex)
        {
            Logger.Error($"PowerPointFunctionHandler: Failed to insert image for property '{propertyPath}' - {ex.Message}");
            // Replace with placeholder text
            textElement.Text = textElement.Text.Replace(match.Value, $"[Image: {propertyPath}]");
        }
    }

    /// <summary>
    /// Retrieves image data from the data object using property path
    /// </summary>
    private static byte[]? GetImageData(object data, string propertyPath)
    {
        try
        {
            var value = GetPropertyValue(data, propertyPath);
            return value switch
            {
                byte[] bytes => bytes,
                string path when File.Exists(path) => File.ReadAllBytes(path),
                _ => null
            };
        }
        catch (Exception ex)
        {
            Logger.Error($"PowerPointFunctionHandler: Error getting image data for '{propertyPath}' - {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Gets property value using reflection with support for nested properties
    /// </summary>
    private static object? GetPropertyValue(object obj, string propertyPath)
    {
        var current = obj;
        var parts = propertyPath.Split('.');

        foreach (var part in parts)
        {
            if (current == null)
                return null;

            var property = current.GetType().GetProperty(part);
            if (property == null)
                return null;

            current = property.GetValue(current);
        }

        return current;
    }

    /// <summary>
    /// Inserts an image into the slide at the text element's position
    /// </summary>
    private static void InsertImage(SlidePart slidePart, A.Text textElement, byte[] imageData, int? width, int? height, bool preserveAspectRatio)
    {
        try
        {
            // Validate image data first
            if (imageData == null || imageData.Length == 0)
            {
                Logger.Warning("PowerPointFunctionHandler: Image data is null or empty");
                return;
            }

            // Find the shape containing this text element
            var shape = textElement.Ancestors<Shape>().FirstOrDefault();
            if (shape == null)
            {
                Logger.Debug("PowerPointFunctionHandler: Could not find parent shape for text element");
                return;
            }            // Add image part to slide with proper error handling
            ImagePart imagePart;
            try
            {
                imagePart = slidePart.AddImagePart(ImagePartType.Png);
            }
            catch (Exception ex)
            {
                Logger.Error($"PowerPointFunctionHandler: Failed to add image part: {ex.Message}");
                return;
            }

            // Feed image data with proper stream handling
            try
            {
                using (var stream = new MemoryStream(imageData))
                {
                    stream.Position = 0;
                    imagePart.FeedData(stream);
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"PowerPointFunctionHandler: Failed to feed image data: {ex.Message}");
                try
                {
                    slidePart.DeletePart(imagePart);
                }
                catch
                {
                    // Ignore cleanup errors
                }
                return;
            }

            var imagePartId = slidePart.GetIdOfPart(imagePart);

            // Get shape properties for positioning
            var spPr = shape.ShapeProperties;
            var transform = spPr?.Transform2D;
              
            // Use existing shape dimensions if width/height not specified
            var finalWidth = width ?? (int?)((transform?.Extents?.Cx?.Value / 9525) ?? 200);
            var finalHeight = height ?? (int?)((transform?.Extents?.Cy?.Value / 9525) ?? 150);

            // Validate dimensions
            if (finalWidth <= 0) finalWidth = 200;
            if (finalHeight <= 0) finalHeight = 150;

            // Create picture element
            var picture = CreatePicture(imagePartId, finalWidth.Value, finalHeight.Value, preserveAspectRatio);
            
            // Replace the shape with the picture
            var parent = shape.Parent;
            if (parent != null)
            {
                try
                {
                    parent.ReplaceChild(picture, shape);
                    Logger.Debug($"PowerPointFunctionHandler: Successfully inserted image with dimensions {finalWidth}x{finalHeight}");
                }
                catch (Exception ex)
                {
                    Logger.Error($"PowerPointFunctionHandler: Failed to replace shape with picture: {ex.Message}");
                    try
                    {
                        slidePart.DeletePart(imagePart);
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                    throw;
                }
            }
            else
            {
                Logger.Warning("PowerPointFunctionHandler: Shape parent is null, cannot replace with picture");
                try
                {
                    slidePart.DeletePart(imagePart);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
        catch (Exception ex)
        {
            Logger.Error($"PowerPointFunctionHandler: Error inserting image: {ex.Message}");
            throw;
        }
    }    /// <summary>
    /// Creates a picture element with the specified properties
    /// </summary>
    private static Picture CreatePicture(string imagePartId, int width, int height, bool preserveAspectRatio)
    {
        var widthEmu = width * 9525L;
        var heightEmu = height * 9525L;

        return new Picture(
            new NonVisualPictureProperties(
                new NonVisualDrawingProperties { Id = 1, Name = "Picture" },
                new NonVisualPictureDrawingProperties(
                    new A.PictureLocks { NoChangeAspect = preserveAspectRatio }
                )
            ),
            new BlipFill(
                new A.Blip { Embed = imagePartId },
                new A.Stretch(new A.FillRectangle())
            ),
            new ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = widthEmu, Cy = heightEmu }
                ),
                new A.PresetGeometry(
                    new A.AdjustValueList()
                ) { Preset = A.ShapeTypeValues.Rectangle }
            )
        );
    }
}
