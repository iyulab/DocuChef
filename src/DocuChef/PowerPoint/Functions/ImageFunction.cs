using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Functions;

/// <summary>
/// Image-related functions for PowerPoint processing according to PPT syntax guidelines
/// </summary>
internal static class ImageFunction
{
    /// <summary>
    /// Creates a PowerPoint function for image handling
    /// </summary>
    public static PowerPointFunction Create()
    {
        return new PowerPointFunction
        {
            Name = "Image",
            Description = "Inserts an image into a PowerPoint shape according to ppt.Image syntax",
            Handler = ProcessImageFunction
        };
    }

    /// <summary>
    /// Process image function: ppt.Image("imageProperty", width: 300, height: 200, preserveAspectRatio: true)
    /// </summary>
    private static object ProcessImageFunction(PowerPointContext context, object value, string[] parameters)
    {
        if (parameters == null || parameters.Length == 0)
        {
            Logger.Warning("Image function called without required path parameter");
            return "[Error: Image path required]";
        }

        string imagePath = parameters[0];
        Logger.Debug($"[IMAGE-DEBUG] Processing image with parameter: '{imagePath}'");

        try
        {
            int width = context.Options?.DefaultImageWidth ?? 300;
            int height = context.Options?.DefaultImageHeight ?? 200;
            bool preserveAspectRatio = context.Options?.PreserveImageAspectRatio ?? true;

            // Check for array references in parameters and add information to shape properties
            var arrayMatch = System.Text.RegularExpressions.Regex.Match(imagePath, @"(\w+)\[(\d+)\]");
            if (arrayMatch.Success && arrayMatch.Groups.Count >= 3)
            {
                string arrayName = arrayMatch.Groups[1].Value;
                if (int.TryParse(arrayMatch.Groups[2].Value, out int index))
                {
                    // Check if index is valid for this array
                    if (context.Variables.TryGetValue(arrayName, out var arrayObj) && arrayObj != null)
                    {
                        int arraySize = CollectionHelper.GetCollectionCount(arrayObj);
                        if (index >= arraySize)
                        {
                            Logger.Warning($"[IMAGE-DEBUG] Array index out of bounds: {arrayName}[{index}] >= {arraySize}");

                            // Hide the shape if it exists and index is out of bounds
                            if (context.Shape?.ShapeObject != null)
                            {
                                PowerPointShapeHelper.HideShape(context.Shape.ShapeObject);
                                Logger.Debug($"[IMAGE-DEBUG] Hiding shape due to invalid array index");
                                return ""; // Return empty string to avoid error message
                            }

                            return ""; // Return empty string instead of error message
                        }
                    }

                    // Add metadata to shape if valid
                    if (context.Shape?.ShapeObject != null)
                    {
                        var nvProps = context.Shape.ShapeObject.NonVisualShapeProperties;
                        if (nvProps?.NonVisualDrawingProperties != null)
                        {
                            // Add array reference to shape description (alt text)
                            string arrayRef = $"{arrayName}[{index}]";
                            nvProps.NonVisualDrawingProperties.Description = arrayRef;
                            Logger.Debug($"[IMAGE-DEBUG] Added array reference to shape description: {arrayRef}");
                        }
                    }
                }
            }

            // Resolve array references
            if (imagePath.Contains('[') && imagePath.Contains(']'))
            {
                string resolvedPath = ResolveArrayIndexedPath(context, imagePath);

                // Check if resolution resulted in an error
                if (resolvedPath.StartsWith("[Error:"))
                {
                    Logger.Warning($"[IMAGE-DEBUG] Failed to resolve array path: {resolvedPath}");

                    // Hide the shape if it exists
                    if (context.Shape?.ShapeObject != null)
                    {
                        PowerPointShapeHelper.HideShape(context.Shape.ShapeObject);
                        Logger.Debug($"[IMAGE-DEBUG] Hiding shape due to array resolution error");
                        return ""; // Return empty string to avoid error message
                    }

                    return ""; // Return empty string instead of error message
                }

                imagePath = resolvedPath;
            }
            // Resolve property paths
            else if (imagePath.Contains('.'))
            {
                imagePath = ResolvePropertyPath(context, imagePath);
            }
            // Direct variable reference
            else if (context.Variables.TryGetValue(imagePath, out var pathObj))
            {
                imagePath = pathObj?.ToString();
            }

            // Use ImageHelper if available
            if (ClosedXML.Report.XLCustom.Functions.ImageHelper.GetImageFromPathOrUrl != null)
            {
                try
                {
                    string resolvedImagePath = ClosedXML.Report.XLCustom.Functions.ImageHelper.GetImageFromPathOrUrl(imagePath);
                    if (!string.IsNullOrEmpty(resolvedImagePath))
                    {
                        Logger.Debug($"[IMAGE-DEBUG] Image resolved with ImageHelper: {resolvedImagePath}");
                        imagePath = resolvedImagePath;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"[IMAGE-DEBUG] Error using ImageHelper: {ex.Message}");
                }
            }

            // Parse named parameters
            ParseImageParameters(parameters, ref width, ref height, ref preserveAspectRatio);

            // Validate image file exists
            if (!File.Exists(imagePath))
            {
                Logger.Warning($"[IMAGE-DEBUG] Image file not found: {imagePath}");

                // Hide the shape if it exists
                if (context.Shape?.ShapeObject != null)
                {
                    PowerPointShapeHelper.HideShape(context.Shape.ShapeObject);
                    Logger.Debug($"[IMAGE-DEBUG] Hiding shape due to missing image file");
                    return ""; // Return empty string to avoid error message
                }

                return $""; // Return empty string instead of error message
            }

            var fileInfo = new FileInfo(imagePath);
            Logger.Debug($"Image file exists: {imagePath}, size: {fileInfo.Length} bytes");

            // Process the image in the shape
            if (context.Shape?.ShapeObject != null && context.SlidePart != null)
            {
                Logger.Debug($"Processing image in shape: File={imagePath}, Width={width}, Height={height}, PreserveAspectRatio={preserveAspectRatio}");
                Logger.Debug($"Shape ID: {context.Shape.ShapeObject.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value}");
                Logger.Debug($"Shape Name: {context.Shape.Name}");

                bool success = ProcessImageInShape(context.SlidePart, context.Shape.ShapeObject, imagePath, width, height, preserveAspectRatio);

                if (success)
                {
                    Logger.Info("Successfully processed image in shape");
                    return "";
                }
                else
                {
                    Logger.Warning("Failed to process image in shape");

                    // Hide the shape on processing failure
                    PowerPointShapeHelper.HideShape(context.Shape.ShapeObject);
                    Logger.Debug($"[IMAGE-DEBUG] Hiding shape due to image processing failure");
                    return "";
                }
            }

            Logger.Warning($"Invalid context for image processing - shape: {context.Shape?.ShapeObject != null}, slide part: {context.SlidePart != null}");
            return "";
        }
        catch (Exception ex)
        {
            Logger.Error($"Error processing image: {ex.Message}", ex);

            // Hide the shape on exception
            if (context.Shape?.ShapeObject != null)
            {
                PowerPointShapeHelper.HideShape(context.Shape.ShapeObject);
                Logger.Debug($"[IMAGE-DEBUG] Hiding shape due to exception: {ex.Message}");
            }

            return "";
        }
    }

    /// <summary>
    /// Resolve array indexed path (e.g., Items[0].ImageUrl)
    /// </summary>
    private static string ResolveArrayIndexedPath(PowerPointContext context, string path)
    {
        var match = System.Text.RegularExpressions.Regex.Match(path, @"^(\w+)\[(\d+)\](\.(.+))?$");
        if (!match.Success)
            return path;

        string arrayName = match.Groups[1].Value;
        int index = int.Parse(match.Groups[2].Value);
        var propertyPath = match.Groups[4].Success ? match.Groups[4].Value : null;

        Logger.Debug($"[IMAGE-DEBUG] Detected array reference: array={arrayName}, index={index}, property={propertyPath}");

        if (!context.Variables.TryGetValue(arrayName, out var arrayObj) || arrayObj == null)
        {
            Logger.Warning($"[IMAGE-DEBUG] Array not found: {arrayName}");
            return $"[Error: Array not found: {arrayName}]";
        }

        // Verify array index is within bounds
        int arraySize = CollectionHelper.GetCollectionCount(arrayObj);
        if (index >= arraySize)
        {
            Logger.Warning($"[IMAGE-DEBUG] Array index out of bounds: {arrayName}[{index}] >= {arraySize}");
            return $"[Error: Array index out of bounds: {index} >= {arraySize}]";
        }

        object item = CollectionHelper.GetItemAtIndex(arrayObj, index);
        if (item == null)
        {
            Logger.Warning($"[IMAGE-DEBUG] Array item not found: {arrayName}[{index}]");
            return $"[Error: Array item not found: {arrayName}[{index}]]";
        }

        Logger.Debug($"[IMAGE-DEBUG] Successfully retrieved item at index {index}");

        if (string.IsNullOrEmpty(propertyPath))
            return item.ToString();

        object propValue = ResolveNestedProperty(item, propertyPath);
        if (propValue == null)
        {
            Logger.Warning($"[IMAGE-DEBUG] Property '{propertyPath}' not found or null");
            return $"[Error: Property '{propertyPath}' not found]";
        }

        return propValue.ToString();
    }

    /// <summary>
    /// Resolve property path
    /// </summary>
    private static string ResolvePropertyPath(PowerPointContext context, string path)
    {
        var resolvedPath = context.ResolveVariable(path);
        if (resolvedPath != null)
        {
            Logger.Debug($"[IMAGE-DEBUG] Resolved image path from property path: {resolvedPath}");
            return resolvedPath.ToString();
        }
        return path;
    }

    /// <summary>
    /// Resolve nested property from object
    /// </summary>
    private static object? ResolveNestedProperty(object obj, string propertyPath)
    {
        var props = propertyPath.Split('.');
        object? current = obj;

        foreach (var prop in props)
        {
            if (current == null)
                return null;

            var property = current.GetType().GetProperty(prop);
            if (property == null)
                return null;

            current = property.GetValue(current);
        }

        return current;
    }

    /// <summary>
    /// Parse image parameters
    /// </summary>
    private static void ParseImageParameters(string[] parameters, ref int width, ref int height, ref bool preserveAspectRatio)
    {
        for (int i = 1; i < parameters.Length; i++)
        {
            string param = parameters[i];
            int colonIndex = param.IndexOf(':');
            if (colonIndex <= 0)
                continue;

            string paramName = param.Substring(0, colonIndex).Trim();
            string paramValue = param.Substring(colonIndex + 1).Trim();

            switch (paramName.ToLowerInvariant())
            {
                case "width":
                    if (int.TryParse(paramValue, out int w))
                        width = w;
                    break;
                case "height":
                    if (int.TryParse(paramValue, out int h))
                        height = h;
                    break;
                case "preserveaspectratio":
                    if (bool.TryParse(paramValue, out bool p))
                        preserveAspectRatio = p;
                    break;
            }
        }
    }

    /// <summary>
    /// Process image in shape by replacing the shape with a picture
    /// </summary>
    private static bool ProcessImageInShape(SlidePart slidePart, P.Shape shape, string imagePath, int width, int height, bool preserveAspectRatio)
    {
        try
        {
            Logger.Debug($"Starting ProcessImageInShape: Path={imagePath}, Shape ID={shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value}");

            var nvsp = shape.NonVisualShapeProperties;
            var nvdp = nvsp?.NonVisualDrawingProperties;
            var transform = shape.ShapeProperties?.Transform2D;

            uint shapeId = nvdp?.Id?.Value ?? 1000u;
            string shapeName = nvdp?.Name?.Value ?? "Image_Shape";
            long shapeX = transform?.Offset?.X?.Value ?? 1524000;
            long shapeY = transform?.Offset?.Y?.Value ?? 1524000;
            long shapeWidth = transform?.Extents?.Cx?.Value ?? (long)(width * 9525);
            long shapeHeight = transform?.Extents?.Cy?.Value ?? (long)(height * 9525);

            Logger.Debug($"Original shape: ID={shapeId}, Name={shapeName}, X={shapeX}, Y={shapeY}, Width={shapeWidth}, Height={shapeHeight}");

            string contentType = Path.GetExtension(imagePath).GetContentType();
            if (string.IsNullOrEmpty(contentType))
            {
                Logger.Warning($"Unsupported image format: {Path.GetExtension(imagePath)}");
                return false;
            }

            string relationshipId = GenerateUniqueRelationshipId(slidePart);
            Logger.Debug($"Generated relationship ID: {relationshipId}");

            var imagePart = PowerPointHelper.CreateImagePart(slidePart, contentType, relationshipId);
            if (imagePart == null)
            {
                Logger.Error("Failed to create image part");
                return false;
            }

            using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
            {
                imagePart.FeedData(stream);
            }
            Logger.Debug("Image data fed to part successfully");

            A.Outline outline = null;
            var originalOutline = shape.ShapeProperties?.GetFirstChild<A.Outline>();

            if (originalOutline != null)
            {
                outline = PowerPointHelper.CloneOutline(originalOutline);
                Logger.Debug("Outline cloned from original shape");
            }
            else
            {
                outline = PowerPointHelper.CreateDefaultOutline(9525, "808080");
                Logger.Debug("Created default outline");
            }

            // Preserve description (alt text) from original shape
            string description = nvdp?.Description?.Value;

            var picture = PowerPointHelper.CreatePicture(
                relationshipId,
                shapeId, // Keep the same ID for proper replacement
                shapeName, // Keep original shape name
                shapeX,
                shapeY,
                shapeWidth,
                shapeHeight,
                preserveAspectRatio,
                outline
            );

            // Set description (alt text) on new picture from original shape
            if (!string.IsNullOrEmpty(description))
            {
                picture.NonVisualPictureProperties.NonVisualDrawingProperties.Description = description;
            }

            var parent = shape.Parent;
            if (parent == null)
            {
                Logger.Error("Cannot find parent element for the shape");
                return false;
            }

            parent.InsertAfter(picture, shape);
            parent.RemoveChild(shape);

            Logger.Debug("Replaced original shape with new Picture element");
            return true;
        }
        catch (Exception ex)
        {
            Logger.Error($"Error replacing shape with image: {ex.Message}", ex);
            return false;
        }
    }

    /// <summary>
    /// Generate unique relationship ID for slide part
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
    }
}