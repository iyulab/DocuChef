namespace DocuChef.Presentation.Functions;

/// <summary>
/// PowerPoint-specific functions that can be called from DollarSignEngine expressions
/// </summary>
public class PPTFunctions
{
    private readonly Dictionary<string, object> _variables;
    private readonly Dictionary<string, byte[]> _imageCache = new();

    public PPTFunctions(Dictionary<string, object> variables)
    {
        _variables = variables;
    }

    /// <summary>
    /// Returns a placeholder for image insertion
    /// The actual image insertion will be handled by PowerPointFunctionHandler
    /// </summary>
    /// <param name="propertyPath">Path to the image data property</param>
    /// <returns>Image placeholder text</returns>
    public string Image(string propertyPath)
    {
        try
        {
            // Get image data from variables
            var imageData = GetImageData(propertyPath);
            if (imageData == null)
            {
                Logger.Debug($"PPTFunctions: No image data found for property '{propertyPath}'");
                return $"[Missing Image: {propertyPath}]";
            }

            // Return a unique placeholder that will be processed later
            var placeholder = $"__PPT_IMAGE_{Guid.NewGuid():N}__{propertyPath}__";
            _imageCache[placeholder] = imageData;
            
            Logger.Debug($"PPTFunctions: Created image placeholder for '{propertyPath}'");
            return placeholder;
        }
        catch (Exception ex)
        {
            Logger.Error($"PPTFunctions: Error processing image '{propertyPath}' - {ex.Message}");
            return $"[Image Error: {propertyPath}]";
        }
    }

    /// <summary>
    /// Returns a placeholder for image insertion with size parameters
    /// </summary>
    /// <param name="propertyPath">Path to the image data property</param>
    /// <param name="width">Image width in pixels</param>
    /// <param name="height">Image height in pixels</param>
    /// <param name="preserveAspectRatio">Whether to preserve aspect ratio</param>
    /// <returns>Image placeholder text</returns>
    public string Image(string propertyPath, int width, int height, bool preserveAspectRatio = true)
    {
        try
        {
            var imageData = GetImageData(propertyPath);
            if (imageData == null)
            {
                Logger.Debug($"PPTFunctions: No image data found for property '{propertyPath}'");
                return $"[Missing Image: {propertyPath}]";
            }

            var placeholder = $"__PPT_IMAGE_{Guid.NewGuid():N}__{propertyPath}__{width}__{height}__{preserveAspectRatio}__";
            _imageCache[placeholder] = imageData;
            
            Logger.Debug($"PPTFunctions: Created image placeholder for '{propertyPath}' with size {width}x{height}");
            return placeholder;
        }
        catch (Exception ex)
        {
            Logger.Error($"PPTFunctions: Error processing image '{propertyPath}' - {ex.Message}");
            return $"[Image Error: {propertyPath}]";
        }
    }

    /// <summary>
    /// Gets the cached image data for a placeholder
    /// </summary>
    /// <param name="placeholder">The image placeholder</param>
    /// <returns>Image data or null if not found</returns>
    public byte[]? GetCachedImageData(string placeholder)
    {
        return _imageCache.TryGetValue(placeholder, out var data) ? data : null;
    }

    /// <summary>
    /// Gets all image placeholders and their data
    /// </summary>
    /// <returns>Dictionary of placeholders and image data</returns>
    public Dictionary<string, byte[]> GetAllImageCache()
    {
        return new Dictionary<string, byte[]>(_imageCache);
    }

    /// <summary>
    /// Retrieves image data from variables using property path
    /// </summary>
    private byte[]? GetImageData(string propertyPath)
    {
        try
        {
            var value = GetPropertyValue(_variables, propertyPath);
            return value switch
            {
                byte[] bytes => bytes,
                string path when File.Exists(path) => File.ReadAllBytes(path),
                _ => null
            };
        }
        catch (Exception ex)
        {
            Logger.Error($"PPTFunctions: Error getting image data for '{propertyPath}' - {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Gets property value using nested property path
    /// </summary>
    private static object? GetPropertyValue(Dictionary<string, object> variables, string propertyPath)
    {
        if (variables.TryGetValue(propertyPath, out var directValue))
        {
            return directValue;
        }

        // Handle nested property paths like "Object.Property"
        var parts = propertyPath.Split('.');
        if (parts.Length == 1)
        {
            return null;
        }

        var current = variables.TryGetValue(parts[0], out var rootValue) ? rootValue : null;
        
        for (int i = 1; i < parts.Length && current != null; i++)
        {
            var property = current.GetType().GetProperty(parts[i]);
            if (property == null)
                return null;
            
            current = property.GetValue(current);
        }

        return current;
    }
}
