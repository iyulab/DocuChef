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
    }    /// <summary>
         /// Gets all image placeholders and their data
         /// </summary>
         /// <returns>Dictionary of placeholders and image data</returns>
    public Dictionary<string, byte[]> GetAllImageCache()
    {
        Logger.Debug($"PPTFunctions.GetAllImageCache: Returning {_imageCache.Count} cached items");
        foreach (var item in _imageCache)
        {
            Logger.Debug($"PPTFunctions.GetAllImageCache: Key='{item.Key}', Value length={item.Value.Length}");
        }
        return new Dictionary<string, byte[]>(_imageCache);
    }/// <summary>
     /// Retrieves image data from variables using property path
     /// </summary>
    private byte[]? GetImageData(string propertyPath)
    {
        try
        {
            Logger.Debug($"PPTFunctions: GetImageData called with propertyPath='{propertyPath}'");
            var value = GetPropertyValue(_variables, propertyPath);
            Logger.Debug($"PPTFunctions: GetPropertyValue returned: {value?.GetType().Name ?? "null"} = '{value}'");

            // If we found a value in variables, process it
            if (value != null)
            {
                var result = value switch
                {
                    byte[] bytes => bytes,
                    string filePath when File.Exists(filePath) => File.ReadAllBytes(filePath),
                    _ => null
                };

                if (result != null)
                {
                    Logger.Debug($"PPTFunctions: GetImageData returning byte array with {result.Length} bytes from variable");
                    return result;
                }
            }

            // If not found in variables, treat propertyPath as a direct file path
            if (File.Exists(propertyPath))
            {
                Logger.Debug($"PPTFunctions: propertyPath '{propertyPath}' is a valid file, reading directly");
                var fileData = File.ReadAllBytes(propertyPath);
                Logger.Debug($"PPTFunctions: GetImageData returning byte array with {fileData.Length} bytes from file");
                return fileData;
            }

            Logger.Debug($"PPTFunctions: File.Exists('{propertyPath}') = false");
            Logger.Debug($"PPTFunctions: GetImageData returning null");
            return null;
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

    /// <summary>
    /// Restores image cache from another PPTFunctions instance
    /// </summary>
    /// <param name="imageCache">Image cache to restore</param>
    public void RestoreImageCache(Dictionary<string, byte[]> imageCache)
    {
        if (imageCache == null) return;

        foreach (var item in imageCache)
        {
            _imageCache[item.Key] = item.Value;
        }

        Logger.Debug($"PPTFunctions: Restored {imageCache.Count} cached images");
    }
}
