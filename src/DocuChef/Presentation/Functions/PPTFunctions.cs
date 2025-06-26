using System.Text.Json;

namespace DocuChef.Presentation.Functions;

/// <summary>
/// PowerPoint-specific functions that can be called from DollarSignEngine expressions
/// </summary>
public class PPTFunctions
{
    private readonly Dictionary<string, object> _variables;
    private readonly Dictionary<string, string> _imageCache = new();

    public PPTFunctions(Dictionary<string, object> variables)
    {
        _variables = variables;
    }

    /// <summary>
    /// Returns a placeholder for image insertion
    /// The actual image insertion will be handled by PowerPointFunctionHandler
    /// </summary>
    public string? Image(object propertyPathObj)
    {
        if (propertyPathObj == null) return null;

        try
        {
            // Convert the input to string, handling various types including JsonElement
            string propertyPath;
            if (propertyPathObj is System.Text.Json.JsonElement jsonElement)
            {
                propertyPath = jsonElement.GetString() ?? jsonElement.ToString();
            }
            else
            {
                propertyPath = propertyPathObj?.ToString() ?? "";
            }

            Logger.Debug($"PPTFunctions.Image: Received {propertyPathObj?.GetType().Name ?? "null"}, converted to '{propertyPath}'");

            // Generate a unique GUID for this image placeholder
            var guid = Guid.NewGuid().ToString("N");
            var placeholder = $"__PPT_IMAGE_{guid}__{propertyPath}__";

            if (_imageCache.ContainsKey(placeholder))
            {
                Logger.Debug($"PPTFunctions: Returning cached image placeholder for '{propertyPath}'");
                return placeholder;
            }            // Download/resolve the actual image and cache it
            var imagePath = ClosedXML.Report.XLCustom.Functions.ImageHelper.GetImageFromPathOrUrl(propertyPath) ?? string.Empty;
            _imageCache[placeholder] = imagePath;

            Logger.Debug($"PPTFunctions: Created image placeholder '{placeholder}' for path '{propertyPath}'");
            return placeholder;
        }
        catch (Exception ex)
        {
            Logger.Error($"PPTFunctions: Error processing image '{propertyPathObj}' - {ex.Message}");
            return $"[Image Error: {propertyPathObj}]";
        }
    }

    /// <summary>
    /// Gets all image placeholders and their data
    /// </summary>
    /// <returns>Dictionary of placeholders and image data</returns>
    public Dictionary<string, string> GetAllImageCache()
    {
        Logger.Debug($"PPTFunctions.GetAllImageCache: Returning {_imageCache.Count} cached items");
        foreach (var item in _imageCache)
        {
            if (item.Value == null) continue;
            Logger.Debug($"PPTFunctions.GetAllImageCache: Key='{item.Key}', Value length={item.Value.Length}");
        }
        return new Dictionary<string, string>(_imageCache);
    }

    /// <summary>
    /// Restores image cache from another PPTFunctions instance
    /// </summary>
    /// <param name="imageCache">Image cache to restore</param>
    public void RestoreImageCache(Dictionary<string, string> imageCache)
    {
        if (imageCache == null) return;

        foreach (var item in imageCache)
        {
            _imageCache[item.Key] = item.Value;
        }

        Logger.Debug($"PPTFunctions: Restored {imageCache.Count} cached images");
    }
}
