using System;
using System.IO;

namespace DocuChef.Extensions;

/// <summary>
/// Extension methods for file operations
/// </summary>
public static class FileExtensions
{
    /// <summary>
    /// Ensures a directory exists for a file path
    /// </summary>
    public static string EnsureDirectoryExists(this string filePath)
    {
        if (string.IsNullOrEmpty(filePath))
            return filePath;

        string directory = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        return filePath;
    }

    /// <summary>
    /// Gets content type based on file extension
    /// </summary>
    public static string GetContentType(this string fileExtension)
    {
        if (string.IsNullOrEmpty(fileExtension))
            return null;

        if (fileExtension.StartsWith("."))
            fileExtension = fileExtension.Substring(1);

        return fileExtension.ToLowerInvariant() switch
        {
            "png" => "image/png",
            "jpg" => "image/jpeg",
            "jpeg" => "image/jpeg",
            "gif" => "image/gif",
            "bmp" => "image/bmp",
            "tiff" => "image/tiff",
            "tif" => "image/tiff",
            "xlsx" => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "xls" => "application/vnd.ms-excel",
            "pptx" => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
            "ppt" => "application/vnd.ms-powerpoint",
            "docx" => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            "doc" => "application/vnd.ms-word",
            "pdf" => "application/pdf",
            _ => null
        };
    }

    /// <summary>
    /// Creates a temporary file path with specified extension
    /// </summary>
    public static string GetTempFilePath(this string extension)
    {
        extension = extension.StartsWith(".") ? extension : $".{extension}";
        return Path.Combine(Path.GetTempPath(), $"DocuChef_{Guid.NewGuid().ToString("N")}{extension}");
    }

    /// <summary>
    /// Safely copies a stream to a file, ensuring directory exists
    /// </summary>
    public static void CopyToFile(this Stream source, string destination)
    {
        if (source == null)
            throw new ArgumentNullException(nameof(source));

        if (string.IsNullOrEmpty(destination))
            throw new ArgumentNullException(nameof(destination));

        destination.EnsureDirectoryExists();

        using var fileStream = new FileStream(destination, FileMode.Create, FileAccess.Write);
        source.Position = 0;
        source.CopyTo(fileStream);
    }
}

/// <summary>
/// Extension methods for collections and objects
/// </summary>
public static class CommonExtensions
{
    /// <summary>
    /// Gets a dictionary of properties and their values from an object
    /// </summary>
    public static Dictionary<string, object> GetProperties(this object source)
    {
        if (source == null)
            return new Dictionary<string, object>();

        var result = new Dictionary<string, object>();
        var type = source.GetType();

        // Get public properties
        foreach (var prop in type.GetProperties(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
        {
            if (prop.CanRead)
            {
                try
                {
                    var value = prop.GetValue(source);
                    result[prop.Name] = value;
                }
                catch
                {
                    // Skip properties that throw exceptions
                }
            }
        }

        // Get public fields
        foreach (var field in type.GetFields(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance))
        {
            try
            {
                var value = field.GetValue(source);
                result[field.Name] = value;
            }
            catch
            {
                // Skip fields that throw exceptions
            }
        }

        return result;
    }
}