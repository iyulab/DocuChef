using System.Reflection;

namespace DocuChef.Extensions;

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

        // If already a dictionary, convert it
        if (source is IDictionary<string, object> dictionary)
        {
            return new Dictionary<string, object>(dictionary);
        }
        else if (source is IDictionary genericDict)
        {
            var result = new Dictionary<string, object>();
            foreach (DictionaryEntry entry in genericDict)
            {
                result[entry.Key.ToString()] = entry.Value;
            }
            return result;
        }

        var resultDict = new Dictionary<string, object>();
        var type = source.GetType();

        // Handle ExpandoObject specially
        if (source is System.Dynamic.ExpandoObject expandoObj)
        {
            return new Dictionary<string, object>(expandoObj as IDictionary<string, object>);
        }

        // Get public properties
        foreach (var prop in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (prop.CanRead)
            {
                try
                {
                    var value = prop.GetValue(source);
                    resultDict[prop.Name] = value;
                }
                catch
                {
                    // Skip properties that throw exceptions
                }
            }
        }

        // Get public fields
        foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance))
        {
            try
            {
                var value = field.GetValue(source);
                resultDict[field.Name] = value;
            }
            catch
            {
                // Skip fields that throw exceptions
            }
        }

        return resultDict;
    }

    /// <summary>
    /// Determines if an object is a complex type (not a simple value type)
    /// </summary>
    public static bool IsComplexType(this object obj)
    {
        if (obj == null)
            return false;

        var type = obj.GetType();

        // Simple types aren't complex
        if (obj is string || obj is ValueType)
            return false;

        // Collections that aren't dictionaries aren't complex
        if (obj is IEnumerable && !(obj is IDictionary))
            return false;

        return true;
    }

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
    /// Recursively processes nested collection items
    /// </summary>
    private static void ProcessNestedCollectionItems(
        System.Collections.ArrayList result,
        IEnumerable currentCollection,
        string[] segments,
        int currentLevel,
        List<object> ancestry)
    {
        foreach (var item in currentCollection)
        {
            if (item == null)
                continue;

            // Create a new ancestry chain for this branch
            var currentAncestry = new List<object>(ancestry);
            currentAncestry.Add(item);

            // If we've reached the target nesting level, add the item to results
            if (currentLevel == segments.Length - 1)
            {
                // Create a nested item with full hierarchy information
                var nestedItem = new DocuChef.Presentation.Core.NestedCollectionItem
                {
                    Item = item,
                    HierarchyPath = segments,
                    Ancestry = currentAncestry.ToArray()
                };

                result.Add(nestedItem);
            }
            // If we're not at the target level yet, continue down the hierarchy
            else if (currentLevel < segments.Length - 1)
            {
                // Get the child property name from the next segment
                string childPropertyName = segments[currentLevel + 1];

                // Try to get the child collection property
                object childValue = null;

                // Try dictionary first
                if (item is IDictionary<string, object> dict && dict.ContainsKey(childPropertyName))
                {
                    childValue = dict[childPropertyName];
                }
                else
                {
                    // Try reflection
                    var childProp = item.GetType().GetProperty(
                        childPropertyName,
                        BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

                    if (childProp == null)
                    {
                        Logger.Debug($"Child property not found: {childPropertyName} on {item.GetType().Name}");
                        continue;
                    }

                    // Get the child collection value
                    childValue = childProp.GetValue(item);
                }

                // If it's a collection (but not a string), process it recursively
                if (childValue is IEnumerable childCollection && !(childValue is string))
                {
                    ProcessNestedCollectionItems(result, childCollection, segments, currentLevel + 1, currentAncestry);
                }
            }
        }
    }
}