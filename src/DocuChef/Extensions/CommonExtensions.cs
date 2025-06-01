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
                var key = entry.Key?.ToString();
                if (key != null)
                {
                    result[key] = entry.Value ?? string.Empty;
                }
            }
            return result;
        }

        var resultDict = new Dictionary<string, object>();
        var type = source.GetType();        // Handle ExpandoObject specially
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
                    resultDict[prop.Name] = value ?? string.Empty;
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
                resultDict[field.Name] = value ?? string.Empty;
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
}