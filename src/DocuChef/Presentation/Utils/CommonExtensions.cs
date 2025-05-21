namespace DocuChef.Presentation.Utils;

/// <summary>
/// Extension methods for collections and objects
/// </summary>
internal static class CommonExtensions
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