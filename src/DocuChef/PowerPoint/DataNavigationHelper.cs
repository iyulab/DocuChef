using System.Reflection;

namespace DocuChef.PowerPoint;

/// <summary>
/// Helper class for navigating data objects and collections
/// </summary>
internal static class DataNavigationHelper
{
    /// <summary>
    /// Gets a property value from an object
    /// </summary>
    public static object GetPropertyValue(object obj, string propertyName)
    {
        if (obj == null || string.IsNullOrEmpty(propertyName))
            return null;

        // Check if the object is a dictionary
        if (obj is IDictionary dictionary && dictionary.Contains(propertyName))
            return dictionary[propertyName];

        // Try to get property using reflection (case insensitive)
        var property = obj.GetType().GetProperty(propertyName,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

        if (property != null && property.CanRead)
        {
            try
            {
                return property.GetValue(obj);
            }
            catch
            {
                // Ignore exceptions, return null
            }
        }

        // Try to get field using reflection (case insensitive)
        var field = obj.GetType().GetField(propertyName,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

        if (field != null)
        {
            try
            {
                return field.GetValue(obj);
            }
            catch
            {
                // Ignore exceptions, return null
            }
        }

        return null;
    }

    /// <summary>
    /// Gets an item at the specified index from a collection
    /// </summary>
    public static object GetItemAtIndex(object collection, int index)
    {
        if (collection == null || index < 0)
            return null;

        // Handle arrays
        if (collection is Array array)
        {
            return index < array.Length ? array.GetValue(index) : null;
        }

        // Handle IList
        if (collection is IList list)
        {
            return index < list.Count ? list[index] : null;
        }

        // Handle IEnumerable
        if (collection is IEnumerable enumerable && !(collection is string))
        {
            int currentIndex = 0;
            foreach (var item in enumerable)
            {
                if (currentIndex == index)
                    return item;
                currentIndex++;
            }
        }

        return null;
    }

    /// <summary>
    /// Gets the count of items in a collection
    /// </summary>
    public static int GetCollectionCount(object obj)
    {
        if (obj == null)
            return 0;

        // Handle standard collections
        if (obj is ICollection collection)
            return collection.Count;

        if (obj is Array array)
            return array.Length;

        // Try Count property via reflection
        var countProperty = obj.GetType().GetProperty("Count");
        if (countProperty != null && countProperty.PropertyType == typeof(int))
        {
            try
            {
                return (int)countProperty.GetValue(obj);
            }
            catch
            {
                // Continue to other methods
            }
        }

        // Try Length property via reflection
        var lengthProperty = obj.GetType().GetProperty("Length");
        if (lengthProperty != null && lengthProperty.PropertyType == typeof(int))
        {
            try
            {
                return (int)lengthProperty.GetValue(obj);
            }
            catch
            {
                // Continue to other methods
            }
        }

        // For any IEnumerable, count by enumerating
        if (obj is IEnumerable enumerable)
        {
            int count = 0;
            foreach (var _ in enumerable)
                count++;
            return count;
        }

        return 0;
    }

    /// <summary>
    /// Gets all items from a collection as a list
    /// </summary>
    public static List<object> GetCollectionItems(object collection)
    {
        var result = new List<object>();

        if (collection == null)
            return result;

        // Convert various collection types to a list
        if (collection is IEnumerable enumerable && !(collection is string))
        {
            foreach (var item in enumerable)
            {
                result.Add(item);
            }
        }

        return result;
    }
}