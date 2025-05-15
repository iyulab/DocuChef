namespace DocuChef.Helpers;

/// <summary>
/// Helper methods for collection operations
/// </summary>
public static class CollectionHelper
{
    /// <summary>
    /// Get collection count for any object
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

        // Handle generic collections
        if (obj is IList list)
            return list.Count;

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
    /// Get item at index from any collection
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

        // Handle lists
        if (collection is IList list)
        {
            return index < list.Count ? list[index] : null;
        }

        // Handle indexer via reflection
        var indexerProperty = collection.GetType().GetProperty("Item");
        if (indexerProperty != null)
        {
            var parameters = indexerProperty.GetIndexParameters();
            if (parameters.Length == 1 && parameters[0].ParameterType == typeof(int))
            {
                try
                {
                    return indexerProperty.GetValue(collection, new object[] { index });
                }
                catch
                {
                    return null;
                }
            }
        }

        // Handle IEnumerable
        if (collection is IEnumerable enumerable)
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
}