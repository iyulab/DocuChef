using System.Reflection;

namespace DocuChef.Presentation.Core;

/// <summary>
/// Provides extension methods for accessing object properties dynamically
/// </summary>
internal static class ObjectExtensions
{
    // Cache for property info lookup to improve performance
    private static readonly Dictionary<Type, Dictionary<string, PropertyInfo>> _propertyCache =
        new Dictionary<Type, Dictionary<string, PropertyInfo>>();

    private static readonly object _cacheLock = new object();

    /// <summary>
    /// Gets a property value from an object by property path
    /// </summary>
    public static string GetPropertyValue(this object obj, string propertyPath)
    {
        if (obj == null || string.IsNullOrEmpty(propertyPath))
            return string.Empty;

        try
        {
            // Handle case of NestedCollectionItem first
            if (obj is NestedCollectionItem nestedItem)
            {
                obj = nestedItem.Item;
                if (obj == null)
                    return string.Empty;
            }

            // Handle nested properties (e.g., "Customer.Name")
            string[] propertyNames = propertyPath.Split('.');
            object currentObj = obj;

            foreach (string propertyName in propertyNames)
            {
                if (currentObj == null)
                    return string.Empty;

                // Get property using reflection with case-insensitive lookup and caching
                PropertyInfo property = GetCachedPropertyInfo(currentObj, propertyName);

                if (property == null)
                {
                    Logger.Debug($"Property not found: {propertyName} on {currentObj.GetType().Name}");
                    return string.Empty;
                }

                currentObj = property.GetValue(currentObj);
            }

            // Convert result to string
            return currentObj?.ToString() ?? string.Empty;
        }
        catch (Exception ex)
        {
            Logger.Debug($"Error accessing property '{propertyPath}': {ex.Message}");
            return string.Empty;
        }
    }

    /// <summary>
    /// Gets a property info by name with case-insensitive lookup and caching
    /// </summary>
    private static PropertyInfo GetCachedPropertyInfo(object obj, string propertyName)
    {
        if (obj == null || string.IsNullOrEmpty(propertyName))
            return null;

        Type objType = obj.GetType();

        // Check cache first
        Dictionary<string, PropertyInfo> typeCache;
        lock (_cacheLock)
        {
            if (!_propertyCache.TryGetValue(objType, out typeCache))
            {
                typeCache = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
                _propertyCache[objType] = typeCache;
            }
        }

        // Try to get from cache using thread-safe approach
        PropertyInfo result = null;
        bool cacheHit;

        lock (_cacheLock)
        {
            cacheHit = typeCache.TryGetValue(propertyName, out result);
        }

        if (!cacheHit)
        {
            // Not in cache, get using reflection
            result = objType.GetProperty(
                propertyName,
                BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);

            // Cache the result (even if null)
            lock (_cacheLock)
            {
                typeCache[propertyName] = result;
            }
        }

        return result;
    }

    /// <summary>
    /// Gets a collection from an object by property path, supporting multi-level nested collections
    /// </summary>
    public static IEnumerable GetCollection(this object obj, string propertyPath)
    {
        if (obj == null || string.IsNullOrEmpty(propertyPath))
            return null;

        try
        {
            // Special case: if obj is a Dictionary<string, object> and propertyPath is a key
            if (obj is IDictionary<string, object> dict && dict.ContainsKey(propertyPath))
            {
                var value = dict[propertyPath];
                if (value is IEnumerable && !(value is string))
                    return value as IEnumerable;
                return null;
            }

            // Handle different collection accessing patterns
            if (propertyPath.Contains('_'))
            {
                // Multi-level nested collection with underscore notation (e.g., "Categories_Products")
                return GetMultiLevelNestedCollection(obj, propertyPath);
            }
            else
            {
                // Regular property path (e.g., "Categories", "Customer.Orders")
                return GetRegularCollection(obj, propertyPath);
            }
        }
        catch (Exception ex)
        {
            Logger.Debug($"Error accessing collection '{propertyPath}': {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Gets a multi-level nested collection using the underscore notation (e.g., "Categories_Products")
    /// </summary>
    private static IEnumerable GetMultiLevelNestedCollection(object obj, string propertyPath)
    {
        // Split the property path by underscore to get hierarchy segments
        string[] segments = propertyPath.Split('_');
        if (segments.Length < 2)
            return null;

        Logger.Debug($"Processing multi-level nested collection: {propertyPath} with {segments.Length} levels");

        // Get the top-level collection
        string topLevelCollectionName = segments[0];
        IEnumerable topCollection = GetCollectionFromObject(obj, topLevelCollectionName);

        if (topCollection == null)
        {
            Logger.Debug($"Top-level collection '{topLevelCollectionName}' not found");
            return null;
        }

        // Create a result collection to store the nested items
        var result = new System.Collections.ArrayList();

        // Process each item in the top collection recursively
        ProcessNestedCollectionItems(result, topCollection, segments, 0, new List<object>());

        Logger.Debug($"Processed {result.Count} items in multi-level nested collection");

        return result.Count > 0 ? result : null;
    }

    /// <summary>
    /// Gets a collection from an object using various methods
    /// </summary>
    private static IEnumerable GetCollectionFromObject(object obj, string collectionName)
    {
        // Special case: if obj is a Dictionary<string, object> and collectionName is a key
        if (obj is IDictionary<string, object> dict && dict.ContainsKey(collectionName))
        {
            var value = dict[collectionName];
            if (value is IEnumerable && !(value is string))
                return value as IEnumerable;
            return null;
        }

        // Try regular collection method
        return GetRegularCollection(obj, collectionName);
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
                var nestedItem = new NestedCollectionItem
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
                PropertyInfo childProp = GetCachedPropertyInfo(item, childPropertyName);

                if (childProp == null)
                {
                    Logger.Debug($"Child property not found: {childPropertyName} on {item.GetType().Name}");
                    continue;
                }

                // Get the child collection value
                object childValue = childProp.GetValue(item);

                // If it's a collection (but not a string), process it recursively
                if (childValue is IEnumerable childCollection && !(childValue is string))
                {
                    ProcessNestedCollectionItems(result, childCollection, segments, currentLevel + 1, currentAncestry);
                }
            }
        }
    }

    /// <summary>
    /// Gets a regular collection using a property path with dot notation
    /// </summary>
    private static IEnumerable GetRegularCollection(object obj, string propertyPath)
    {
        // Unwrap NestedCollectionItem if necessary
        if (obj is NestedCollectionItem nestedItem)
        {
            obj = nestedItem.Item;
            if (obj == null)
                return null;
        }

        // Handle nested properties using dot notation (e.g., "Customer.Orders")
        string[] propertyNames = propertyPath.Split('.');
        object currentObj = obj;

        foreach (string propertyName in propertyNames)
        {
            if (currentObj == null)
                return null;

            // Get property using reflection with caching
            var property = GetCachedPropertyInfo(currentObj, propertyName);

            if (property == null)
            {
                Logger.Debug($"Property not found in collection path: {propertyName} on {currentObj.GetType().Name}");
                return null;
            }

            currentObj = property.GetValue(currentObj);
        }

        // Return result if it's a collection (but not a string)
        if (currentObj is IEnumerable enumerable && !(currentObj is string))
            return enumerable;

        return null;
    }

    /// <summary>
    /// Counts items in a collection
    /// </summary>
    public static int Count(this IEnumerable collection)
    {
        if (collection == null)
            return 0;

        // Use LINQ Count extension if possible
        if (collection is ICollection c)
            return c.Count;

        // Otherwise, manually count items
        int count = 0;
        foreach (var item in collection)
        {
            count++;
        }

        return count;
    }
}

/// <summary>
/// Represents an item in a nested collection with full hierarchy information
/// </summary>
internal class NestedCollectionItem
{
    /// <summary>
    /// The actual collection item
    /// </summary>
    public object Item { get; set; }

    /// <summary>
    /// The hierarchy path segments from root to this item
    /// </summary>
    public string[] HierarchyPath { get; set; }

    /// <summary>
    /// The ancestry chain from root to this item (array of parent objects)
    /// </summary>
    public object[] Ancestry { get; set; }

    /// <summary>
    /// Gets the parent object at the specified level (0-based)
    /// </summary>
    public object GetAncestor(int level)
    {
        if (Ancestry == null || level < 0 || level >= Ancestry.Length)
            return null;

        return Ancestry[level];
    }

    /// <summary>
    /// Gets the direct parent of this item
    /// </summary>
    public object Parent => Ancestry?.Length > 1 ? Ancestry[Ancestry.Length - 2] : null;

    /// <summary>
    /// Gets the nesting level of this item (0 for top level)
    /// </summary>
    public int NestingLevel => Ancestry != null ? Ancestry.Length - 1 : 0;

    /// <summary>
    /// Returns the string representation of the item
    /// </summary>
    public override string ToString()
    {
        return Item?.ToString() ?? string.Empty;
    }
}