using System.Reflection;

namespace DocuChef.Presentation.Core;

/// <summary>
/// Provides extension methods for accessing object properties dynamically
/// </summary>
public static class ObjectExtensions
{
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

                // Get property using reflection with case-insensitive lookup
                PropertyInfo property = GetPropertyInfo(currentObj, propertyName);

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
    /// Gets a property info by name with case-insensitive lookup
    /// </summary>
    private static PropertyInfo GetPropertyInfo(object obj, string propertyName)
    {
        return obj.GetType().GetProperty(
            propertyName,
            BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
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
            // Handle different collection accessing patterns
            if (propertyPath.Contains("_"))
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
        IEnumerable topCollection = GetRegularCollection(obj, topLevelCollectionName);

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
                var childProp = GetPropertyInfo(item, childPropertyName);

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

            // Get property using reflection
            var property = GetPropertyInfo(currentObj, propertyName);

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