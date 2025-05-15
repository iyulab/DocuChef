namespace DocuChef.PowerPoint;

/// <summary>
/// Context for PowerPoint template processing
/// </summary>
public partial class PowerPointContext
{
    /// <summary>
    /// Current indices for array collections, used for slide batches and nested data references
    /// </summary>
    public Dictionary<string, int> CurrentIndices { get; set; } = new Dictionary<string, int>();

    /// <summary>
    /// Keeps track of processed nested collections for data binding
    /// </summary>
    public Dictionary<string, object> ProcessedNestedCollections { get; set; } = new Dictionary<string, object>();

    /// <summary>
    /// Resolves a nested data reference by parent and child names
    /// </summary>
    /// <param name="parentName">The name of the parent collection/object</param>
    /// <param name="childName">The name of the child property/collection</param>
    /// <returns>The resolved value or null if not found</returns>
    public object ResolveNestedData(string parentName, string childName)
    {
        if (string.IsNullOrEmpty(parentName) || string.IsNullOrEmpty(childName))
            return null;

        // Check if combined name exists as a direct variable
        string combinedName = $"{parentName}_{childName}";
        if (Variables.TryGetValue(combinedName, out var directValue))
            return directValue;

        // Try to resolve from parent object
        if (!Variables.TryGetValue(parentName, out var parentObj) || parentObj == null)
            return null;

        // Get current index for the parent collection
        int currentIndex = CurrentIndices.GetValueOrDefault(parentName, 0);
        Logger.Debug($"Resolving nested data: {parentName}[{currentIndex}].{childName}");

        // Get item at current index if parent is a collection
        object targetObj = parentObj;

        if (parentObj is IList list && currentIndex >= 0 && currentIndex < list.Count)
        {
            targetObj = list[currentIndex];
            Logger.Debug($"Resolved {parentName} to item at index {currentIndex}");
        }
        else if (parentObj is Array array && currentIndex >= 0 && currentIndex < array.Length)
        {
            targetObj = array.GetValue(currentIndex);
            Logger.Debug($"Resolved {parentName} to array item at index {currentIndex}");
        }
        else
        {
            Logger.Debug($"Using {parentName} directly as it is not a collection or index is out of range");
        }

        // Get child property from the target object
        if (targetObj != null)
        {
            var property = targetObj.GetType().GetProperty(childName);
            if (property != null && property.CanRead)
            {
                var result = property.GetValue(targetObj);
                Logger.Debug($"Resolved {parentName}[{currentIndex}].{childName} to {(result != null ? result.GetType().Name : "null")}");

                // Cache the result for future use
                ProcessedNestedCollections[$"{parentName}[{currentIndex}].{childName}"] = result;

                return result;
            }
            else
            {
                Logger.Warning($"Property {childName} not found or not readable on {targetObj.GetType().Name}");
            }
        }
        else
        {
            Logger.Warning($"Target object is null for {parentName}[{currentIndex}]");
        }

        return null;
    }

    /// <summary>
    /// Gets direct access to a nested collection at a specific path
    /// </summary>
    /// <param name="nestedPath">Path in format "Parent[index].Child"</param>
    /// <returns>The resolved collection or null if not found</returns>
    public object GetNestedCollection(string nestedPath)
    {
        if (string.IsNullOrEmpty(nestedPath))
            return null;

        // Check if we've already processed this path
        if (ProcessedNestedCollections.TryGetValue(nestedPath, out var cachedValue))
            return cachedValue;

        // Parse path: Parent[index].Child
        var match = System.Text.RegularExpressions.Regex.Match(nestedPath, @"(\w+)\[(\d+)\]\.(\w+)");
        if (!match.Success || match.Groups.Count < 4)
            return null;

        string parentName = match.Groups[1].Value;
        int parentIndex = int.Parse(match.Groups[2].Value);
        string childName = match.Groups[3].Value;

        // Store original index value
        int originalIndex = CurrentIndices.ContainsKey(parentName) ? CurrentIndices[parentName] : 0;

        try
        {
            // Set temporary index and resolve
            CurrentIndices[parentName] = parentIndex;
            return ResolveNestedData(parentName, childName);
        }
        finally
        {
            // Restore original index
            CurrentIndices[parentName] = originalIndex;
        }
    }

    /// <summary>
    /// Extracts all child properties from a parent object at current index
    /// and adds them as ParentName_PropertyName variables
    /// </summary>
    /// <param name="parentName">The name of the parent collection/object</param>
    /// <returns>True if properties were successfully extracted, false otherwise</returns>
    public bool ExtractNestedProperties(string parentName)
    {
        if (string.IsNullOrEmpty(parentName) || !Variables.TryGetValue(parentName, out var parentObj) || parentObj == null)
            return false;

        int currentIndex = CurrentIndices.GetValueOrDefault(parentName, 0);
        Logger.Debug($"Extracting nested properties from {parentName}[{currentIndex}]");

        // Get item at current index if parent is a collection
        object targetObj = parentObj;

        if (parentObj is IList list && currentIndex >= 0 && currentIndex < list.Count)
        {
            targetObj = list[currentIndex];
            Logger.Debug($"Using item at index {currentIndex} from list with {list.Count} items");
        }
        else if (parentObj is Array array && currentIndex >= 0 && currentIndex < array.Length)
        {
            targetObj = array.GetValue(currentIndex);
            Logger.Debug($"Using item at index {currentIndex} from array with {array.Length} items");
        }
        else
        {
            Logger.Debug($"Using {parentName} directly as it is not a collection or index is out of range");
        }

        if (targetObj == null)
        {
            Logger.Warning($"Target object is null for {parentName}[{currentIndex}]");
            return false;
        }

        // Extract all properties from the target object
        bool anyPropertiesExtracted = false;
        int propertyCount = 0;

        foreach (var prop in targetObj.GetType().GetProperties())
        {
            if (!prop.CanRead)
                continue;

            try
            {
                var value = prop.GetValue(targetObj);
                string variableName = $"{parentName}_{prop.Name}";
                Variables[variableName] = value;

                propertyCount++;
                anyPropertiesExtracted = true;

                // Also add it with indexed notation for consistency
                string indexedName = $"{parentName}[{currentIndex}].{prop.Name}";
                ProcessedNestedCollections[indexedName] = value;

                // If this property is itself a collection, extract its count as a variable too
                if (value != null)
                {
                    int count = CollectionHelper.GetCollectionCount(value);
                    if (count > 0)
                    {
                        string countName = $"{parentName}_{prop.Name}_Count";
                        Variables[countName] = count;
                        Logger.Debug($"Added collection count variable {countName} = {count}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error extracting property {prop.Name} from {parentName}: {ex.Message}");
            }
        }

        Logger.Debug($"Extracted {propertyCount} properties from {parentName}[{currentIndex}]");
        return anyPropertiesExtracted;
    }

    /// <summary>
    /// Sets the current index for a collection and extracts its nested properties
    /// </summary>
    /// <param name="collectionName">The name of the collection</param>
    /// <param name="index">The index to set</param>
    /// <param name="extractProperties">Whether to extract properties from the item at this index</param>
    /// <returns>True if index was set and properties were extracted (if requested), false otherwise</returns>
    public bool SetCollectionIndex(string collectionName, int index, bool extractProperties = true)
    {
        if (string.IsNullOrEmpty(collectionName) || !Variables.ContainsKey(collectionName))
        {
            Logger.Warning($"Collection '{collectionName}' not found in variables");
            return false;
        }

        // Check if index is in range
        var collection = Variables[collectionName];
        int count = CollectionHelper.GetCollectionCount(collection);

        if (index < 0 || index >= count)
        {
            Logger.Warning($"Index {index} is out of range for collection '{collectionName}' with {count} items");
            return false;
        }

        // Set the current index
        CurrentIndices[collectionName] = index;
        Logger.Debug($"Set collection index '{collectionName}' to {index}");

        // Extract properties if requested
        if (extractProperties)
        {
            return ExtractNestedProperties(collectionName);
        }

        return true;
    }
}