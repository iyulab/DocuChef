namespace DocuChef.PowerPoint;

/// <summary>
/// Handles navigation of hierarchical data paths with optimized caching
/// </summary>
public class PathNavigator
{
    private readonly Dictionary<string, object> _variables;
    private readonly Dictionary<string, object> _cache = new();
    private readonly int _maxCacheSize = 200;

    /// <summary>
    /// Creates a new path navigator with the specified variables
    /// </summary>
    public PathNavigator(Dictionary<string, object> variables)
    {
        _variables = variables ?? throw new ArgumentNullException(nameof(variables));
    }

    /// <summary>
    /// Resolves the value at the specified path
    /// </summary>
    public object ResolveValue(HierarchicalPath path)
    {
        if (path == null || path.Segments.Count == 0)
            return null;

        string cacheKey = path.ToString();

        // Check cache first
        if (_cache.TryGetValue(cacheKey, out var cachedValue))
        {
            Logger.Debug($"Cache hit for path: {cacheKey}");
            return cachedValue;
        }

        Logger.Debug($"Resolving path: {cacheKey}");

        // Start with the root segment
        var rootSegment = path.GetRoot();
        if (!_variables.TryGetValue(rootSegment.Name, out var currentObj) || currentObj == null)
        {
            Logger.Debug($"Root not found: {rootSegment.Name}");
            return null;
        }

        // Apply index if specified for root
        if (rootSegment.Index.HasValue)
        {
            currentObj = GetItemAtIndex(currentObj, rootSegment.Index.Value);
            if (currentObj == null)
            {
                Logger.Debug($"Index {rootSegment.Index.Value} not found in root {rootSegment.Name}");
                return null;
            }
        }

        // Navigate through remaining segments
        for (int i = 1; i < path.Segments.Count; i++)
        {
            var segment = path.Segments[i];

            if (currentObj == null)
            {
                Logger.Debug($"Null object before processing segment {i}: {segment.Name}");
                return null;
            }

            // Handle property navigation
            currentObj = DataNavigationHelper.GetPropertyValue(currentObj, segment.Name);

            if (currentObj == null)
            {
                Logger.Debug($"Property not found: {segment.Name}");
                return null;
            }

            // Apply index if specified
            if (segment.Index.HasValue)
            {
                currentObj = GetItemAtIndex(currentObj, segment.Index.Value);
                if (currentObj == null)
                {
                    Logger.Debug($"Index {segment.Index.Value} not found in {segment.Name}");
                    return null;
                }
            }
        }

        // Store in cache (with potential cleanup)
        if (_cache.Count >= _maxCacheSize)
        {
            CleanupCache();
        }

        _cache[cacheKey] = currentObj;
        return currentObj;
    }

    /// <summary>
    /// Resolves a value using dynamic indices from a context
    /// </summary>
    public object ResolveValueWithContext(HierarchicalPath path, Dictionary<string, int> indices)
    {
        if (path == null || path.Segments.Count == 0)
            return null;

        // Log the indices for debugging
        Logger.Debug($"Resolving path with context: {path}");

        if (Logger.MinimumLevel <= Logger.LogLevel.Debug && indices.Count > 0)
        {
            Logger.Debug("Context indices:");
            foreach (var kvp in indices.OrderBy(x => x.Key))
            {
                Logger.Debug($"  {kvp.Key} = {kvp.Value}");
            }
        }

        // Check for direct underscore path match in indices
        string underscoreKey = path.ToUnderscoreFormat();
        bool hasContextIndex = false;
        int pathIndex = 0;

        // Look for index for the complete path
        if (indices.TryGetValue(underscoreKey, out pathIndex))
        {
            // If we have an index for the complete path, adjust the last segment
            Logger.Debug($"Found context index for complete path '{underscoreKey}': {pathIndex}");
            hasContextIndex = true;
        }

        // Create a copy of the path with indices applied from context
        var resolvedPath = ApplyContextIndices(path, indices, hasContextIndex, pathIndex);

        Logger.Debug($"Resolved path with indices: {resolvedPath}");
        return ResolveValue(resolvedPath);
    }

    /// <summary>
    /// Resolves a path specified as a string
    /// </summary>
    public object ResolveValue(string pathString)
    {
        if (string.IsNullOrEmpty(pathString))
            return null;

        return ResolveValue(new HierarchicalPath(pathString));
    }

    /// <summary>
    /// Resolves a path string using dynamic indices from a context
    /// </summary>
    public object ResolveValueWithContext(string pathString, Dictionary<string, int> indices)
    {
        if (string.IsNullOrEmpty(pathString))
            return null;

        return ResolveValueWithContext(new HierarchicalPath(pathString), indices);
    }

    /// <summary>
    /// Gets the count of items in a collection at the specified path
    /// </summary>
    public int GetCollectionCount(HierarchicalPath path)
    {
        var collection = ResolveValue(path);
        if (collection == null)
            return 0;

        return DataNavigationHelper.GetCollectionCount(collection);
    }

    /// <summary>
    /// Gets all items in a collection at the specified path
    /// </summary>
    public List<object> GetCollectionItems(HierarchicalPath path)
    {
        var collection = ResolveValue(path);
        if (collection == null)
            return new List<object>();

        return DataNavigationHelper.GetCollectionItems(collection);
    }

    /// <summary>
    /// Gets an item at a specific index from a collection at the specified path
    /// </summary>
    public object GetItemAtIndex(HierarchicalPath path, int index)
    {
        var collection = ResolveValue(path);
        if (collection == null)
            return null;

        return GetItemAtIndex(collection, index);
    }

    /// <summary>
    /// Converts a path to a context key for indices
    /// </summary>
    public static string PathToContextKey(HierarchicalPath path)
    {
        return path?.ToUnderscoreFormat() ?? string.Empty;
    }

    /// <summary>
    /// Applies context indices to a path
    /// </summary>
    private HierarchicalPath ApplyContextIndices(
        HierarchicalPath path,
        Dictionary<string, int> indices,
        bool hasContextIndex = false,
        int contextIndex = 0)
    {
        var result = new HierarchicalPath();
        var currentPath = new HierarchicalPath();

        for (int i = 0; i < path.Segments.Count; i++)
        {
            // Copy the segment with its existing index
            var segment = path.Segments[i];
            var newSegment = new PathSegment(segment.Name);

            // Set index from explicit index in the path
            if (segment.Index.HasValue)
            {
                newSegment.Index = segment.Index.Value;
                Logger.Debug($"Using explicit index for segment '{segment.Name}': {segment.Index.Value}");
            }
            // Use context index if this is the last segment and we have a context index for the full path
            else if (i == path.Segments.Count - 1 && hasContextIndex)
            {
                newSegment.Index = contextIndex;
                Logger.Debug($"Using context index for last segment '{segment.Name}': {contextIndex}");
            }
            // Otherwise check for indices in the context
            else
            {
                // Add this segment to the current path
                currentPath.AddSegment(segment.Name);

                // Look for indices using progressively more specific paths
                string contextKey = PathToContextKey(currentPath);
                if (indices.TryGetValue(contextKey, out int index))
                {
                    newSegment.Index = index;
                    Logger.Debug($"Found index for partial path '{contextKey}': {index}");
                }
                // Also check direct segment name
                else if (indices.TryGetValue(segment.Name, out index))
                {
                    newSegment.Index = index;
                    Logger.Debug($"Found index for segment name '{segment.Name}': {index}");
                }
            }

            result.Segments.Add(newSegment);
        }

        return result;
    }

    /// <summary>
    /// Gets an item at the specified index from a collection
    /// </summary>
    private object GetItemAtIndex(object collection, int index)
    {
        return DataNavigationHelper.GetItemAtIndex(collection, index);
    }

    /// <summary>
    /// Cleans up the cache when it gets too large
    /// </summary>
    private void CleanupCache()
    {
        // Simple strategy: Remove half of the oldest entries
        int removeCount = _cache.Count / 2;
        var keysToRemove = _cache.Keys.Take(removeCount).ToList();

        foreach (var key in keysToRemove)
        {
            _cache.Remove(key);
        }
    }

    /// <summary>
    /// Clears the path resolution cache
    /// </summary>
    public void ClearCache()
    {
        _cache.Clear();
    }
}