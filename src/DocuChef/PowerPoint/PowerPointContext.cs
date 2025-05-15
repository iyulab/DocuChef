namespace DocuChef.PowerPoint;

/// <summary>
/// PowerPoint template processing context with reduced complexity
/// </summary>
public class PowerPointContext
{
    /// <summary>
    /// Variables for template processing
    /// </summary>
    public Dictionary<string, object> Variables { get; set; } = new();

    /// <summary>
    /// Global variables with dynamic evaluation
    /// </summary>
    public Dictionary<string, Func<object>> GlobalVariables { get; set; } = new();

    /// <summary>
    /// PowerPoint functions
    /// </summary>
    public Dictionary<string, PowerPointFunction> Functions { get; set; } = new();

    /// <summary>
    /// PowerPoint processing options
    /// </summary>
    public PowerPointOptions Options { get; set; }

    /// <summary>
    /// Current slide context
    /// </summary>
    public SlideContext Slide { get; set; } = new();

    /// <summary>
    /// Current shape context
    /// </summary>
    public ShapeContext Shape { get; set; } = new();

    /// <summary>
    /// Current SlidePart being processed
    /// </summary>
    public SlidePart SlidePart { get; set; }

    /// <summary>
    /// Dictionary of current indices for hierarchical paths
    /// </summary>
    public Dictionary<string, int> HierarchicalIndices { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Path navigator for resolving nested data
    /// </summary>
    public PathNavigator Navigator { get; private set; }

    /// <summary>
    /// Processed slide hierarchy maps for tracking complex slide relationships
    /// </summary>
    public Dictionary<string, SlideHierarchyMap> ProcessedHierarchies { get; } = new();

    /// <summary>
    /// Initializes the navigator for hierarchical data resolution
    /// </summary>
    public void InitializeNavigator()
    {
        Navigator = new PathNavigator(Variables);
    }

    /// <summary>
    /// Sets the current index for a hierarchical path
    /// </summary>
    public void SetHierarchicalIndex(HierarchicalPath path, int index)
    {
        if (path == null || path.Segments.Count == 0)
            return;

        // Set index for the full path
        string contextKey = PathNavigator.PathToContextKey(path);
        HierarchicalIndices[contextKey] = index;
        Logger.Debug($"Set hierarchical index for '{contextKey}' to {index}");

        // Also set indices for individual segments and partial paths
        if (path.Segments.Count > 1)
        {
            // For underscore notation (Categories_Products)
            string currentPath = "";
            for (int i = 0; i < path.Segments.Count; i++)
            {
                var segment = path.Segments[i];
                currentPath = i == 0 ? segment.Name : $"{currentPath}_{segment.Name}";

                if (i == path.Segments.Count - 1)
                {
                    HierarchicalIndices[currentPath] = index;
                    Logger.Debug($"Set underscore index for '{currentPath}' to {index}");

                    // Also set the last segment name directly as some templates might use it
                    HierarchicalIndices[segment.Name] = index;
                    Logger.Debug($"Set segment index for '{segment.Name}' to {index}");
                }
            }

            // For dot notation (Categories.Products)
            for (int i = 1; i <= path.Segments.Count; i++)
            {
                var partialPath = path.GetPathUpTo(i);
                string dotKey = partialPath.ToString();

                if (i == path.Segments.Count)
                {
                    HierarchicalIndices[dotKey] = index;
                    Logger.Debug($"Set dot notation index for '{dotKey}' to {index}");
                }
            }
        }
        else if (path.Segments.Count == 1)
        {
            // Special case for single-segment paths
            var segment = path.Segments[0];
            HierarchicalIndices[segment.Name] = index;
            Logger.Debug($"Set direct segment index for '{segment.Name}' to {index}");
        }
    }

    /// <summary>
    /// Gets the current index for a hierarchical path
    /// </summary>
    public int GetHierarchicalIndex(HierarchicalPath path)
    {
        if (path == null || path.Segments.Count == 0)
            return 0;

        string contextKey = PathNavigator.PathToContextKey(path);
        if (HierarchicalIndices.TryGetValue(contextKey, out int index))
        {
            Logger.Debug($"Found hierarchical index for '{contextKey}': {index}");
            return index;
        }

        // Try alternative formats if full path not found
        if (path.Segments.Count > 1)
        {
            // Try underscore format
            string underscoreKey = path.ToUnderscoreFormat();
            if (HierarchicalIndices.TryGetValue(underscoreKey, out index))
            {
                Logger.Debug($"Found underscore index for '{underscoreKey}': {index}");
                return index;
            }

            // Try dot format
            string dotKey = path.ToString();
            if (HierarchicalIndices.TryGetValue(dotKey, out index))
            {
                Logger.Debug($"Found dot notation index for '{dotKey}': {index}");
                return index;
            }

            // Try last segment name directly
            string lastSegmentName = path.Segments.Last().Name;
            if (HierarchicalIndices.TryGetValue(lastSegmentName, out index))
            {
                Logger.Debug($"Found direct segment index for '{lastSegmentName}': {index}");
                return index;
            }
        }

        Logger.Debug($"No index found for path '{path}', using default 0");
        return 0;
    }

    /// <summary>
    /// Resolves a value at a hierarchical path using current context indices
    /// </summary>
    public object ResolveHierarchicalValue(HierarchicalPath path)
    {
        if (Navigator == null)
            InitializeNavigator();

        return Navigator.ResolveValueWithContext(path, HierarchicalIndices);
    }

    /// <summary>
    /// Resolves a value at a hierarchical path specified as a string
    /// </summary>
    public object ResolveHierarchicalValue(string pathString)
    {
        if (string.IsNullOrEmpty(pathString))
            return null;

        return ResolveHierarchicalValue(new HierarchicalPath(pathString));
    }

    /// <summary>
    /// Gets the count of items in a collection at the specified path
    /// </summary>
    public int GetCollectionCount(HierarchicalPath path)
    {
        if (Navigator == null)
            InitializeNavigator();

        return Navigator.GetCollectionCount(path);
    }

    /// <summary>
    /// Gets a hierarchy map or creates a new one if it doesn't exist
    /// </summary>
    public SlideHierarchyMap GetOrCreateHierarchyMap(string mapId)
    {
        if (!ProcessedHierarchies.TryGetValue(mapId, out var map))
        {
            map = new SlideHierarchyMap { Id = mapId };
            ProcessedHierarchies[mapId] = map;
        }
        return map;
    }

    /// <summary>
    /// Displays current hierarchical indices for debugging
    /// </summary>
    public void DumpHierarchicalIndices()
    {
        Logger.Debug("Current hierarchical indices:");
        foreach (var kvp in HierarchicalIndices.OrderBy(x => x.Key))
        {
            Logger.Debug($"  {kvp.Key} = {kvp.Value}");
        }
    }
}

/// <summary>
/// Context for slide processing
/// </summary>
public class SlideContext
{
    /// <summary>
    /// Slide ID
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// Slide notes
    /// </summary>
    public string Notes { get; set; }
}

/// <summary>
/// Context for shape processing
/// </summary>
public class ShapeContext
{
    /// <summary>
    /// Shape name
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Shape type
    /// </summary>
    public string Type { get; set; }

    /// <summary>
    /// Shape ID
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// Shape text
    /// </summary>
    public string Text { get; set; }

    /// <summary>
    /// The actual shape object
    /// </summary>
    public Shape ShapeObject { get; set; }
}

/// <summary>
/// Represents a mapping between slides and hierarchical data paths
/// </summary>
public class SlideHierarchyMap
{
    /// <summary>
    /// Unique identifier for the map
    /// </summary>
    public string Id { get; set; }

    /// <summary>
    /// The hierarchical path being processed
    /// </summary>
    public HierarchicalPath Path { get; set; }

    /// <summary>
    /// Mappings between slide IDs and indices
    /// </summary>
    public Dictionary<string, SlideIndexMapping> SlideMappings { get; } = new();

    /// <summary>
    /// Gets a slide mapping or creates a new one if it doesn't exist
    /// </summary>
    public SlideIndexMapping GetOrCreateSlideMapping(string slideId)
    {
        if (!SlideMappings.TryGetValue(slideId, out var mapping))
        {
            mapping = new SlideIndexMapping { SlideId = slideId };
            SlideMappings[slideId] = mapping;
        }
        return mapping;
    }
}

/// <summary>
/// Maps a slide to specific indices in a hierarchical path
/// </summary>
public class SlideIndexMapping
{
    /// <summary>
    /// Slide identifier
    /// </summary>
    public string SlideId { get; set; }

    /// <summary>
    /// Index mappings for each level of the path
    /// </summary>
    public Dictionary<string, int> PathIndices { get; } = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Base index for the current segment being iterated
    /// </summary>
    public int BaseIndex { get; set; }

    /// <summary>
    /// Items per slide for batch processing
    /// </summary>
    public int ItemsPerSlide { get; set; } = 1;

    /// <summary>
    /// Actual number of items on this slide
    /// </summary>
    public int TotalItems { get; set; } = 0;
}