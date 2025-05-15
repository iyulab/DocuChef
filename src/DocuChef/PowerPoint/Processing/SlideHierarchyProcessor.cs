using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Processes hierarchical foreach directives by handling slide duplication and data mapping
/// </summary>
internal class SlideHierarchyProcessor
{
    private readonly PowerPointContext _context;
    private readonly Dictionary<string, object> _variables;

    /// <summary>
    /// Creates a new instance of SlideHierarchyProcessor
    /// </summary>
    public SlideHierarchyProcessor(PowerPointContext context, Dictionary<string, object> variables)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _variables = variables ?? throw new ArgumentNullException(nameof(variables));

        // Ensure Navigator is initialized
        if (_context.Navigator == null)
            _context.InitializeNavigator();
    }

    /// <summary>
    /// Process hierarchical foreach directive
    /// </summary>
    public SlideProcessingResult ProcessHierarchicalForeach(
        PresentationPart presentationPart,
        SlidePart slidePart,
        Directive directive)
    {
        var result = new SlideProcessingResult
        {
            SlidePart = slidePart,
            WasProcessed = false
        };

        if (directive == null || directive.Path == null || directive.Path.Segments.Count == 0)
        {
            Logger.Warning("Invalid hierarchical foreach directive");
            return result;
        }

        // Get hierarchical path from directive
        string pathString = directive.Value.Trim();
        HierarchicalPath path = directive.Path;

        Logger.Debug($"Processing hierarchical foreach directive for path: {pathString}");

        // Extract parameters
        int maxItemsPerSlide = directive.GetParameterAsInt("max", 1);
        int offset = directive.GetParameterAsInt("offset", 0);
        bool skipEmpty = directive.GetParameterAsBool("skipEmpty", true);

        // Check if the path exists and resolve to a collection
        if (!VerifyPathExists(path, skipEmpty))
        {
            Logger.Warning($"Path does not exist or is empty: {path}");
            result.WasProcessed = true; // Mark as processed even if invalid
            return result;
        }

        // Generate a unique map ID for this hierarchy
        string mapId = GenerateHierarchyMapId(slidePart, path);
        var hierarchyMap = _context.GetOrCreateHierarchyMap(mapId);
        hierarchyMap.Path = path;

        // Process the hierarchy levels
        return ProcessHierarchyLevels(
            presentationPart,
            slidePart,
            path,
            maxItemsPerSlide,
            offset,
            hierarchyMap);
    }

    /// <summary>
    /// Verify that a hierarchical path exists and references a collection
    /// </summary>
    private bool VerifyPathExists(HierarchicalPath path, bool skipEmpty)
    {
        if (path == null || path.Segments.Count == 0)
            return false;

        // Check if root exists
        var rootSegment = path.GetRoot();
        if (rootSegment == null || !_variables.TryGetValue(rootSegment.Name, out var rootObj) || rootObj == null)
        {
            Logger.Warning($"Root collection not found: {rootSegment?.Name ?? "null"}");
            return false;
        }

        // If single segment path, check if it's a collection
        if (path.Segments.Count == 1)
        {
            int count = DataNavigationHelper.GetCollectionCount(rootObj);
            if (count == 0 && skipEmpty)
            {
                Logger.Warning($"Root collection is empty: {rootSegment.Name}");
                return false;
            }
            return true;
        }

        // Navigate through the path to verify it exists
        object currentObj = rootObj;
        for (int i = 1; i < path.Segments.Count; i++)
        {
            var pathToCheck = path.GetPathUpTo(i + 1);
            var resolved = _context.Navigator.ResolveValue(pathToCheck);

            if (resolved == null)
            {
                Logger.Warning($"Path segment not found: {pathToCheck}");
                return false;
            }

            // If at last segment, check if it's a collection
            if (i == path.Segments.Count - 1)
            {
                int count = DataNavigationHelper.GetCollectionCount(resolved);
                if (count == 0 && skipEmpty)
                {
                    Logger.Warning($"Collection is empty at path: {pathToCheck}");
                    return false;
                }
            }

            currentObj = resolved;
        }

        return true;
    }

    /// <summary>
    /// Process all levels of a hierarchical path
    /// </summary>
    private SlideProcessingResult ProcessHierarchyLevels(
        PresentationPart presentationPart,
        SlidePart slidePart,
        HierarchicalPath path,
        int maxItemsPerSlide,
        int offset,
        SlideHierarchyMap hierarchyMap)
    {
        var result = new SlideProcessingResult
        {
            SlidePart = slidePart,
            WasProcessed = true
        };

        // Ensure path refers to a collection
        var collectionPath = SimplifyPathToCollection(path);

        Logger.Debug($"Using simplified collection path: {collectionPath}");

        // Resolve the collection
        var collection = _context.Navigator.ResolveValue(collectionPath);
        if (collection == null)
        {
            Logger.Warning($"Could not resolve collection at path: {collectionPath}");
            return result;
        }

        // Get collection count
        int totalItems = DataNavigationHelper.GetCollectionCount(collection);
        if (totalItems == 0)
        {
            Logger.Debug($"Collection at path {collectionPath} is empty");
            return result;
        }

        Logger.Debug($"Processing collection with {totalItems} items, max {maxItemsPerSlide} per slide");

        // Ensure maxItemsPerSlide is valid
        if (maxItemsPerSlide <= 0)
        {
            maxItemsPerSlide = 5; // Default to 5 if invalid
            Logger.Debug($"Invalid maxItemsPerSlide, using default: {maxItemsPerSlide}");
        }

        // Calculate needed slides
        int slidesNeeded = (int)Math.Ceiling((double)(totalItems - offset) / maxItemsPerSlide);
        if (slidesNeeded <= 0)
        {
            Logger.Debug($"No slides needed for collection (offset: {offset}, items: {totalItems})");
            return result;
        }

        // Create hierarchy index context
        var hierarchyIndices = new Dictionary<string, int>(_context.HierarchicalIndices);

        // Store the original slide position
        int slidePosition = SlideHelper.FindSlidePosition(presentationPart, slidePart);

        // Process each batch of items
        Logger.Debug($"Creating {slidesNeeded} slides for collection at path {collectionPath}");

        for (int slideIndex = 0; slideIndex < slidesNeeded; slideIndex++)
        {
            // Calculate start and end indices for this slide
            int startIdx = offset + (slideIndex * maxItemsPerSlide);
            int endIdx = Math.Min(startIdx + maxItemsPerSlide - 1, totalItems - 1);
            int itemsOnSlide = endIdx - startIdx + 1;

            Logger.Debug($"Processing slide {slideIndex + 1} with items {startIdx} to {endIdx} ({itemsOnSlide} items)");

            // Create or use slide
            SlidePart currentSlide;
            if (slideIndex == 0)
            {
                // Use the original slide for the first batch
                currentSlide = slidePart;
                Logger.Debug("Using original slide for first batch");
            }
            else
            {
                // Clone the original slide for subsequent batches
                currentSlide = SlideHelper.CloneSlide(presentationPart, slidePart);

                // Insert after the previous slide
                int insertPosition = slidePosition + slideIndex;
                SlideHelper.InsertSlide(presentationPart, currentSlide, insertPosition);

                Logger.Debug($"Created new slide at position {insertPosition}");
                result.GeneratedSlides.Add(currentSlide);
            }

            // Store slide mapping information
            string slideId = presentationPart.GetIdOfPart(currentSlide);
            var slideMapping = hierarchyMap.GetOrCreateSlideMapping(slideId);
            slideMapping.BaseIndex = startIdx;
            slideMapping.ItemsPerSlide = maxItemsPerSlide;
            slideMapping.TotalItems = itemsOnSlide;

            // Set hierarchical indices for this slide
            UpdateHierarchicalIndicesForSlide(collectionPath, startIdx, hierarchyIndices, slideMapping);
        }

        // Restore original indices
        _context.HierarchicalIndices.Clear();
        foreach (var kvp in hierarchyIndices)
        {
            _context.HierarchicalIndices[kvp.Key] = kvp.Value;
        }

        return result;
    }

    /// <summary>
    /// Simplify a path to reference just the collection
    /// </summary>
    private HierarchicalPath SimplifyPathToCollection(HierarchicalPath path)
    {
        // If this is already a simple path, return it
        if (path.Segments.Count == 1 && !path.Segments[0].Index.HasValue)
            return path;

        // Create a simplified path with just the root collection
        var rootSegment = path.GetRoot();
        if (rootSegment == null)
            return path;

        var simplified = new HierarchicalPath();
        simplified.AddSegment(rootSegment.Name);
        return simplified;
    }

    /// <summary>
    /// Update hierarchical indices for a slide
    /// </summary>
    private void UpdateHierarchicalIndicesForSlide(
        HierarchicalPath path,
        int baseIndex,
        Dictionary<string, int> originalIndices,
        SlideIndexMapping slideMapping)
    {
        // Store the full path index
        string rootKey = PathNavigator.PathToContextKey(path);
        slideMapping.PathIndices[rootKey] = baseIndex;

        // For combined path navigation
        string currentPath = string.Empty;

        // Process each segment of the path
        for (int i = 0; i < path.Segments.Count; i++)
        {
            var segment = path.Segments[i];

            // Build segment key for both dot notation and underscore notation
            currentPath = i == 0 ? segment.Name : $"{currentPath}_{segment.Name}";
            string dotPath = string.Join(".", path.Segments.Take(i + 1).Select(s => s.Name));

            // Assign index from base index if this is the main collection
            if (i == path.Segments.Count - 1)
            {
                slideMapping.PathIndices[currentPath] = baseIndex;
                slideMapping.PathIndices[dotPath] = baseIndex;

                // Also store individual segment name if it's the last one
                slideMapping.PathIndices[segment.Name] = baseIndex;
            }
            // Otherwise use existing index if available
            else if (originalIndices.TryGetValue(currentPath, out int existingIndex))
            {
                slideMapping.PathIndices[currentPath] = existingIndex;
                slideMapping.PathIndices[dotPath] = existingIndex;
            }

            Logger.Debug($"Set index for path '{currentPath}' to {slideMapping.PathIndices[currentPath]}");
        }
    }

    /// <summary>
    /// Generate a unique ID for a hierarchy map
    /// </summary>
    private string GenerateHierarchyMapId(SlidePart slidePart, HierarchicalPath path)
    {
        string slideId = slidePart != null ?
            _context.Variables.ContainsKey("_document") ?
                ((PresentationDocument)_context.Variables["_document"]).PresentationPart.GetIdOfPart(slidePart) :
                "unknown" :
            "unknown";

        return $"hierarchy_{path.ToUnderscoreFormat()}_{slideId}";
    }
}