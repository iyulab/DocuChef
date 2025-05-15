using DocuChef.PowerPoint.Processing.ArrayProcessing;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Class for processing multi-level nested data structures in PowerPoint templates
/// </summary>
internal class NestedDataProcessor
{
    private readonly PowerPointProcessor _processor;
    private readonly PowerPointContext _context;

    /// <summary>
    /// Initialize nested data processor
    /// </summary>
    public NestedDataProcessor(PowerPointProcessor processor, PowerPointContext context)
    {
        _processor = processor ?? throw new ArgumentNullException(nameof(processor));
        _context = context ?? throw new ArgumentNullException(nameof(context));

        // Ensure context has CurrentIndices initialized
        if (_context.CurrentIndices == null)
        {
            _context.CurrentIndices = new Dictionary<string, int>();
        }

        // Ensure ProcessedNestedCollections is initialized
        if (_context.ProcessedNestedCollections == null)
        {
            _context.ProcessedNestedCollections = new Dictionary<string, object>();
        }
    }

    /// <summary>
    /// Process a PowerPoint document with nested data structure support
    /// </summary>
    public void Process()
    {
        var presentationPart = _processor.GetPresentationPart();
        if (presentationPart == null)
            throw new DocuChefException("Invalid presentation part");

        var slideIds = _processor.GetSlideIds(presentationPart);

        Logger.Info("Phase 1: Preparing nested data structures...");

        // 1. First pass: Analyze data and prepare nested structures
        PrepareNestedDataStructures();

        Logger.Info("Phase 2: Analyzing and preparing slides for nested data structures...");

        // 2. Slide processing
        ProcessSlides(presentationPart, slideIds);

        // Final save
        presentationPart.Presentation.Save();
        Logger.Info("PowerPoint template processing with nested data support completed successfully");
    }

    /// <summary>
    /// Process all slides, handling nested data structures
    /// </summary>
    private void ProcessSlides(PresentationPart presentationPart, IEnumerable<SlideId> slideIds)
    {
        var slideProcessor = new SlideProcessor(_processor, _context);
        var bindingProcessor = new BindingProcessor(_processor, _context);

        // Track processed slides to avoid duplication
        var processedSlideIds = new HashSet<string>();

        // Process original slides
        foreach (var slideId in slideIds.ToList())
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);
            string slidePartId = presentationPart.GetIdOfPart(slidePart);

            // Skip already processed slides
            if (processedSlideIds.Contains(slidePartId))
                continue;

            // 1. First try to process as a nested data slide
            slideProcessor.AnalyzeAndPrepareNestedData(presentationPart, slidePart);

            // 2. Then try standard array processing if not already handled
            if (!_context.ProcessedArraySlides.Contains(slidePartId))
            {
                slideProcessor.AnalyzeAndPrepareSlide(presentationPart, slidePart);
            }

            // Mark as processed
            processedSlideIds.Add(slidePartId);
        }

        // Process all slides including new ones generated during processing
        var allSlideIds = _processor.GetSlideIds(presentationPart);
        foreach (var slideId in allSlideIds)
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId);

            // Set the appropriate context for this slide
            DetermineSlideContext(slidePart);

            // Apply bindings with the determined context
            bindingProcessor.ApplyBindings(slidePart);
        }
    }

    /// <summary>
    /// Prepare nested data structures from variables
    /// </summary>
    private void PrepareNestedDataStructures()
    {
        Logger.Debug("Preparing nested data structures");

        // 1. Process explicit mappings
        ProcessExplicitMappings();

        // 2. Auto-detect and prepare nested collections
        ProcessCommonCollections();

        // 3. Generate multi-level collection indices
        GenerateMultiLevelIndices();
    }

    /// <summary>
    /// Process explicit nested collection mappings
    /// </summary>
    private void ProcessExplicitMappings()
    {
        // Look for explicit nested collection mappings
        if (_context.Variables.TryGetValue("_nestedCollections", out var mappingsObj) &&
            mappingsObj is Dictionary<string, (string ParentName, string ChildName)> mappings)
        {
            foreach (var mapping in mappings)
            {
                // Create the mapping variable
                string targetName = mapping.Key;
                var (parentName, childName) = mapping.Value;

                // Resolve and add the variable
                var value = _context.ResolveNestedData(parentName, childName);
                if (value != null)
                {
                    _context.Variables[targetName] = value;
                    Logger.Debug($"Created nested data mapping: {targetName} = {parentName}.{childName}");
                }
            }
        }
    }

    /// <summary>
    /// Process collections that may contain nested data
    /// </summary>
    private void ProcessCommonCollections()
    {
        // Look for collections in the context variables
        foreach (var entry in _context.Variables)
        {
            string collectionName = entry.Key;
            object collectionObj = entry.Value;

            // Skip non-collection variables and system variables
            if (collectionObj == null || collectionName.StartsWith("_"))
                continue;

            int count = CollectionHelper.GetCollectionCount(collectionObj);
            if (count > 0)
            {
                Logger.Debug($"Found potential parent collection: {collectionName} with {count} items");

                // Extract properties from first item for inspection
                _context.SetCollectionIndex(collectionName, 0);

                // Try to identify child collections in the first item
                var item = CollectionHelper.GetItemAtIndex(collectionObj, 0);
                if (item != null)
                {
                    IdentifyChildCollections(collectionName, item, 1);
                }
            }
        }
    }

    /// <summary>
    /// Generate multi-level indices for nested collections
    /// </summary>
    private void GenerateMultiLevelIndices()
    {
        // Look for nested collection variables (with underscore in name)
        var nestedCollections = _context.Variables.Keys
            .Where(k => k.Contains('_') && !k.StartsWith("_"))
            .ToList();

        foreach (var nestedName in nestedCollections)
        {
            var parts = nestedName.Split('_');
            if (parts.Length < 2)
                continue;

            // Get the parent collection
            string parentName = parts[0];
            if (!_context.Variables.TryGetValue(parentName, out var parentObj) || parentObj == null)
                continue;

            int parentCount = CollectionHelper.GetCollectionCount(parentObj);
            if (parentCount == 0)
                continue;

            // Process each level of nesting
            for (int i = 0; i < parentCount; i++)
            {
                // Set index for parent level
                _context.SetCollectionIndex(parentName, i);

                // Build path incrementally to set indices for each level
                string currentPath = parentName;
                object currentObj = CollectionHelper.GetItemAtIndex(parentObj, i);

                for (int level = 1; level < parts.Length; level++)
                {
                    string childName = parts[level];
                    currentPath += $"_{childName}";

                    // Get property value
                    var prop = currentObj?.GetType().GetProperty(childName);
                    if (prop == null)
                        break;

                    var childObj = prop.GetValue(currentObj);
                    if (childObj == null)
                        break;

                    int childCount = CollectionHelper.GetCollectionCount(childObj);

                    // Store the child collection
                    string fullPath = $"{parentName}[{i}].{string.Join(".", parts.Skip(1).Take(level))}";
                    _context.ProcessedNestedCollections[fullPath] = childObj;

                    // Store count info
                    _context.Variables[$"{currentPath}_Count"] = childCount;

                    // If this is not the last level and has items, continue to next level
                    if (level < parts.Length - 1 && childCount > 0)
                    {
                        // Set index for this level
                        _context.SetCollectionIndex(currentPath, 0);

                        // Process the first item
                        currentObj = CollectionHelper.GetItemAtIndex(childObj, 0);
                    }
                    else
                    {
                        break;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Recursively identify child collections within an object, with nesting level tracking
    /// </summary>
    private void IdentifyChildCollections(string parentPath, object item, int level, int maxLevel = 3)
    {
        if (item == null || level > maxLevel)
            return;

        foreach (var prop in item.GetType().GetProperties())
        {
            if (!prop.CanRead)
                continue;

            try
            {
                var value = prop.GetValue(item);
                if (value == null)
                    continue;

                int count = CollectionHelper.GetCollectionCount(value);

                // If it's a collection with items
                if (count > 0)
                {
                    // Generate path for this level
                    string targetName = $"{parentPath}_{prop.Name}";

                    Logger.Debug($"Identified nested collection (level {level}): {targetName} with {count} items");

                    // Add as a nested collection mapping
                    _context.Variables[targetName] = value;

                    // Also add the count
                    string countName = $"{targetName}_Count";
                    _context.Variables[countName] = count;

                    // Continue identifying nested collections in the first item
                    if (level < maxLevel && count > 0)
                    {
                        var firstItem = CollectionHelper.GetItemAtIndex(value, 0);
                        if (firstItem != null)
                        {
                            IdentifyChildCollections(targetName, firstItem, level + 1, maxLevel);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Error identifying child collection {prop.Name}: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Determine and set context for a specific slide
    /// </summary>
    private void DetermineSlideContext(SlidePart slidePart)
    {
        if (slidePart == null)
            return;

        var document = _context.Variables.TryGetValue("_document", out var docObj)
            ? docObj as PresentationDocument
            : null;

        string slideId = document?.PresentationPart != null
            ? document.PresentationPart.GetIdOfPart(slidePart)
            : slidePart.Uri.ToString();

        // Set slide basic context
        _context.SlidePart = slidePart;
        if (_context.Slide == null)
        {
            _context.Slide = new SlideContext();
        }
        _context.Slide.Notes = slidePart.GetNotes();

        Logger.Debug($"Determining context for slide: {slideId}");

        // Extract directives from slide notes
        var directives = DirectiveParser.ParseDirectives(_context.Slide.Notes ?? "");

        // Look for processed slides and identify context from slide ID
        DetectContextFromProcessedId(slideId);

        // Look for foreach directives that indicate needed context
        foreach (var directive in directives)
        {
            if (directive.Name.Equals("foreach", StringComparison.OrdinalIgnoreCase))
            {
                string collectionPath = directive.Value.Trim();
                SetContextFromCollectionPath(collectionPath);
            }
            else if (directive.Name.Equals("foreach-nested", StringComparison.OrdinalIgnoreCase))
            {
                string parent = directive.GetParameter("parent");
                string child = directive.GetParameter("child");

                if (!string.IsNullOrEmpty(parent) && !string.IsNullOrEmpty(child))
                {
                    SetContextFromCollectionPath($"{parent}_{child}");
                }
            }
        }
    }

    /// <summary>
    /// Detect context from a processed slide ID
    /// </summary>
    private void DetectContextFromProcessedId(string slideId)
    {
        // Look through processed slides to find matching patterns
        foreach (var processedId in _context.ProcessedArraySlides)
        {
            if (slideId.Contains(processedId))
            {
                // Try to detect multi-level collections like Parent_Child_Grandchild
                var multiLevelMatch = System.Text.RegularExpressions.Regex.Match(
                    processedId, @"^(\w+(?:_\w+)*)_(\d+)");

                if (multiLevelMatch.Success && multiLevelMatch.Groups.Count > 2)
                {
                    string path = multiLevelMatch.Groups[1].Value;
                    if (int.TryParse(multiLevelMatch.Groups[2].Value, out int index))
                    {
                        // Handle nested paths like Parent_Child_Grandchild
                        SetContextFromCollectionPath(path, index);
                        return;
                    }
                }

                // Try to detect nested collection pattern Parent_index_Child_index
                var nestedMatch = System.Text.RegularExpressions.Regex.Match(
                    processedId, @"(\w+)_(\d+)_(\w+)_(\d+)");

                if (nestedMatch.Success && nestedMatch.Groups.Count > 4)
                {
                    string parent = nestedMatch.Groups[1].Value;
                    string child = nestedMatch.Groups[3].Value;

                    if (int.TryParse(nestedMatch.Groups[2].Value, out int parentIndex) &&
                        int.TryParse(nestedMatch.Groups[4].Value, out int childIndex))
                    {
                        // Set parent index
                        _context.SetCollectionIndex(parent, parentIndex);

                        // Set combined path index
                        _context.SetCollectionIndex($"{parent}_{child}", childIndex);

                        // Also set the current batch start index
                        int maxItemsPerBatch = GetMaxItemsPerBatch($"{parent}_{child}");
                        _context.CurrentIndices[$"{child}_batch_start"] = childIndex * maxItemsPerBatch;

                        Logger.Debug($"Set nested context: {parent}[{parentIndex}].{child}[{childIndex}]");
                        return;
                    }
                }

                // Try to detect batch pattern for any collection
                var batchMatch = System.Text.RegularExpressions.Regex.Match(
                    processedId, @"batch_(\w+(?:_\w+)*)_(\d+)");

                if (batchMatch.Success && batchMatch.Groups.Count > 2)
                {
                    string collection = batchMatch.Groups[1].Value;
                    if (int.TryParse(batchMatch.Groups[2].Value, out int batchIndex))
                    {
                        int maxItemsPerBatch = GetMaxItemsPerBatch(collection);
                        int startIndex = batchIndex * maxItemsPerBatch;

                        // Handle both simple and nested collections
                        SetContextFromCollectionPath(collection, startIndex);

                        // Also store the batch start index
                        var parts = collection.Split('_');
                        if (parts.Length > 1)
                        {
                            string lastPart = parts[parts.Length - 1];
                            _context.CurrentIndices[$"{lastPart}_batch_start"] = startIndex;
                        }

                        Logger.Debug($"Set batch context: {collection} batch {batchIndex}, start={startIndex}");
                        return;
                    }
                }
            }
        }
    }

    /// <summary>
    /// Set context based on a collection path (e.g., "Parent_Child_Grandchild")
    /// </summary>
    private void SetContextFromCollectionPath(string path, int index = 0)
    {
        if (string.IsNullOrEmpty(path))
            return;

        var parts = path.Split('_');
        if (parts.Length == 0)
            return;

        // For single-level collection
        if (parts.Length == 1)
        {
            string collection = parts[0];
            if (_context.Variables.ContainsKey(collection))
            {
                _context.SetCollectionIndex(collection, index);
                _context.ExtractNestedProperties(collection);
                Logger.Debug($"Set context for {collection}[{index}]");
            }
            return;
        }

        // For multi-level nested collections
        string rootName = parts[0];
        if (!_context.Variables.TryGetValue(rootName, out var rootObj) || rootObj == null)
            return;

        // Set index for root level
        _context.SetCollectionIndex(rootName, 0);

        // Set index for combined path
        _context.SetCollectionIndex(path, index);

        // Extract properties for the full path
        _context.ExtractNestedProperties(path);

        Logger.Debug($"Set context for nested path {path}[{index}]");

        // Also ensure each intermediate collection is available
        string currentPath = rootName;
        for (int i = 1; i < parts.Length; i++)
        {
            string nextPart = parts[i];
            string nextPath = i == 1 ? $"{rootName}_{nextPart}" : $"{currentPath}_{nextPart}";

            // If this is not the full path, set index to 0
            if (i < parts.Length - 1)
            {
                _context.SetCollectionIndex(nextPath, 0);
            }

            currentPath = nextPath;
        }
    }

    /// <summary>
    /// Get maximum items per batch for a collection from directives or defaults
    /// </summary>
    private int GetMaxItemsPerBatch(string collection)
    {
        // Default value
        int defaultMax = 1;

        // Check slide notes across all processed slides
        foreach (var slideId in _context.ProcessedArraySlides)
        {
            string slidePart = slideId;

            // Look for a slide with this collection in a forEach directive
            if (slidePart.Contains(collection))
            {
                // Find the slide part
                var document = _context.Variables.TryGetValue("_document", out var docObj)
                    ? docObj as PresentationDocument
                    : null;

                if (document?.PresentationPart == null)
                    continue;

                // Fixed: Use proper method to get slide parts and their IDs
                foreach (var slidePartRel in document.PresentationPart.SlideParts)
                {
                    // Get the relationship ID for this slide part
                    string slidePartId = document.PresentationPart.GetIdOfPart(slidePartRel);

                    if (slidePartId.Contains(slidePart))
                    {
                        string notes = slidePartRel.GetNotes();
                        var directives = DirectiveParser.ParseDirectives(notes);

                        foreach (var directive in directives)
                        {
                            if (directive.Name.Equals("foreach", StringComparison.OrdinalIgnoreCase) &&
                                directive.Value.Contains(collection))
                            {
                                int max = directive.GetParameterAsInt("max", -1);
                                if (max > 0)
                                    return max;
                            }
                        }
                    }
                }
            }
        }

        // If no directive found, try to detect from arrays in context
        if (_context.Variables.TryGetValue(collection, out var collectionObj) && collectionObj != null)
        {
            // Try to auto-detect from references in any slide
            var document = _context.Variables.TryGetValue("_document", out var docObj)
                ? docObj as PresentationDocument
                : null;

            if (document?.PresentationPart != null)
            {
                int maxIndex = -1;
                var parts = collection.Split('_');
                string lastPart = parts.Length > 0 ? parts[parts.Length - 1] : collection;

                foreach (var slidePart in document.PresentationPart.SlideParts)
                {
                    foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
                    {
                        foreach (var text in shape.Descendants<A.Text>())
                        {
                            // Look for array index patterns
                            var matches = System.Text.RegularExpressions.Regex.Matches(
                                text.Text, $@"\$\{{.*{lastPart}\[(\d+)\]");

                            foreach (System.Text.RegularExpressions.Match match in matches)
                            {
                                if (match.Groups.Count > 1 && int.TryParse(match.Groups[1].Value, out int index))
                                {
                                    maxIndex = Math.Max(maxIndex, index);
                                }
                            }
                        }
                    }
                }

                if (maxIndex >= 0)
                    return maxIndex + 1;
            }
        }

        return defaultMax;
    }
}