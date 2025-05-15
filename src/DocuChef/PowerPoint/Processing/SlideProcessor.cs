using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Processor coordinator for PowerPoint template processing
/// </summary>
internal class SlideProcessor
{
    private readonly PowerPointContext _context;
    private readonly IExpressionEvaluator _evaluator;
    private readonly ShapeProcessor _shapeProcessor;

    /// <summary>
    /// Initialize slide processor
    /// </summary>
    public SlideProcessor(IExpressionEvaluator evaluator, PowerPointContext context)
    {
        _evaluator = evaluator ?? throw new ArgumentNullException(nameof(evaluator));
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _shapeProcessor = new ShapeProcessor(evaluator, context);
    }

    /// <summary>
    /// Analyze slide and prepare duplicates if needed
    /// </summary>
    public void AnalyzeAndPrepareSlide(PresentationPart presentationPart, SlidePart slidePart, Dictionary<string, object> variables)
    {
        string slideId = presentationPart.GetIdOfPart(slidePart);
        Logger.Debug($"Analyzing slide {slideId} for hierarchical data and directives");

        // First check for slide directives in notes
        var directiveProcessor = new DirectiveProcessor(_context, variables);
        var dirResult = directiveProcessor.ProcessDirectives(presentationPart, slidePart);

        if (dirResult.WasProcessed)
        {
            // Slide was already processed by directives, no need for further processing
            Logger.Debug($"Slide {slideId} processed by directives");
            return;
        }

        // If no directives found, use automatic hierarchical reference detection
        AutoDetectAndProcessHierarchicalReferences(presentationPart, slidePart, variables);
    }

    /// <summary>
    /// Auto-detect hierarchical references and process the slide accordingly
    /// </summary>
    private void AutoDetectAndProcessHierarchicalReferences(
        PresentationPart presentationPart,
        SlidePart slidePart,
        Dictionary<string, object> variables)
    {
        string slideId = presentationPart.GetIdOfPart(slidePart);
        Logger.Debug($"Using automatic hierarchical reference detection for slide {slideId}");

        // Find all hierarchical references in the slide
        var hierarchicalPaths = FindHierarchicalReferencesInSlide(slidePart);
        if (!hierarchicalPaths.Any())
        {
            Logger.Debug("No hierarchical references found in slide");
            return;
        }

        // Group by root collection name
        var groupedPaths = hierarchicalPaths
            .GroupBy(p => p.GetRoot()?.Name)
            .Where(g => g.Key != null);

        Logger.Debug($"Found {groupedPaths.Count()} root collections referenced in slide");

        // Process each group
        foreach (var group in groupedPaths)
        {
            string rootName = group.Key;
            Logger.Debug($"Processing references for root collection '{rootName}' in slide");

            // Check if the root is actually a collection
            if (!variables.TryGetValue(rootName, out var rootCollection) || rootCollection == null)
            {
                Logger.Warning($"Root collection '{rootName}' not found in variables");
                continue;
            }

            // Verify it's a collection (not a single item with properties)
            int collectionSize = DataNavigationHelper.GetCollectionCount(rootCollection);
            if (collectionSize <= 0)
            {
                Logger.Warning($"Root object '{rootName}' is not a collection or is empty");
                continue;
            }

            Logger.Debug($"Collection '{rootName}' contains {collectionSize} items");

            // Find the best candidate path - prefer direct array references first
            var candidatePaths = group
                .Where(p => p.Segments.Count > 0 && p.Segments[0].Name.Equals(rootName, StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Sort by preference: paths with array indices first, then by depth (deeper paths first)
            var sortedPaths = candidatePaths
                .OrderByDescending(p => p.Segments.Any(s => s.Index.HasValue)) // Array indices first
                .ThenByDescending(p => p.Segments.Count) // Deeper paths next
                .ToList();

            if (sortedPaths.Any())
            {
                var deepestPath = sortedPaths.First();

                // Create a simplified path that just references the root collection
                var simplifiedPath = new HierarchicalPath();
                simplifiedPath.AddSegment(rootName);

                Logger.Debug($"Using collection path: {rootName} (simplified from {deepestPath})");

                // Create a directive for auto-processing
                var autoDirective = new Directive
                {
                    Name = "foreach",
                    Value = simplifiedPath.ToString(),
                    Path = simplifiedPath
                };

                // Determine max items per slide from template design
                int maxItemsPerSlide = DetermineMaxItemsPerSlide(slidePart, deepestPath);

                // Validate detected items per slide against collection references
                int explicitIndices = FindExplicitArrayIndices(slidePart, rootName).Count;
                if (explicitIndices > 0)
                {
                    // If we found explicit indices, use the highest one + 1
                    maxItemsPerSlide = explicitIndices;
                    Logger.Debug($"Using {maxItemsPerSlide} items per slide based on explicit indices");
                }
                else if (maxItemsPerSlide <= 0)
                {
                    // Default to 5 if no pattern detected
                    maxItemsPerSlide = 5;
                    Logger.Debug($"No clear pattern detected, defaulting to {maxItemsPerSlide} items per slide");
                }

                autoDirective.Parameters["max"] = maxItemsPerSlide.ToString();
                Logger.Debug($"Set max items per slide: {maxItemsPerSlide}");

                // Process using the hierarchical processor
                var processor = new SlideHierarchyProcessor(_context, variables);
                var result = processor.ProcessHierarchicalForeach(presentationPart, slidePart, autoDirective);

                if (result.WasProcessed)
                {
                    Logger.Debug($"Successfully auto-processed hierarchical references for '{deepestPath}'");
                    Logger.Debug($"Generated {result.GeneratedSlides.Count} additional slides");
                    break; // Only process the first group successfully
                }
                else
                {
                    Logger.Warning($"Failed to process collection '{rootName}' with directive");
                }
            }
        }
    }

    /// <summary>
    /// Determine maximum items per slide based on array index references
    /// </summary>
    private int DetermineMaxItemsPerSlide(SlidePart slidePart, HierarchicalPath path)
    {
        Logger.Debug($"Determining max items per slide for path: {path}");
        string rootName = path.GetRoot()?.Name ?? "";

        // First, try to find explicit array indices in the slide
        var indices = FindExplicitArrayIndices(slidePart, rootName);

        if (indices.Count > 0)
        {
            int maxIndex = indices.Max();
            int itemsPerSlide = maxIndex + 1; // Count from 0
            Logger.Debug($"Found explicit array indices: {string.Join(", ", indices.OrderBy(i => i))}");
            Logger.Debug($"Highest index: {maxIndex}, items per slide: {itemsPerSlide}");
            return itemsPerSlide;
        }

        // If no explicit indices found, try to determine from shape patterns
        int itemsFromPatterns = DetectItemsFromShapePatterns(slidePart);
        if (itemsFromPatterns > 1)
        {
            Logger.Debug($"Detected {itemsFromPatterns} items per slide from shape patterns");
            return itemsFromPatterns;
        }

        // Default to a reasonable number if no pattern detected
        // For most templates, 5 is common for lists or catalog items
        int defaultItemsPerSlide = 5;
        Logger.Debug($"No clear pattern detected, using default of {defaultItemsPerSlide} items per slide");
        return defaultItemsPerSlide;
    }

    /// <summary>
    /// Find explicit array indices for a collection in a slide
    /// </summary>
    private HashSet<int> FindExplicitArrayIndices(SlidePart slidePart, string collectionName)
    {
        var indices = new HashSet<int>();

        if (string.IsNullOrEmpty(collectionName) || slidePart?.Slide == null)
            return indices;

        // Pattern to match array indices with the collection name
        // Example: Items[0], Items[1], etc.
        var indexPattern = new System.Text.RegularExpressions.Regex(
            $@"\$\{{{collectionName}\[(\d+)\]",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        );

        // Check all text elements in shapes
        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null)
                continue;

            foreach (var text in shape.Descendants<A.Text>())
            {
                if (string.IsNullOrEmpty(text.Text))
                    continue;

                var matches = indexPattern.Matches(text.Text);
                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    if (match.Groups.Count >= 2 && int.TryParse(match.Groups[1].Value, out int index))
                    {
                        indices.Add(index);
                    }
                }
            }
        }

        return indices;
    }

    /// <summary>
    /// Detect number of items per slide from recurring shape patterns
    /// </summary>
    private int DetectItemsFromShapePatterns(SlidePart slidePart)
    {
        // This is a heuristic to detect repeating patterns of shapes that might
        // represent items in a collection (like product items in a catalog)

        // Get all shapes with potential naming patterns
        var shapes = slidePart.Slide.Descendants<P.Shape>()
            .Where(s => s.GetShapeName() != null)
            .ToList();

        // Look for numeric suffixes in shape names (e.g. "Image1", "Image2", "Image3")
        var numericSuffixGroups = shapes
            .Select(s => System.Text.RegularExpressions.Regex.Match(s.GetShapeName(), @"^(.+?)(\d+)$"))
            .Where(m => m.Success)
            .GroupBy(m => m.Groups[1].Value)
            .Where(g => g.Count() > 1)
            .ToDictionary(
                g => g.Key,
                g => g.Select(m => int.Parse(m.Groups[2].Value)).OrderBy(i => i).ToList()
            );

        // If we found groups with numeric suffixes, find the largest group that has sequential numbers
        if (numericSuffixGroups.Any())
        {
            var largestGroup = numericSuffixGroups
                .OrderByDescending(g => g.Value.Count)
                .FirstOrDefault();

            if (largestGroup.Value != null && largestGroup.Value.Count > 1)
            {
                // Check if numbers are sequential starting from a common base
                var numbers = largestGroup.Value;
                int baseNumber = numbers.Min();
                bool isSequential = true;

                for (int i = 0; i < numbers.Count; i++)
                {
                    if (numbers[i] != baseNumber + i)
                    {
                        isSequential = false;
                        break;
                    }
                }

                if (isSequential)
                {
                    Logger.Debug($"Found sequential shape pattern: {largestGroup.Key} with {numbers.Count} items");
                    return numbers.Count;
                }
            }
        }

        // Fallback: Look for patterns in vertical positioning of shapes
        // This assumes items are arranged vertically with similar spacing

        var verticalGroups = shapes
            .GroupBy(s => s.ShapeProperties?.Transform2D?.Offset?.Y?.Value)
            .Where(g => g.Key.HasValue && g.Count() > 1)
            .OrderByDescending(g => g.Count())
            .ToList();

        if (verticalGroups.Count >= 2)
        {
            // If we have at least 2 rows of shapes with similar counts, this might be a list
            var topRows = verticalGroups.Take(3).ToList(); // Consider up to 3 rows
            int maxCount = topRows.Max(g => g.Count());
            int avgCount = (int)Math.Round(topRows.Average(g => g.Count()));

            // If most rows have similar shape counts, this is likely a pattern
            if (avgCount >= 2 && avgCount <= 5 && Math.Abs(maxCount - avgCount) <= 1)
            {
                Logger.Debug($"Detected tabular pattern with approximately {avgCount} items per row");
                return avgCount;
            }
        }

        // Default: no pattern detected
        return 1;
    }

    /// <summary>
    /// Find all hierarchical references in a slide
    /// </summary>
    private List<HierarchicalPath> FindHierarchicalReferencesInSlide(SlidePart slidePart)
    {
        var result = new List<HierarchicalPath>();

        if (slidePart?.Slide == null)
            return result;

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null)
                continue;

            // Search all text elements for expressions
            foreach (var text in shape.Descendants<A.Text>())
            {
                if (string.IsNullOrEmpty(text.Text))
                    continue;

                // Look for ${...} expressions
                var matches = System.Text.RegularExpressions.Regex.Matches(
                    text.Text, @"\${([^{}]+)}");

                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    string expression = match.Groups[1].Value;

                    // Skip PowerPoint function expressions
                    if (expression.StartsWith("ppt."))
                        continue;

                    // Check if expression contains hierarchical references
                    if (expression.Contains('.') || expression.Contains('[') ||
                        (expression.Contains('_') && !expression.StartsWith('_')))
                    {
                        try
                        {
                            // Parse and add the path
                            var path = new HierarchicalPath(expression);
                            if (path.Segments.Count > 0)
                            {
                                result.Add(path);
                                Logger.Debug($"Found hierarchical reference: {expression}");
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Warning($"Error parsing hierarchical reference '{expression}': {ex.Message}");
                        }
                    }
                }
            }
        }

        return result;
    }

    /// <summary>
    /// Apply bindings to all shapes in a slide
    /// </summary>
    public void ApplyBindings(SlidePart slidePart, Dictionary<string, object> variables)
    {
        // Get the presentation document from context
        var presentationDoc = _context.Variables.ContainsKey("_document")
            ? _context.Variables["_document"] as PresentationDocument
            : null;

        if (presentationDoc == null)
        {
            Logger.Warning("Presentation document not found in context");
            return;
        }

        string slideId = presentationDoc.PresentationPart.GetIdOfPart(slidePart);
        Logger.Debug($"Applying bindings to slide {slideId}");

        // Set slide context
        _context.SlidePart = slidePart;
        _context.Slide = new SlideContext
        {
            Id = slideId,
            Notes = slidePart.GetNotes()
        };

        // Check for hierarchy mapping for this slide
        ApplyHierarchicalIndicesForSlide(slideId);

        // First pass: pre-scan all shapes to identify and hide those with invalid array references
        // This helps prevent processing shapes that will be hidden anyway
        PreScanAndHideInvalidShapes(slidePart, variables);

        // Second pass: process all remaining shapes
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();

        Logger.Debug($"Processing {shapes.Count} shapes in slide {slideId}");

        int processedShapeCount = 0;
        foreach (var shape in shapes)
        {
            // Skip shapes that are already hidden
            if (ShapeHelper.IsShapeHidden(shape))
                continue;

            if (_shapeProcessor.ProcessShape(shape, variables))
            {
                processedShapeCount++;
            }
        }

        Logger.Debug($"Processed {processedShapeCount} shapes in slide {slideId}");
        slidePart.Slide.Save();
    }

    /// <summary>
    /// Pre-scan shapes to hide those with invalid array references
    /// </summary>
    private void PreScanAndHideInvalidShapes(SlidePart slidePart, Dictionary<string, object> variables)
    {
        var shapes = slidePart.Slide.Descendants<P.Shape>().ToList();
        int hiddenCount = 0;

        foreach (var shape in shapes)
        {
            if (ShapeHelper.IsShapeHidden(shape))
                continue;

            // Check if shape contains array references
            if (ContainsArrayReferences(shape))
            {
                // Check if the array references are valid
                if (ContainsInvalidArrayReferences(shape, variables))
                {
                    ShapeHelper.HideShape(shape);
                    hiddenCount++;
                    Logger.Debug($"Pre-hiding shape '{shape.GetShapeName()}' due to invalid array references");
                }
            }
        }

        if (hiddenCount > 0)
        {
            Logger.Debug($"Pre-hidden {hiddenCount} shapes with invalid array references");
        }
    }

    /// <summary>
    /// Check if a shape contains array references
    /// </summary>
    private bool ContainsArrayReferences(P.Shape shape)
    {
        if (shape.TextBody == null)
            return false;

        foreach (var text in shape.Descendants<A.Text>())
        {
            if (string.IsNullOrEmpty(text.Text))
                continue;

            // Look for ${...} expressions with array indices
            if (text.Text.Contains("[") && text.Text.Contains("]") && text.Text.Contains("${"))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Check if a shape contains invalid array references
    /// </summary>
    private bool ContainsInvalidArrayReferences(P.Shape shape, Dictionary<string, object> variables)
    {
        if (shape.TextBody == null)
            return false;

        // Find all array references in the shape text
        var arrayPattern = new System.Text.RegularExpressions.Regex(@"\${(\w+)\[(\d+)\]");

        foreach (var text in shape.Descendants<A.Text>())
        {
            if (string.IsNullOrEmpty(text.Text))
                continue;

            var matches = arrayPattern.Matches(text.Text);
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                if (match.Groups.Count < 3)
                    continue;

                string arrayName = match.Groups[1].Value;
                string indexStr = match.Groups[2].Value;

                if (!int.TryParse(indexStr, out int arrayIndex))
                    continue;

                // Check if collection exists
                if (!variables.TryGetValue(arrayName, out var arrayObj) || arrayObj == null)
                    continue;

                // Get collection size
                int collectionSize = DataNavigationHelper.GetCollectionCount(arrayObj);

                // Check if index is valid (directly or through context mapping)
                bool validIndex = false;

                // Check context indices first
                if (_context.HierarchicalIndices.TryGetValue($"{arrayName}[{arrayIndex}]", out int mappedIndex))
                {
                    // Special case: -1 is explicitly marked as invalid
                    if (mappedIndex == -1)
                        return true;

                    // Check if the mapped index is valid
                    validIndex = mappedIndex >= 0 && mappedIndex < collectionSize;
                }
                else
                {
                    // Direct index check
                    validIndex = arrayIndex >= 0 && arrayIndex < collectionSize;
                }

                if (!validIndex)
                {
                    Logger.Debug($"Invalid array reference: {arrayName}[{arrayIndex}], collection size: {collectionSize}");
                    return true;
                }
            }
        }

        return false;
    }

    /// <summary>
    /// Apply hierarchical indices for a slide based on mapping
    /// </summary>
    private void ApplyHierarchicalIndicesForSlide(string slideId)
    {
        // Clear previous indices first
        _context.HierarchicalIndices.Clear();

        // Look for hierarchy maps containing this slide
        bool foundMapping = false;
        foreach (var hierarchyMap in _context.ProcessedHierarchies.Values)
        {
            if (hierarchyMap.SlideMappings.TryGetValue(slideId, out var slideMapping))
            {
                Logger.Debug($"Found hierarchy mapping for slide {slideId}");
                foundMapping = true;

                // Apply indices from mapping
                foreach (var kvp in slideMapping.PathIndices)
                {
                    _context.HierarchicalIndices[kvp.Key] = kvp.Value;
                    Logger.Debug($"Applied index {kvp.Value} for path {kvp.Key}");
                }

                // Handle array indices for the items on this slide
                string rootName = hierarchyMap.Path?.GetRoot()?.Name;
                if (!string.IsNullOrEmpty(rootName))
                {
                    int baseIndex = slideMapping.BaseIndex;
                    int itemsPerSlide = slideMapping.ItemsPerSlide;
                    int totalItems = slideMapping.TotalItems;

                    // For each template array reference (e.g., Items[0], Items[1], etc.)
                    for (int i = 0; i < itemsPerSlide; i++)
                    {
                        string indexPath = $"{rootName}[{i}]";

                        // If this index is applicable to this slide, map it
                        if (i < totalItems)
                        {
                            int actualIndex = baseIndex + i;
                            _context.HierarchicalIndices[indexPath] = actualIndex;
                            Logger.Debug($"Mapped array index {indexPath} to {actualIndex}");

                            // Also track property paths with this index
                            // This ensures expressions like Items[0].Name work correctly
                            _context.HierarchicalIndices[$"{rootName}{i}"] = actualIndex;
                        }
                        else
                        {
                            // For indices beyond the slide's items, map to a safe value
                            // This ensures that out-of-range references don't crash but can be hidden
                            _context.HierarchicalIndices[indexPath] = -1;
                            Logger.Debug($"Mapped out-of-range array index {indexPath} to -1");
                        }
                    }
                }

                // Store additional context information about the slide mapping
                _context.Variables["_slideMapping"] = slideMapping;
                _context.Variables["_baseIndex"] = slideMapping.BaseIndex;
                _context.Variables["_itemsPerSlide"] = slideMapping.ItemsPerSlide;
                _context.Variables["_totalItems"] = slideMapping.TotalItems;
                _context.Variables["_collectionPath"] = hierarchyMap.Path?.ToString() ?? "";

                // No need to continue looking for mappings
                break;
            }
        }

        if (!foundMapping)
        {
            Logger.Debug($"No hierarchy mapping found for slide {slideId}");
        }

        // Debug dump of indices
        if (Logger.MinimumLevel <= Logger.LogLevel.Debug)
        {
            _context.DumpHierarchicalIndices();
        }
    }
}