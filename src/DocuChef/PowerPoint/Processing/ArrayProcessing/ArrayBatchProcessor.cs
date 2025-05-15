using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing.ArrayProcessing;

/// <summary>
/// Processes array-based batch operations for PowerPoint slides
/// </summary>
internal partial class ArrayBatchProcessor
{
    private readonly PowerPointContext _context;
    private readonly Dictionary<string, object> _variables;

    /// <summary>
    /// Creates a new instance of the ArrayBatchProcessor
    /// </summary>
    public ArrayBatchProcessor(PowerPointContext context, Dictionary<string, object> variables)
    {
        _context = context ?? throw new ArgumentNullException(nameof(context));
        _variables = variables ?? throw new ArgumentNullException(nameof(variables));
    }

    /// <summary>
    /// Process array batch with the given parameters
    /// </summary>
    public SlideProcessingResult ProcessArrayBatch(
        PresentationPart presentationPart,
        SlidePart slidePart,
        ArrayBatchParameters parameters)
    {
        var result = new SlideProcessingResult
        {
            SlidePart = slidePart,
            WasProcessed = false
        };

        if (parameters == null || string.IsNullOrEmpty(parameters.CollectionName))
        {
            Logger.Warning("Invalid array batch parameters");
            return result;
        }

        // For nested collections, we need special handling
        if (parameters.IsNestedCollection)
        {
            // Resolve the nested collection first
            ResolveNestedCollection(parameters);
        }

        // After resolving the nested collection (if needed), continue with standard processing
        string collectionName = parameters.CollectionName;
        Logger.Debug($"Processing array batch for collection: {collectionName}");

        // Check if collection exists
        if (!_variables.TryGetValue(collectionName, out var collectionObj) || collectionObj == null)
        {
            Logger.Warning($"Collection not found: {collectionName}");
            return result;
        }

        // Get collection count
        int itemCount = CollectionHelper.GetCollectionCount(collectionObj);
        if (itemCount == 0)
        {
            // Mark as processed but don't create slides for empty collections
            Logger.Warning($"Collection is empty: {collectionName}");
            result.WasProcessed = true;

            // For empty collections, remove the original slide
            RemoveSlide(presentationPart, slidePart);

            return result;
        }

        // Process the batch as usual
        return ProcessBatch(presentationPart, slidePart, parameters, collectionObj, itemCount);
    }

    /// <summary>
    /// Resolves a nested collection based on parameters
    /// </summary>
    private void ResolveNestedCollection(ArrayBatchParameters parameters)
    {
        if (!parameters.IsNestedCollection)
            return;

        // Extract parent and child information
        string parentName = parameters.GetParentCollectionName();
        string childName = parameters.GetChildCollectionName();
        int parentIndex = parameters.GetParentIndex();

        Logger.Debug($"Resolving nested collection: {parentName}[{parentIndex}].{childName}");

        // First check if context has this collection already resolved
        string combinedName = parameters.GetCombinedName();
        if (_variables.ContainsKey(combinedName))
        {
            Logger.Debug($"Found already resolved collection: {combinedName}");
            return;
        }

        // Get parent collection
        if (!_variables.TryGetValue(parentName, out var parentObj) || parentObj == null)
        {
            Logger.Warning($"Parent collection not found: {parentName}");
            return;
        }

        // Get parent item at index
        var parentItem = CollectionHelper.GetItemAtIndex(parentObj, parentIndex);
        if (parentItem == null)
        {
            Logger.Warning($"Parent item not found at index {parentIndex}");
            return;
        }

        // Get child collection from parent item
        var childProp = parentItem.GetType().GetProperty(childName);
        if (childProp == null)
        {
            Logger.Warning($"Child property not found: {childName}");
            return;
        }

        var childCollection = childProp.GetValue(parentItem);
        if (childCollection == null)
        {
            Logger.Warning($"Child collection is null");
            return;
        }

        // Store resolved collection in variables
        _variables[combinedName] = childCollection;
        Logger.Debug($"Resolved nested collection: {combinedName} with {CollectionHelper.GetCollectionCount(childCollection)} items");

        // Also set current index in context if available
        if (_context != null)
        {
            _context.SetCollectionIndex(parentName, parentIndex);
        }
    }

    /// <summary>
    /// Process a batch with resolved collection
    /// </summary>
    private SlideProcessingResult ProcessBatch(
        PresentationPart presentationPart,
        SlidePart slidePart,
        ArrayBatchParameters parameters,
        object collectionObj,
        int itemCount)
    {
        var result = new SlideProcessingResult
        {
            SlidePart = slidePart,
            WasProcessed = false
        };

        string collectionName = parameters.CollectionName;

        // Determine max items per slide
        int maxItemsPerSlide = parameters.MaxItemsPerSlide;

        // If not explicitly specified, auto-detect from slide
        if (maxItemsPerSlide <= 0)
        {
            maxItemsPerSlide = DetectMaxItemsPerSlide(slidePart, collectionName);
            Logger.Debug($"Auto-detected max items per slide: {maxItemsPerSlide}");
        }

        // Ensure at least 1 item per slide
        maxItemsPerSlide = Math.Max(1, maxItemsPerSlide);

        // Get offset
        int offset = Math.Max(0, parameters.Offset);

        // Adjust effective item count based on offset
        int effectiveCount = Math.Max(0, itemCount - offset);
        if (effectiveCount <= 0)
        {
            Logger.Warning($"No items available after offset {offset} for collection with {itemCount} items");
            result.WasProcessed = true;

            // Remove the slide since there are no items to display
            RemoveSlide(presentationPart, slidePart);

            return result;
        }

        // Calculate how many slides needed
        int slidesNeeded = (int)Math.Ceiling((double)effectiveCount / maxItemsPerSlide);
        Logger.Debug($"Collection {collectionName} has {itemCount} items (effective: {effectiveCount} after offset {offset}), " +
                     $"max {maxItemsPerSlide} per slide, need {slidesNeeded} slides");

        // Mark the original slide as processed
        result.WasProcessed = true;

        // Get the position of this slide
        int basePosition = SlideHelper.FindSlidePosition(presentationPart, slidePart);
        int currentPosition = basePosition;

        // Generate appropriate ID prefix based on collection type
        string idPrefix = parameters.IsNestedCollection
            ? $"{parameters.GetParentCollectionName()}_{parameters.GetParentIndex()}_{parameters.GetChildCollectionName()}"
            : collectionName;

        // Process all batches
        for (int batchIndex = 0; batchIndex < slidesNeeded; batchIndex++)
        {
            int batchStartIndex = offset + (batchIndex * maxItemsPerSlide);

            // Make sure batch index is within bounds
            if (batchStartIndex >= itemCount)
                break;

            // Create or clone slide
            SlidePart batchSlidePart;
            if (batchIndex == 0)
            {
                // Use original slide for first batch
                batchSlidePart = slidePart;
                Logger.Debug($"Using original slide for {collectionName} batch 0 (items {batchStartIndex}-{Math.Min(batchStartIndex + maxItemsPerSlide - 1, itemCount - 1)})");
            }
            else
            {
                // Clone the original slide for subsequent batches
                batchSlidePart = SlideHelper.CloneSlide(presentationPart, slidePart);
                SlideHelper.InsertSlide(presentationPart, batchSlidePart, ++currentPosition);
                Logger.Debug($"Created slide at position {currentPosition} for {collectionName} batch {batchIndex} (items {batchStartIndex}-{Math.Min(batchStartIndex + maxItemsPerSlide - 1, itemCount - 1)})");
            }

            // Update array references with batch offset
            UpdateArrayReferences(batchSlidePart, parameters.IsNestedCollection ? parameters.GetChildCollectionName() : collectionName, batchStartIndex);

            // Hide out-of-range items
            HideOutOfRangeItems(batchSlidePart, parameters.IsNestedCollection ? parameters.GetChildCollectionName() : collectionName, batchStartIndex, itemCount, maxItemsPerSlide);

            // Mark as processed
            string processedId = $"{idPrefix}_batch_{batchIndex}";
            if (_context != null)
            {
                _context.ProcessedArraySlides.Add(processedId);
            }
            result.GeneratedSlides.Add(batchSlidePart);
        }

        // Save the presentation to apply changes
        presentationPart.Presentation.Save();

        return result;
    }

    /// <summary>
    /// Auto-detect maximum items per slide based on array references
    /// </summary>
    private int DetectMaxItemsPerSlide(SlidePart slidePart, string arrayName)
    {
        var arrayReferences = FindArrayReferencesInSlide(slidePart)
            .Where(r => r.ArrayName == arrayName)
            .ToList();

        if (!arrayReferences.Any())
            return 1; // Default to 1 if no references found

        // Find the maximum index referenced
        int maxIndex = arrayReferences.Max(r => r.Index);

        // Max items per slide is max index + 1
        return maxIndex + 1;
    }

    /// <summary>
    /// Find all array references in slide
    /// </summary>
    private List<ArrayReference> FindArrayReferencesInSlide(SlidePart slidePart)
    {
        var result = new List<ArrayReference>();

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            var references = PowerPointShapeHelper.FindArrayReferences(shape);
            result.AddRange(references);
        }

        return result;
    }

    /// <summary>
    /// Update array references in slide with specified offset
    /// </summary>
    private void UpdateArrayReferences(SlidePart slidePart, string arrayName, int offset)
    {
        Logger.Debug($"Updating array references: {arrayName} with offset {offset}");

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null)
                continue;

            foreach (var text in shape.Descendants<A.Text>())
            {
                if (string.IsNullOrEmpty(text.Text) || !text.Text.Contains(arrayName))
                    continue;

                string original = text.Text;

                // 1. Process ${ArrayName[n]} pattern
                var dollarPattern = new Regex($@"\$\{{{arrayName}\[(\d+)\]([^\}}]*)\}}");
                text.Text = dollarPattern.Replace(text.Text, match => {
                    int index = int.Parse(match.Groups[1].Value);
                    string suffix = match.Groups[2].Value;
                    return $"${{{arrayName}[{index + offset}]{suffix}}}";
                });

                // 2. Process direct ArrayName[n] pattern
                var directPattern = new Regex($@"(?<!\$\{{){arrayName}\[(\d+)\]");
                text.Text = directPattern.Replace(text.Text, match => {
                    int index = int.Parse(match.Groups[1].Value);
                    return $"{arrayName}[{index + offset}]";
                });

                if (text.Text != original)
                {
                    Logger.Debug($"Updated text from '{original}' to '{text.Text}'");
                }
            }
        }
    }

    /// <summary>
    /// Hide shapes with out-of-range array indices
    /// </summary>
    private void HideOutOfRangeItems(SlidePart slidePart, string arrayName, int startIndex, int totalItems, int maxItems)
    {
        Logger.Debug($"Hiding out-of-range shapes: {arrayName}, startIndex={startIndex}, totalItems={totalItems}, maxItems={maxItems}");

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (PowerPointShapeHelper.IsShapeHidden(shape))
                continue;

            bool hasOutOfRangeReference = false;

            // 1. Check text content
            if (shape.TextBody != null)
            {
                foreach (var text in shape.Descendants<A.Text>())
                {
                    if (string.IsNullOrEmpty(text.Text))
                        continue;

                    // Check for array references
                    var matches = Regex.Matches(text.Text, $@"\$\{{{arrayName}\[(\d+)\]");
                    foreach (Match match in matches)
                    {
                        if (match.Groups.Count < 2)
                            continue;

                        int referencedIndex = int.Parse(match.Groups[1].Value);

                        // Check if referenced index is out of range or beyond this batch's range
                        if (referencedIndex >= totalItems)
                        {
                            hasOutOfRangeReference = true;
                            Logger.Debug($"Found out-of-range reference: {arrayName}[{referencedIndex}] >= {totalItems}");
                            break;
                        }

                        // Check batch range
                        int batchEndIndex = startIndex + maxItems - 1;
                        if (referencedIndex < startIndex || referencedIndex > batchEndIndex)
                        {
                            hasOutOfRangeReference = true;
                            Logger.Debug($"Found out-of-batch reference: {arrayName}[{referencedIndex}] outside range {startIndex}-{batchEndIndex}");
                            break;
                        }
                    }

                    if (hasOutOfRangeReference)
                        break;
                }
            }

            // 2. Check shape name for array references
            if (!hasOutOfRangeReference)
            {
                string shapeName = shape.GetShapeName();
                if (!string.IsNullOrEmpty(shapeName))
                {
                    // Check for direct array references in shape name
                    var directMatches = Regex.Matches(shapeName, $@"{arrayName}\[(\d+)\]");
                    foreach (Match match in directMatches)
                    {
                        if (match.Groups.Count < 2)
                            continue;

                        int referencedIndex = int.Parse(match.Groups[1].Value);

                        // Check range
                        if (referencedIndex >= totalItems ||
                            referencedIndex < startIndex ||
                            referencedIndex > startIndex + maxItems - 1)
                        {
                            hasOutOfRangeReference = true;
                            Logger.Debug($"Found out-of-range reference in shape name: {shapeName}");
                            break;
                        }
                    }
                }
            }

            // 3. Check shape alt text (Description) for array references
            if (!hasOutOfRangeReference)
            {
                var altText = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Description?.Value;
                if (!string.IsNullOrEmpty(altText) && altText.Contains(arrayName))
                {
                    // Check for array references in alt text
                    var matches = Regex.Matches(altText, $@"{arrayName}\[(\d+)\]");
                    foreach (Match match in matches)
                    {
                        if (match.Groups.Count < 2)
                            continue;

                        int referencedIndex = int.Parse(match.Groups[1].Value);

                        // Check range
                        if (referencedIndex >= totalItems ||
                            referencedIndex < startIndex ||
                            referencedIndex > startIndex + maxItems - 1)
                        {
                            hasOutOfRangeReference = true;
                            Logger.Debug($"Found out-of-range reference in alt text: {arrayName}[{referencedIndex}]");
                            break;
                        }
                    }
                }
            }

            // Hide the shape if out-of-range references were found
            if (hasOutOfRangeReference)
            {
                PowerPointShapeHelper.HideShape(shape);
                Logger.Debug($"Hidden shape with out-of-range reference to {arrayName}: {shape.GetShapeName()}");
            }
        }
    }

    /// <summary>
    /// Remove a slide from presentation
    /// </summary>
    private void RemoveSlide(PresentationPart presentationPart, SlidePart slidePart)
    {
        var slideIds = presentationPart.Presentation.SlideIdList;
        string relationshipId = presentationPart.GetIdOfPart(slidePart);
        var slideIdToRemove = slideIds.Elements<SlideId>()
            .FirstOrDefault(s => s.RelationshipId == relationshipId);

        if (slideIdToRemove != null)
        {
            slideIds.RemoveChild(slideIdToRemove);
            Logger.Debug($"Removed slide from presentation");
        }
    }
}