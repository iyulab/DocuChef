using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing.ArrayProcessing;

/// <summary>
/// Extension to ArrayBatchProcessor for handling nested collections
/// </summary>
internal partial class ArrayBatchProcessor
{
    /// <summary>
    /// Override of the standard ProcessArrayBatch method to handle nested collections
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
}