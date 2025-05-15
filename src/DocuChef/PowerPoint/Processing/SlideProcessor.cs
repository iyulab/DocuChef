using DocuChef.PowerPoint.Helpers;

namespace DocuChef.PowerPoint.Processing;

/// <summary>
/// Processor responsible for slide analysis and preparation
/// </summary>
internal partial class SlideProcessor
{
    private readonly PowerPointProcessor _mainProcessor;
    private readonly PowerPointContext _context;

    /// <summary>
    /// Initialize slide processor
    /// </summary>
    public SlideProcessor(PowerPointProcessor mainProcessor, PowerPointContext context)
    {
        _mainProcessor = mainProcessor ?? throw new ArgumentNullException(nameof(mainProcessor));
        _context = context ?? throw new ArgumentNullException(nameof(context));
    }

    /// <summary>
    /// Analyze slide and prepare duplicates if needed
    /// </summary>
    public void AnalyzeAndPrepareSlide(PresentationPart presentationPart, SlidePart slidePart)
    {
        string slideId = presentationPart.GetIdOfPart(slidePart);
        Logger.Debug($"Analyzing slide {slideId} for array references");

        // 1. Analyze array references in this slide
        var arrayReferences = FindArrayReferencesInSlide(slidePart);
        if (!arrayReferences.Any())
        {
            Logger.Debug("No array references found in slide");
            return;
        }

        // 2. Process by array
        foreach (var arrayGroup in arrayReferences.GroupBy(r => r.ArrayName))
        {
            string arrayName = arrayGroup.Key;
            int maxIndex = arrayGroup.Max(r => r.Index);
            int itemsPerSlide = maxIndex + 1;

            Logger.Debug($"Found array '{arrayName}' with max index {maxIndex} in slide");

            // Check data array size
            if (!_context.Variables.TryGetValue(arrayName, out var arrayObj) || arrayObj == null)
            {
                Logger.Warning($"Array '{arrayName}' not found in variables");
                continue;
            }

            int totalItems = CollectionHelper.GetCollectionCount(arrayObj);
            Logger.Debug($"Array '{arrayName}' has {totalItems} total items, {itemsPerSlide} items per slide");

            // Important: Always hide out-of-range items on the original slide
            // This also applies when Items.Count is less than designed elements
            HideOutOfRangeItems(slidePart, arrayName, 0, totalItems);
            Logger.Debug($"Applied range check to original slide for array '{arrayName}'");

            // Determine if duplication is needed
            if (totalItems <= itemsPerSlide)
            {
                Logger.Debug($"No duplication needed for array '{arrayName}'");
                continue;  // No duplication needed
            }

            // 3. Calculate number of slides needed (ceiling division)
            int slidesNeeded = (int)Math.Ceiling((double)totalItems / itemsPerSlide);
            Logger.Info($"Array '{arrayName}' requires {slidesNeeded} slides for {totalItems} items ({itemsPerSlide} items per slide)");

            // 4. Duplicate additional slides (starting from the second slide, as the first one already exists)
            int baseSlidePosition = SlideHelper.FindSlidePosition(presentationPart, slidePart);

            for (int i = 1; i < slidesNeeded; i++)
            {
                // Calculate batch start index
                int batchStartIndex = i * itemsPerSlide;
                Logger.Debug($"Creating slide {i + 1} for batch starting at index {batchStartIndex}");

                // Clone slide
                var newSlidePart = SlideHelper.CloneSlide(presentationPart, slidePart);

                // Update array indices in the cloned slide
                UpdateArrayIndices(newSlidePart, arrayName, batchStartIndex);

                // Insert slide
                SlideHelper.InsertSlide(presentationPart, newSlidePart, baseSlidePosition + i);
                Logger.Debug($"Inserted duplicated slide at position {baseSlidePosition + i}");

                // Hide out-of-range items
                HideOutOfRangeItems(newSlidePart, arrayName, batchStartIndex, totalItems);

                // Mark as processed slide
                _context.ProcessedArraySlides.Add(presentationPart.GetIdOfPart(newSlidePart));
            }
        }

        // Save changes
        presentationPart.Presentation.Save();
    }

    /// <summary>
    /// Find all array references in slide
    /// </summary>
    private List<ArrayReference> FindArrayReferencesInSlide(SlidePart slidePart)
    {
        var result = new List<ArrayReference>();

        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (shape.TextBody == null)
                continue;

            var references = PowerPointShapeHelper.FindArrayReferences(shape);
            result.AddRange(references);
        }

        return result;
    }

    /// <summary>
    /// Update array indices in slide with specified offset
    /// </summary>
    private void UpdateArrayIndices(SlidePart slidePart, string arrayName, int offset)
    {
        Logger.Debug($"Updating array indices for '{arrayName}' with offset {offset}");

        int updatedShapeCount = 0;
        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            bool shapeUpdated = false;

            // Process all text elements
            var texts = shape.Descendants<A.Text>().ToList();
            foreach (var text in texts)
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

                // 2. Process direct ArrayName[n] pattern (function arguments, etc.)
                var directPattern = new Regex($@"(?<!\$\{{){arrayName}\[(\d+)\]");
                text.Text = directPattern.Replace(text.Text, match => {
                    int index = int.Parse(match.Groups[1].Value);
                    return $"{arrayName}[{index + offset}]";
                });

                if (text.Text != original)
                {
                    Logger.Debug($"Updated text from '{original}' to '{text.Text}'");
                    shapeUpdated = true;
                }
            }

            if (shapeUpdated)
                updatedShapeCount++;
        }

        Logger.Debug($"Updated array indices in {updatedShapeCount} shapes");
        slidePart.Slide.Save();
    }

    /// <summary>
    /// Hide shapes with out-of-range array indices, ensuring all items beyond total count are hidden
    /// </summary>
    private void HideOutOfRangeItems(SlidePart slidePart, string arrayName, int startIndex, int totalItems)
    {
        Logger.Debug($"Checking for out-of-range items in array '{arrayName}': startIndex={startIndex}, totalItems={totalItems}");

        int hiddenShapeCount = 0;
        foreach (var shape in slidePart.Slide.Descendants<P.Shape>())
        {
            if (PowerPointShapeHelper.IsShapeHidden(shape))
                continue;

            // Check all array references in this shape
            var references = PowerPointShapeHelper.FindArrayReferences(shape)
                            .Where(r => r.ArrayName == arrayName)
                            .ToList();

            if (!references.Any())
                continue;

            // Hide if there are out-of-range references
            foreach (var reference in references)
            {
                // Calculate actual index (for cloned slides with offset applied indices)
                int actualIndex = reference.Index;

                // Hide if index is beyond total items
                if (actualIndex >= totalItems)
                {
                    Logger.Debug($"Hiding shape '{shape.GetShapeName()}' with reference to {arrayName}[{actualIndex}] (>= {totalItems})");
                    PowerPointShapeHelper.HideShape(shape);
                    hiddenShapeCount++;
                    break;
                }

                // Additionally, hide if out of current slide's batch range
                // Only show indices in range: startIndex ~ startIndex + (slidesPerBatch - 1)
                int slidesPerBatch = GetSlidesPerBatch(slidePart, arrayName);
                int batchEndIndex = startIndex + slidesPerBatch - 1;

                // Calculate local index (local index for items displayed on cloned slide)
                int localIndex = actualIndex - startIndex;

                if (localIndex < 0 || localIndex >= slidesPerBatch)
                {
                    Logger.Debug($"Hiding shape '{shape.GetShapeName()}' with out-of-batch index: {arrayName}[{actualIndex}] (local index {localIndex} outside of batch 0-{slidesPerBatch - 1})");
                    PowerPointShapeHelper.HideShape(shape);
                    hiddenShapeCount++;
                    break;
                }
            }
        }

        Logger.Debug($"Hidden {hiddenShapeCount} shapes with out-of-range references");
        slidePart.Slide.Save();
    }

    /// <summary>
    /// Get number of items per batch for the given slide and array
    /// </summary>
    private int GetSlidesPerBatch(SlidePart slidePart, string arrayName)
    {
        // Find all index references for the array in this slide and return max index + 1
        var references = FindArrayReferencesInSlide(slidePart)
                        .Where(r => r.ArrayName == arrayName)
                        .ToList();

        if (!references.Any())
            return 0;

        return references.Max(r => r.Index) + 1;
    }
}